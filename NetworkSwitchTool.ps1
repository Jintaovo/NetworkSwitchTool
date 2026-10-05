<#
=====================================================================
 网络连接查看与切换工具  NetworkSwitchTool.ps1
---------------------------------------------------------------------
 切换方式：调整「接口跃点数（Interface Metric）」来改变实际流量出口，
           两块网卡都保持启用状态，不断开任何网卡，可随时来回切换。

 功能：
   1) 界面实时显示当前网络连接信息（流量出口、IP、掩码、网关、DNS、
      MAC、链路速率、跃点数、DHCP、无线 SSID / 信号 / 信道 / 速率等）。
   2) 按钮「本地网络切换」：把流量出口在本地网络 A / B 之间来回切换。
   3) 按钮「切换到 WiFi」：连接 WiFi 并把流量出口切到无线网卡。
   4) 设置窗口：选择网卡、WiFi、跃点数值，并可「恢复自动跃点」。

 依赖：Windows PowerShell 5.1 + .NET WinForms（系统自带，无需安装）。
 权限：调整跃点数 / 连接 WiFi 需要管理员权限，脚本会自行请求提升。
 数据源：.NET NetworkInterface + netsh + route（不依赖 WMI，兼容性更好）。
=====================================================================
#>
[CmdletBinding()]
param(
    [switch]$Probe,      # 只输出采集到的信息（控制台文本），不启动界面
    [switch]$NoElevate   # 不在启动时请求管理员提升
)

$ErrorActionPreference = 'Continue'
Set-StrictMode -Off

# ---------------------------------------------------------------------
# 路径与全局状态
# ---------------------------------------------------------------------
if ($PSScriptRoot) { $script:ScriptDir = $PSScriptRoot }
else { $script:ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition }
if (-not $script:ScriptDir) { $script:ScriptDir = (Get-Location).Path }

$script:SelfPath = $PSCommandPath
if (-not $script:SelfPath) { $script:SelfPath = Join-Path $script:ScriptDir 'NetworkSwitchTool.ps1' }

$script:ConfigPath  = Join-Path $script:ScriptDir 'config.json'
$script:LogPath     = Join-Path $script:ScriptDir 'switch-log.txt'
$script:ConfigCache = $null
$script:Busy        = $false
$script:LastError   = ''
$script:MainForm    = $null
$script:UiFont      = $null
$script:LegacyEnc   = $null
$script:AdminMapCache     = $null
$script:AdminMapCacheTime = [datetime]::MinValue
$script:AdminMapSig       = ''
$script:AdminMapTtlSec    = 600    # netsh interface show interface 很慢（实测 ~4.5s），必须缓存
$script:LastRefreshMs     = 0
$script:LastSnapshot      = $null  # Update-Ui 采集的最新快照，供自动回切看门狗复用
$script:LastFallback      = [datetime]::MinValue
$script:ButtonPanel       = $null  # 按钮面板（忙碌时统一禁用）

$script:DefaultConfig = [ordered]@{
    localA               = ''      # 本地网络 A 的网卡名称（netsh 中的接口名）
    localB               = ''      # 本地网络 B 的网卡名称
    wifiAdapter          = ''      # 无线网卡名称
    wifiSsid             = ''      # 要连接的无线网络名称
    disableLocalWhenWifi = $false  # 切到 WiFi 后是否"禁用"本地网卡（默认只调跃点数）
    preferMetric         = 10      # 首选出口网卡的接口跃点数（越小越优先）
    deprioritizeMetric   = 1000    # 其它网卡的接口跃点数
    lastTarget           = ''      # 上次切换的目标网卡名，用于按钮严格 A↔B 交替
    lastWired            = ''      # 切到 WiFi 之前的那块有线网卡，用于「切回有线」
    autoFallback         = $true   # WiFi 断开/不可用时自动把流量切回有线
    wifiAltFallback      = $true   # 指定的 WiFi 连不上时，自动尝试其它"可见且已保存"的网络
    wifiConnectSeconds   = 15      # 连接单个 WiFi 网络最多等多少秒（越短越快失败）
    autoRefreshSeconds   = 3       # 界面自动刷新间隔，0 = 关闭
    switchTimeoutSeconds = 30      # 等待网卡启用 / 无线连接的超时时间
}

# ---------------------------------------------------------------------
# 基础工具
# ---------------------------------------------------------------------
function Write-Log {
    param([string]$Message)
    $line = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    try { Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8 } catch { }
}

function Get-Config {
    if ($script:ConfigCache) { return $script:ConfigCache }
    $cfg = [ordered]@{}
    foreach ($k in $script:DefaultConfig.Keys) { $cfg[$k] = $script:DefaultConfig[$k] }
    if (Test-Path -LiteralPath $script:ConfigPath) {
        try {
            $raw = Get-Content -LiteralPath $script:ConfigPath -Raw -Encoding UTF8
            if ($raw -and $raw.Trim() -ne '') {
                $obj = $raw | ConvertFrom-Json
                foreach ($k in $script:DefaultConfig.Keys) {
                    if ($obj.PSObject.Properties.Name -contains $k -and $null -ne $obj.$k) { $cfg[$k] = $obj.$k }
                }
            }
        } catch { Write-Log "读取 config.json 失败：$($_.Exception.Message)" }
    }
    $script:ConfigCache = $cfg
    return $cfg
}

function Save-Config {
    param($Config)
    try {
        ($Config | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath $script:ConfigPath -Encoding UTF8
        $script:ConfigCache = $null
        Write-Log '配置已保存'
        return $true
    } catch {
        Write-Log "保存 config.json 失败：$($_.Exception.Message)"
        return $false
    }
}

function Quote-NetshArg {
    param([string]$Value)
    if ($null -eq $Value) { return '""' }
    return '"' + ($Value -replace '"', '\"') + '"'
}

function Get-LegacyEncoding {
    # 非 UTF-8 的原生输出用系统 OEM 代码页（中文系统 = 936）。
    # 注意不能用 [Console]::OutputEncoding 兜底：某些进程里它本身就是 UTF-8，
    # 那样 GBK 字节会被再解错一次。
    if ($script:LegacyEnc) { return $script:LegacyEnc }
    $enc = $null
    try { $enc = [System.Text.Encoding]::GetEncoding([int][System.Globalization.CultureInfo]::CurrentCulture.TextInfo.OEMCodePage) } catch { $enc = $null }
    if (-not $enc) { try { $enc = [System.Text.Encoding]::Default } catch { $enc = [System.Text.Encoding]::ASCII } }
    $script:LegacyEnc = $enc
    return $enc
}

function Convert-NativeBytes {
    <#
      ★ 关键修复 ★
      控制台程序的输出编码会随进程而变：实测同一台机器上，普通实例里 netsh 输出 GBK，
      管理员（隐藏控制台）实例里 netsh 输出 UTF-8。若按固定编码解码，界面上看起来只是
      几个字乱码，但"已启用/已禁用"这种状态词也会一起烂掉，导致：
        明明已启用的网卡被判成"已禁用" -> 白白重置一次网卡 + 死等 20 秒超时。
      所以这里读原始字节再判断：先按严格 UTF-8 试，失败才用系统代码页。
    #>
    param([byte[]]$Bytes)
    if (-not $Bytes -or $Bytes.Length -eq 0) { return '' }
    $hasHigh = $false
    foreach ($b in $Bytes) { if ($b -gt 127) { $hasHigh = $true; break } }
    if (-not $hasHigh) { return [System.Text.Encoding]::ASCII.GetString($Bytes) }
    if ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF) {
        return (New-Object System.Text.UTF8Encoding($false)).GetString($Bytes, 3, $Bytes.Length - 3)
    }
    try {
        $strict = New-Object System.Text.UTF8Encoding($false, $true)
        return $strict.GetString($Bytes)
    } catch { }
    return (Get-LegacyEncoding).GetString($Bytes)
}

function Invoke-NativeProcess {
    <#
      用 ProcessStartInfo 直接传命令行字符串，避免 PowerShell 5.1 对参数中
      引号的转义破坏 netsh 语法（例如 name="My Adapter"）。
      输出按原始字节读回后智能解码（见 Convert-NativeBytes）。
      加固：stdout/stderr 用异步读（避免管道写满互锁）、带超时（超时 Kill）、
            try/finally 保证进程与流一定被释放。
    #>
    param([Parameter(Mandatory)][string]$FilePath, [string]$Arguments = '', [int]$TimeoutSec = 20)
    $out = ''
    $code = -1
    $proc = $null
    $ms = $null
    $ems = $null
    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $FilePath
        $psi.Arguments = $Arguments
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $proc = [System.Diagnostics.Process]::Start($psi)
        $ms = New-Object System.IO.MemoryStream
        $ems = New-Object System.IO.MemoryStream
        # 异步读两个流，避免"先读 stdout 到 EOF、子进程却在写 stderr"造成的死锁
        $tOut = $proc.StandardOutput.BaseStream.CopyToAsync($ms)
        $tErr = $proc.StandardError.BaseStream.CopyToAsync($ems)
        $exited = $proc.WaitForExit($TimeoutSec * 1000)
        if (-not $exited) {
            try { $proc.Kill() } catch { }
            try { $null = $proc.WaitForExit(3000) } catch { }
        }
        try { $null = [System.Threading.Tasks.Task]::WaitAll(@($tOut, $tErr), 3000) } catch { }
        if ($exited) { $code = $proc.ExitCode } else { $code = -2 }
        $out = Convert-NativeBytes $ms.ToArray()
        $errText = Convert-NativeBytes $ems.ToArray()
        if ($errText) { $out = $out + $errText }
        if (-not $exited) { $out = $out + ('[命令超时 {0} 秒，已强制终止]' -f $TimeoutSec) }
    } catch {
        $out = $_.Exception.Message
        $code = -1
    } finally {
        if ($ems) { try { $ems.Dispose() } catch { } }
        if ($ms) { try { $ms.Dispose() } catch { } }
        if ($proc) { try { $proc.Dispose() } catch { } }
    }
    return New-Object psobject -Property @{ Output = $out; ExitCode = $code }
}

function Invoke-Netsh {
    param([Parameter(Mandatory)][string]$CommandLine)
    return Invoke-NativeProcess -FilePath (Join-Path $env:SystemRoot 'System32\netsh.exe') -Arguments $CommandLine
}

function Invoke-Route {
    param([string]$CommandLine = 'print -4')
    return Invoke-NativeProcess -FilePath (Join-Path $env:SystemRoot 'System32\route.exe') -Arguments $CommandLine
}

function Get-NetshPairs {
    param([string]$Text)
    $h = @{}
    foreach ($line in ($Text -split "`r?`n")) {
        if ($line -match '^\s*([^:：]+?)\s*[:：]\s*(.*?)\s*$') {
            $k = $Matches[1].Trim()
            $v = $Matches[2].Trim()
            if ($k -ne '' -and -not $h.ContainsKey($k)) { $h[$k] = $v }
        }
    }
    return $h
}

function Get-PairValue {
    param($Pairs, [string[]]$Keys)
    foreach ($k in $Keys) {
        if ($Pairs.ContainsKey($k)) { return $Pairs[$k] }
    }
    return ''
}

function Test-IsAdmin {
    try {
        $id = [Security.Principal.WindowsIdentity]::GetCurrent()
        return (New-Object Security.Principal.WindowsPrincipal($id)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch { return $false }
}

# ---------------------------------------------------------------------
# 网卡信息采集
# ---------------------------------------------------------------------
function Get-AllInterfaces {
    try { return [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() }
    catch { return @() }
}

function Test-VirtualAdapter {
    param([string]$Name, [string]$Description)
    $pattern = 'Virtual|VirtualBox|VMware|Hyper-V|Loopback|TAP-|Npcap|WAN Miniport|Wi-Fi Direct|Wireless Direct|Bluetooth|Pseudo|Teredo|VPN|KM-TEST|Microsoft Hosted|虚拟'
    return (($Description -match $pattern) -or ($Name -match 'Virtual|Pseudo|Loopback|虚拟'))
}

function Get-InterfaceAdminStateMap {
    <#
      netsh interface show interface 会列出「已禁用的网卡」，而 .NET 的
      NetworkInterface 枚举看不到被禁用的网卡，所以这里作为网卡清单的补充来源。
      ★ 本机实测这条命令要 ~4.5 秒（netsh interface ipv4 show interfaces 只要 60ms），
        所以必须缓存：默认 60 秒内复用，用户操作时可 -Force 强制刷新。
      返回 @{ 网卡名 = @{ AdminState=..; State=..; Type=.. } }
    #>
    param([int]$MaxAgeSec = -1, [switch]$Force)
    if ($MaxAgeSec -lt 0) { $MaxAgeSec = $script:AdminMapTtlSec }
    $now = Get-Date
    # 廉价指纹：.NET 能看到的网卡名单。启用/禁用网卡一定会让这个名单变化，
    # 名单没变就没必要重跑那条 4.5 秒的慢命令。
    $sig = ''
    try { $sig = (@(Get-AllInterfaces | ForEach-Object { $_.Name }) | Sort-Object) -join '|' } catch { $sig = '' }
    if (-not $Force -and $script:AdminMapCache -and $sig -eq $script:AdminMapSig -and ((($now - $script:AdminMapCacheTime).TotalSeconds) -lt $MaxAgeSec)) {
        return $script:AdminMapCache
    }

    $map = @{}
    # .NET 只看得到"已启用"的网卡：被禁用的网卡会从 .NET 里消失，
    # 所以"在 CIM 里但不在 .NET 里"就是被禁用的 —— 这条判断与编码无关，最可靠。
    $netNames = @{}
    try { foreach ($ni in (Get-AllInterfaces)) { $netNames[[string]$ni.Name] = $true } } catch { }

    # ---- 1) CIM：Get-NetAdapter 返回的网卡名是 UTF-16，完全不受控制台编码影响 ----
    if (Get-Command Get-NetAdapter -ErrorAction SilentlyContinue) {
        try {
            foreach ($a in @(Get-NetAdapter -ErrorAction Stop)) {
                $nm = [string]$a.Name
                if (-not $nm) { continue }
                $st = [string]$a.Status
                # 已禁用 = CIM 说 Disabled，或者"CIM 列出了但 .NET 看不到且 CIM 也没说它在 Up"
                # （.NET 看不到被禁用的网卡；加后半句是为了避免把极少数 .NET 不枚举但已启用的网卡误启一遍）
                $disabled = ($st -match 'Disabled|已禁用') -or ((-not $netNames.ContainsKey($nm)) -and ($st -notmatch '^(Up|Connected|已连接)$'))
                $map[$nm] = @{ AdminState = $(if ($disabled) { 'Disabled' } else { 'Enabled' }); State = $st; Type = 'CIM' }
            }
        } catch { }
    }

    # ---- 2) 兜底：只有 CIM 不可用才跑 netsh（这条命令本机实测要 4.5 秒）----
    if ($map.Count -eq 0) {
        $r = Invoke-Netsh 'interface show interface'
        $started = $false
        foreach ($line in ($r.Output -split "`r?`n")) {
            if ($line -match '^-{3,}') { $started = $true; continue }
            if ($line.Trim() -eq '') { continue }
            if ($line -match '^\s*(\S+)\s+(\S+)\s+(\S+)\s+(.+?)\s*$') {
                # 先取出各组，再做其它 -match（-match 会覆盖 $Matches）
                $t1 = $Matches[1]; $t2 = $Matches[2]; $t3 = $Matches[3]; $ifName = $Matches[4]
                $known = ($t1 -match '^(Enabled|Disabled|已启用|已禁用)$')
                if (-not $started -and -not $known) { continue }
                $map[$ifName] = @{ AdminState = $t1; State = $t2; Type = $t3 }
            }
        }
    }

    $script:AdminMapCache = $map
    $script:AdminMapCacheTime = $now
    $script:AdminMapSig = $sig
    return $map
}

function Clear-AdminStateCache {
    $script:AdminMapCache = $null
    $script:AdminMapCacheTime = [datetime]::MinValue
    $script:AdminMapSig = ''
}

function Get-InterfaceMetricMap {
    <#
      接口跃点数来源优先级：
        1) Get-NetIPInterface（CIM，权威，并且带"是否自动度量"标志）
        2) netsh interface ipv4 show interfaces（兜底）
    #>
    $map = @{}
    if (Get-Command Get-NetIPInterface -ErrorAction SilentlyContinue) {
        try {
            foreach ($o in @(Get-NetIPInterface -AddressFamily IPv4 -ErrorAction Stop)) {
                $alias = [string]$o.InterfaceAlias
                # 受限环境里 CIM 可能返回空对象，必须过滤，避免把 netsh 兜底也跳过了
                if (-not $alias) { continue }
                if ($null -eq $o.InterfaceMetric) { continue }
                $auto = $null
                try { $auto = ([string]$o.AutomaticMetric -match '^(Enabled|已启用)') } catch { }
                $map[$alias] = @{
                    Index = [int]$o.InterfaceIndex; Metric = [int]$o.InterfaceMetric
                    Mtu = 0; State = [string]$o.ConnectionState; Auto = $auto
                }
            }
        } catch { }
    }
    if ($map.Count -gt 0) { return $map }

    $r = Invoke-Netsh 'interface ipv4 show interfaces'
    $started = $false
    foreach ($line in ($r.Output -split "`r?`n")) {
        if ($line -match '^-{3,}') { $started = $true; continue }
        if (-not $started) { continue }
        if ($line.Trim() -eq '') { continue }
        if ($line -match '^\s*(\d+)\s+(\d+)\s+(\d+)\s+(\S+)\s+(.+?)\s*$') {
            # MTU 可能是 4294967295（Loopback），必须用 long
            $idx = [int]$Matches[1]; $met = [int]$Matches[2]; $mtu = [long]$Matches[3]; $st = $Matches[4]; $nm = $Matches[5]
            $map[$nm] = @{ Index = $idx; Metric = $met; Mtu = $mtu; State = $st; Auto = $null }
        }
    }
    return $map
}

function Get-InterfaceMetricOf {
    # 取单个网卡当前的 IPv4 接口跃点数（写入后回读校验用）
    param([string]$Name)
    $m = Get-InterfaceMetricMap
    if ($m.ContainsKey($Name)) { return [int]$m[$Name].Metric }
    # CIM 没返回这块网卡（刚启用 / 只跑过 IPv6 等）：单独用 netsh 查一次，
    # 否则回读恒为 $null，会把"已经写成功"误判成"写入失败"
    $r = Invoke-Netsh 'interface ipv4 show interfaces'
    $started = $false
    foreach ($line in ($r.Output -split "`r?`n")) {
        if ($line -match '^-{3,}') { $started = $true; continue }
        if (-not $started) { continue }
        if ($line.Trim() -eq '') { continue }
        if ($line -match '^\s*(\d+)\s+(\d+)\s+(\d+)\s+(\S+)\s+(.+?)\s*$') {
            $met = [int]$Matches[2]; $ifName = $Matches[5]
            if ($ifName -eq $Name) { return $met }
        }
    }
    return $null
}

function Get-DefaultRoutes {
    # 解析 route print -4 中的默认路由（0.0.0.0/0）。
    # 用 (Metric, 原始顺序) 做稳定排序：跃点数相同时保持 route print 的原始顺序，
    # 避免 PowerShell 5.1 的 Sort-Object 在并列时顺序随机。
    $rows = @()
    $i = 0
    $r = Invoke-Route 'print -4'
    foreach ($line in ($r.Output -split "`r?`n")) {
        if ($line -match '^\s*0\.0\.0\.0\s+0\.0\.0\.0\s+(\S+)\s+(\S+)\s+(\d+)\s*$') {
            $gw = $Matches[1]; $ip = $Matches[2]; $met = [int]$Matches[3]
            $rows += New-Object psobject -Property @{ Gateway = $gw; InterfaceIp = $ip; Metric = $met; Order = $i }
            $i++
        }
    }
    return @($rows | Sort-Object Metric, Order)
}

function Test-AdminEnabled {
    param($AdminMap, [string]$Name)
    if ($AdminMap -and $AdminMap.ContainsKey($Name)) {
        return ($AdminMap[$Name].AdminState -match '^(Enabled|已启用)')
    }
    return $null   # 未知
}

function Get-AdapterSnapshot {
    <#
      采集全部网卡：.NET 枚举（已启用的）+ netsh 接口列表（含已禁用网卡）。
    #>
    $adminMap = Get-InterfaceAdminStateMap
    $metricMap = Get-InterfaceMetricMap
    $result = @()
    $seen = @{}

    foreach ($ni in (Get-AllInterfaces)) {
        $typeStr = [string]$ni.NetworkInterfaceType
        if ($typeStr -eq 'Loopback' -or $typeStr -eq 'Tunnel') { continue }

        $ipv4 = @(); $ipv6 = @(); $mask = ''; $gw = @(); $dns = @(); $dhcpSrv = @()
        try {
            $props = $ni.GetIPProperties()
            if ($props) {
                foreach ($u in $props.UnicastAddresses) {
                    if ($u.Address.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) {
                        if (-not ($u.Address.ToString() -like '169.254.*')) { $ipv4 += $u.Address.ToString() }
                        if ($u.IPv4Mask -and $u.IPv4Mask.ToString() -ne '0.0.0.0') { $mask = $u.IPv4Mask.ToString() }
                    } elseif ($u.Address.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetworkV6) {
                        $a = $u.Address.ToString()
                        if ($a -notmatch '^fe80' -and $a -notmatch '^::1') { $ipv6 += $a }
                    }
                }
                foreach ($g in $props.GatewayAddresses) {
                    $s = $g.Address.ToString()
                    if ($s -ne '0.0.0.0' -and $s -ne '::') { $gw += $s }
                }
                foreach ($d in $props.DnsAddresses) { $dns += $d.ToString() }
                foreach ($s in $props.DhcpServerAddresses) { $dhcpSrv += $s.ToString() }
            }
        } catch { }

        $speed = 0
        try { if ($ni.Speed -gt 0) { $speed = [math]::Round($ni.Speed / 1e6) } } catch { }

        $mac = ''
        try {
            $raw = $ni.GetPhysicalAddress().ToString()
            if ($raw.Length -eq 12) { $mac = (($raw -split '(.{2})' | Where-Object { $_ -ne '' }) -join '-') }
            else { $mac = $raw }
        } catch { }

        $statusText = switch ([string]$ni.OperationalStatus) {
            'Up'             { '已连接' }
            'Down'           { '未连接' }
            'Dormant'        { '休眠' }
            'NotPresent'     { '不存在' }
            'LowerLayerDown' { '链路断开' }
            default          { [string]$ni.OperationalStatus }
        }

        $met = $null; $idx = $null; $auto = $null
        if ($metricMap.ContainsKey($ni.Name)) {
            $met = $metricMap[$ni.Name].Metric
            $idx = $metricMap[$ni.Name].Index
            $auto = $metricMap[$ni.Name].Auto
        }

        $seen[$ni.Name] = $true
        $result += New-Object psobject -Property @{
            Name         = $ni.Name
            Description  = $ni.Description
            Type         = $typeStr
            Status       = [string]$ni.OperationalStatus
            StatusText   = $statusText
            SpeedMbps    = $speed
            Mac          = $mac
            IpAddresses  = $ipv4
            IpV6         = $ipv6
            SubnetMask   = $mask
            Gateways     = $gw
            Dns          = $dns
            DhcpServer   = $dhcpSrv
            IsDhcp       = ($dhcpSrv.Count -gt 0)
            IsWifi       = ($typeStr -eq 'Wireless80211')
            IsVirtual    = (Test-VirtualAdapter -Name $ni.Name -Description $ni.Description)
            AdminState   = $(if ($adminMap.ContainsKey($ni.Name)) { $adminMap[$ni.Name].AdminState } else { '' })
            AdminEnabled = (Test-AdminEnabled $adminMap $ni.Name)
            Metric       = $met
            MetricAuto   = $auto
            Index        = $idx
            Disabled     = $false
        }
    }

    # 补充：netsh 里存在、但 .NET 看不到的接口（通常是已被禁用的网卡）
    foreach ($name in @($adminMap.Keys)) {
        if ($seen.ContainsKey($name)) { continue }
        if ($name -match 'Loopback|环回') { continue }
        $isWifi = ($name -match 'WLAN|Wi-?Fi|Wireless|无线')
        $result += New-Object psobject -Property @{
            Name         = $name
            Description  = if ($isWifi) { '（已禁用的无线网卡）' } else { '（已禁用的网卡）' }
            Type         = if ($isWifi) { 'Wireless80211' } else { 'Ethernet' }
            Status       = 'Disabled'
            StatusText   = '已禁用'
            SpeedMbps    = 0
            Mac          = ''
            IpAddresses  = @()
            IpV6         = @()
            SubnetMask   = ''
            Gateways     = @()
            Dns          = @()
            DhcpServer   = @()
            IsDhcp       = $false
            IsWifi       = $isWifi
            IsVirtual    = $false
            AdminState   = $adminMap[$name].AdminState
            AdminEnabled = $false
            Metric       = $null
            MetricAuto   = $null
            Index        = $null
            Disabled     = $true
        }
    }
    return $result
}

function Get-LocalAdapters {
    # 注意：NetworkInterfaceType 不只是 'Ethernet'，还可能是 GigabitEthernet / FastEthernetT 等，
    # 所以这里用"非无线、非虚拟、非回环/隧道"来判断，避免漏掉物理网卡。
    param($Snapshot)
    if (-not $Snapshot) { $Snapshot = @(Get-AdapterSnapshot) }
    return @($Snapshot | Where-Object {
        -not $_.IsWifi -and -not $_.IsVirtual -and $_.Type -ne 'Loopback' -and $_.Type -ne 'Tunnel'
    })
}

function Get-WifiAdapters {
    param($Snapshot)
    if (-not $Snapshot) { $Snapshot = @(Get-AdapterSnapshot) }
    return @($Snapshot | Where-Object { $_.IsWifi -and -not $_.IsVirtual })
}

function Get-AdapterByName {
    param($Snapshot, [string]$Name)
    if (-not $Name) { return $null }
    $a = @($Snapshot | Where-Object { $_.Name -eq $Name } | Select-Object -First 1)
    if ($a.Count -eq 0) { return $null }
    return $a[0]
}

function Test-AdapterExists {
    param([string]$Name)
    if (-not $Name) { return $false }
    return ($null -ne (Get-AdapterByName @(Get-AdapterSnapshot) $Name))
}

function Get-WifiProfiles {
    $r = Invoke-Netsh 'wlan show profiles'
    $list = @()
    foreach ($line in ($r.Output -split "`r?`n")) {
        if ($line -match '^\s*(?:All User Profile|所有用户配置文件)\s*[:：]\s*(.+?)\s*$') {
            $n = $Matches[1].Trim()
            if ($n -ne '' -and $list -notcontains $n) { $list += $n }
        }
    }
    return $list
}

function Get-WlanBlocks {
    param([string]$Text)
    $blocks = @()
    $cur = $null
    foreach ($line in ($Text -split "`r?`n")) {
        if ($line -match '^\s*(Name|名称)\s*[:：]\s*(.+?)\s*$') {
            if ($cur) { $blocks += , $cur }
            $cur = @{ Name = $Matches[2]; Text = ($line + "`r`n") }
        } elseif ($cur) {
            $cur.Text += ($line + "`r`n")
        }
    }
    if ($cur) { $blocks += , $cur }
    return $blocks
}

function Get-WlanStatus {
    param([string]$InterfaceName)
    $res = New-Object psobject -Property @{
        Available = $false; Interface = $InterfaceName; State = ''; SSID = ''; BSSID = ''
        Signal = ''; SignalPercent = -1; RadioType = ''; Channel = ''; RxRate = ''
        TxRate = ''; Profile = ''; Error = ''; RadioOff = $false; Raw = ''
    }
    $r = Invoke-Netsh 'wlan show interfaces'
    $res.Raw = $r.Output
    if ($r.ExitCode -ne 0 -and $r.Output -notmatch 'SSID|State|状态') {
        $res.Error = ($r.Output.Trim())
        return $res
    }
    $blocks = @(Get-WlanBlocks $r.Output)
    if ($blocks.Count -eq 0) {
        $res.Error = '未找到无线接口信息'
        return $res
    }
    $blk = @($blocks | Where-Object { $_.Name -eq $InterfaceName } | Select-Object -First 1)
    if ($blk.Count -eq 0) { $blk = $blocks[0] } else { $blk = $blk[0] }

    # 只看"这张网卡自己"的块，避免多张无线网卡时互相污染
    if ($blk.Text -match 'Hardware Off|Software Off|硬件关闭|软件关闭|无线电已关闭') { $res.RadioOff = $true }

    $p = Get-NetshPairs $blk.Text
    $res.Available = $true
    $res.State     = Get-PairValue $p @('State', '状态')
    $res.SSID      = Get-PairValue $p @('SSID')
    $res.BSSID     = Get-PairValue $p @('BSSID')
    $res.Signal    = Get-PairValue $p @('Signal', '信号')
    $res.RadioType = Get-PairValue $p @('Radio type', '无线电类型')
    $res.Channel   = Get-PairValue $p @('Channel', '信道')
    $res.RxRate    = Get-PairValue $p @('Receive rate (Mbps)', '接收速率')
    $res.TxRate    = Get-PairValue $p @('Transmit rate (Mbps)', '传输速率')
    $res.Profile   = Get-PairValue $p @('Profile', '配置文件')
    if ($res.Signal -match '(\d+)') { $res.SignalPercent = [int]$Matches[1] }
    return $res
}

function Test-WlanConnected {
    param($WlanStatus)
    if (-not $WlanStatus -or -not $WlanStatus.Available) { return $false }
    if ($WlanStatus.State -notmatch 'connected|已连接') { return $false }
    if ($WlanStatus.State -match 'disconnected|未连接|已断开') { return $false }
    return $true
}

function Get-ActualExitRoute {
    <#
      判定"实际流量出口"。关键：必须按「接口跃点数 + 路由跃点数」取**最小值**，
      因为系统会同时返回多条默认路由，而它们的排列顺序与跃点数无关
      （route print / Find-NetRoute 的第一条不一定是真正生效的那条）。
      来源优先级：Get-NetRoute（CIM，实时跃点） -> Find-NetRoute -> 路由表解析
    #>
    param($Adapters)
    $cands = @()

    if (Get-Command Get-NetRoute -ErrorAction SilentlyContinue) {
        try {
            foreach ($r in @(Get-NetRoute -AddressFamily IPv4 -DestinationPrefix '0.0.0.0/0' -ErrorAction Stop)) {
                $alias = [string]$r.InterfaceAlias
                if (-not $alias) { continue }
                if ($null -eq $r.InterfaceMetric) { continue }
                $im = [int]$r.InterfaceMetric
                $rm = 0
                try { if ($null -ne $r.RouteMetric) { $rm = [int]$r.RouteMetric } } catch { }
                $cands += New-Object psobject -Property @{
                    InterfaceAlias = $alias; NextHop = [string]$r.NextHop
                    InterfaceMetric = $im; RouteMetric = $rm; Total = ($im + $rm); Source = 'Get-NetRoute'
                }
            }
        } catch { }
    }

    if ($cands.Count -eq 0 -and (Get-Command Find-NetRoute -ErrorAction SilentlyContinue)) {
        try {
            foreach ($r in @(Find-NetRoute -RemoteIPAddress '8.8.8.8' -ErrorAction Stop)) {
                if ($null -eq $r.NextHop) { continue }          # 跳过"源地址"对象
                $alias = [string]$r.InterfaceAlias
                if (-not $alias) { continue }
                if ($null -eq $r.InterfaceMetric) { continue }
                $im = [int]$r.InterfaceMetric
                $rm = 0
                try { if ($null -ne $r.RouteMetric) { $rm = [int]$r.RouteMetric } } catch { }
                $cands += New-Object psobject -Property @{
                    InterfaceAlias = $alias; NextHop = [string]$r.NextHop
                    InterfaceMetric = $im; RouteMetric = $rm; Total = ($im + $rm); Source = 'Find-NetRoute'
                }
            }
        } catch { }
    }

    if ($cands.Count -eq 0 -and $Adapters) {
        # 最后兜底：用已经解析好的路由表（Get-DefaultRoutes 的结果）
        foreach ($r in @($Adapters.Routes)) {
            $ad = @($Adapters.Adapters | Where-Object { $_.IpAddresses -contains $r.InterfaceIp } | Select-Object -First 1)
            if ($ad.Count -eq 0) { continue }
            $cands += New-Object psobject -Property @{
                InterfaceAlias = $ad[0].Name; NextHop = $r.Gateway
                InterfaceMetric = $ad[0].Metric; RouteMetric = 0
                Total = $(if ($null -ne $r.Metric) { [int]$r.Metric } else { 999999 }); Source = 'route print'
            }
        }
    }

    if ($cands.Count -eq 0) { return $null }
    return @($cands | Sort-Object Total, InterfaceAlias)[0]
}

function Get-CurrentExit {
    <#
      当前实际流量出口 = 默认路由中「接口跃点数+路由跃点数」最小的那条，
      再按接口 IP 反查是哪块网卡。
    #>
    param($Snapshot)
    $best = @($Snapshot.Routes | Select-Object -First 1)
    if ($best.Count -eq 0) { return $null }
    $r = $best[0]
    $a = @($Snapshot.Adapters | Where-Object { $_.IpAddresses -contains $r.InterfaceIp } | Select-Object -First 1)
    return New-Object psobject -Property @{
        Adapter     = $(if ($a.Count -gt 0) { $a[0] } else { $null })
        Name        = $(if ($a.Count -gt 0) { $a[0].Name } else { '' })
        Gateway     = $r.Gateway
        InterfaceIp = $r.InterfaceIp
        Metric      = $r.Metric
    }
}

function Get-NetworkSnapshot {
    $adapters = @(Get-AdapterSnapshot)
    $routes = @(Get-DefaultRoutes)
    $snap = New-Object psobject -Property @{
        Time        = Get-Date
        Adapters    = $adapters
        Routes      = $routes
        Active      = $null
        Exit        = $null
        ActualRoute = $null
        Wlan        = $null
        WifiCards   = @()
    }
    $snap.Exit = Get-CurrentExit $snap
    # 用「接口跃点+路由跃点」最小者判定真实出口（route print 的排列顺序不可信）
    $actual = Get-ActualExitRoute $snap
    $snap.ActualRoute = $actual
    if ($actual -and $actual.InterfaceAlias) {
        $aa = Get-AdapterByName $adapters $actual.InterfaceAlias
        if ($aa) {
            $snap.Exit = New-Object psobject -Property @{
                Adapter     = $aa
                Name        = $aa.Name
                Gateway     = $actual.NextHop
                InterfaceIp = $(if ($aa.IpAddresses.Count -gt 0) { $aa.IpAddresses[0] } else { '' })
                Metric      = $actual.InterfaceMetric
                Total       = $actual.Total
                Source      = $actual.Source
            }
        }
    }
    if ($snap.Exit -and $snap.Exit.Adapter) { $snap.Active = $snap.Exit.Adapter }
    else {
        $cands = @($adapters | Where-Object { $_.Status -eq 'Up' -and -not $_.IsVirtual })
        $withGw = @($cands | Where-Object { $_.Gateways.Count -gt 0 } | Select-Object -First 1)
        if ($withGw.Count -gt 0) { $snap.Active = $withGw[0] }
        elseif (@($cands).Count -gt 0) { $snap.Active = @($cands)[0] }
    }
    $snap.WifiCards = @(Get-WifiAdapters $adapters)
    $wlan = $null
    foreach ($w in $snap.WifiCards) {
        $st = Get-WlanStatus -InterfaceName $w.Name
        if ($st.Available -and (Test-WlanConnected $st)) { $wlan = $st; break }
    }
    if (-not $wlan -and $snap.WifiCards.Count -gt 0) { $wlan = Get-WlanStatus -InterfaceName $snap.WifiCards[0].Name }
    $snap.Wlan = $wlan
    return $snap
}

function Format-Metric {
    param($Metric)
    if ($null -eq $Metric) { return '-' }
    return [string]$Metric
}

function Format-MetricText {
    # 显示跃点数，并标明它是否仍是 Windows 自动度量值
    param($Adapter)
    if ($null -eq $Adapter -or $null -eq $Adapter.Metric) { return '-' }
    if ($Adapter.MetricAuto -eq $true) { return ('{0} (自动)' -f $Adapter.Metric) }
    return [string]$Adapter.Metric
}

function Format-NetworkReport {
    param($Snapshot)
    $sb = New-Object System.Text.StringBuilder
    $exitName = ''
    if ($Snapshot.Exit -and $Snapshot.Exit.Name) { $exitName = $Snapshot.Exit.Name }

    $null = $sb.AppendLine(("===== 网络连接信息  {0} =====" -f $Snapshot.Time.ToString('yyyy-MM-dd HH:mm:ss')))
    $null = $sb.AppendLine('')

    if ($Snapshot.Active) {
        $a = $Snapshot.Active
        $kind = if ($a.IsWifi) { '无线 (WiFi)' } else { '有线本地网络' }
        $null = $sb.AppendLine('--- 当前流量出口 ---')
        $null = $sb.AppendLine(('连接名称 : {0}' -f $a.Name))
        $null = $sb.AppendLine(('设备型号 : {0}' -f $a.Description))
        $null = $sb.AppendLine(('连接类型 : {0}' -f $kind))
        $null = $sb.AppendLine(('连接状态 : {0}' -f $a.StatusText))
        $null = $sb.AppendLine(('接口跃点 : {0}   （数值越小越优先；带"自动"表示仍由 Windows 自动度量）' -f (Format-MetricText $a)))
        $null = $sb.AppendLine(('链路速率 : {0}' -f $(if ($a.SpeedMbps -gt 0) { "$($a.SpeedMbps) Mbps" } else { '未知' })))
        $null = $sb.AppendLine(('MAC 地址 : {0}' -f $(if ($a.Mac) { $a.Mac } else { '未知' })))
        $null = $sb.AppendLine(('IPv4 地址 : {0}' -f $(if ($a.IpAddresses.Count -gt 0) { $a.IpAddresses -join ', ' } else { '无' })))
        $null = $sb.AppendLine(('子网掩码 : {0}' -f $(if ($a.SubnetMask) { $a.SubnetMask } else { '无' })))
        $null = $sb.AppendLine(('默认网关 : {0}' -f $(if ($a.Gateways.Count -gt 0) { $a.Gateways -join ', ' } else { '无' })))
        $null = $sb.AppendLine(('DNS 服务器 : {0}' -f $(if ($a.Dns.Count -gt 0) { $a.Dns -join ', ' } else { '无' })))
        $null = $sb.AppendLine(('DHCP 状态 : {0}' -f $(if ($a.IsDhcp) { "是（服务器 $($a.DhcpServer -join ', ')）" } elseif ($a.Status -eq 'Up') { '否（静态配置）' } else { '未知' })))
        if ($a.IpV6.Count -gt 0) { $null = $sb.AppendLine(('IPv6 地址 : {0}' -f ($a.IpV6[0]))) }
        $null = $sb.AppendLine('')
    } else {
        $null = $sb.AppendLine('--- 当前没有可用的默认路由（无网络出口）---')
        $null = $sb.AppendLine('')
    }

    $null = $sb.AppendLine('--- 默认路由（生效跃点＝接口跃点＋路由跃点，最小者才是真正的出口）---')
    if (@($Snapshot.Routes).Count -eq 0) {
        $null = $sb.AppendLine('  （无默认路由）')
    } else {
        foreach ($r in $Snapshot.Routes) {
            $nm = ''; $im = $null
            $ad = @($Snapshot.Adapters | Where-Object { $_.IpAddresses -contains $r.InterfaceIp } | Select-Object -First 1)
            if ($ad.Count -gt 0) { $nm = $ad[0].Name; $im = $ad[0].Metric }
            $mark = if ($nm -and $nm -eq $exitName) { '   <-- 当前出口' } else { '' }
            $null = $sb.AppendLine(('  网关 {0,-16} 本机 {1,-16} 网卡 {2,-14} 接口跃点 {3,-7} 生效跃点 {4,-7}{5}' -f $r.Gateway, $r.InterfaceIp, $nm, (Format-Metric $im), $r.Metric, $mark))
        }
    }
    if ($Snapshot.ActualRoute) {
        $null = $sb.AppendLine(('  出口判定依据：{0}（总跃点 {1}）' -f $Snapshot.ActualRoute.Source, $Snapshot.ActualRoute.Total))
    }
    $null = $sb.AppendLine('')

    if ($Snapshot.Wlan) {
        $w = $Snapshot.Wlan
        $null = $sb.AppendLine('--- 无线信息 ---')
        if ($w.Error) {
            $null = $sb.AppendLine(('无线状态 : 无法读取（{0}）' -f ($w.Error -replace "`r?`n", ' ')))
            $null = $sb.AppendLine('           提示：读取无线信息需要管理员权限。')
        } elseif (Test-WlanConnected $w) {
            $null = $sb.AppendLine(('SSID     : {0}' -f $w.SSID))
            $null = $sb.AppendLine(('信号强度 : {0}' -f $(if ($w.Signal) { $w.Signal } else { '未知' })))
            $null = $sb.AppendLine(('无线电类型 : {0}' -f $w.RadioType))
            $null = $sb.AppendLine(('信道     : {0}' -f $w.Channel))
            $null = $sb.AppendLine(('接收/发送 : {0} / {1} Mbps' -f $w.RxRate, $w.TxRate))
            $null = $sb.AppendLine(('无线网卡 : {0}' -f $w.Interface))
        } else {
            $state = if ($w.State) { $w.State } else { '未连接' }
            $null = $sb.AppendLine(('无线状态 : {0}（无线网卡：{1}）' -f $state, $w.Interface))
        }
        $null = $sb.AppendLine('')
    }

    $null = $sb.AppendLine('--- 全部网络适配器 ---')
    $sorted = @($Snapshot.Adapters | Sort-Object -Property @{ Expression = { if ($_.Disabled) { 3 } elseif ($_.Name -eq $exitName) { 0 } elseif ($_.Status -eq 'Up') { 1 } else { 2 } } }, Name)
    foreach ($a in $sorted) {
        $kind = if ($a.IsWifi) { '无线' } elseif ($a.IsVirtual) { '虚拟' } else { '有线' }
        $ip = if ($a.IpAddresses.Count -gt 0) { $a.IpAddresses[0] } else { '-' }
        $flag = ''
        if ($a.Name -eq $exitName) { $flag = '  <-- 当前流量出口' }
        elseif ($a.Disabled) { $flag = '  [已禁用]' }
        $null = $sb.AppendLine(('  {0,-14} {1,-4} {2,-8} {3,-16} 跃点 {4,-10}{5}' -f $a.Name, $kind, $a.StatusText, $ip, (Format-MetricText $a), $flag))
    }
    return $sb.ToString()
}

# ---------------------------------------------------------------------
# 网卡控制
# ---------------------------------------------------------------------
function Set-AdapterAdminState {
    param([string]$Name, [bool]$Enabled)
    $word = if ($Enabled) { 'enabled' } else { 'disabled' }
    $r = Invoke-Netsh ('interface set interface name=' + (Quote-NetshArg $Name) + ' admin=' + $word)
    if ($r.ExitCode -eq 0) { Clear-AdminStateCache; return $true }

    $script:LastError = $r.Output.Trim()
    try {
        if (Get-Command Enable-NetAdapter -ErrorAction SilentlyContinue) {
            if ($Enabled) { Enable-NetAdapter -Name $Name -Confirm:$false -ErrorAction Stop }
            else { Disable-NetAdapter -Name $Name -Confirm:$false -ErrorAction Stop }
            $script:LastError = ''
            Clear-AdminStateCache
            return $true
        }
    } catch {
        $script:LastError = ($script:LastError + ' ' + $_.Exception.Message).Trim()
    }
    return $false
}

function Set-InterfaceMetric {
    <#
      设置接口跃点数。判定成功的唯一标准是「写入后回读一致」。
        方法 1：Set-NetIPInterface（CIM）—— 先把「自动度量」关掉再写值，最可靠
        方法 2：netsh interface ipv4 set interface metric=...（兜底）
      实测教训：若"自动度量"仍启用，netsh 会返回退出码 0 但值立刻被 Windows 覆盖回自动值。
    #>
    param([string]$Name, [int]$Metric)
    $errors = @()

    if (Get-Command Set-NetIPInterface -ErrorAction SilentlyContinue) {
        # 1) 写活动存储：先关「自动度量」，再写跃点值（最可靠的写法）
        try { Set-NetIPInterface -InterfaceAlias $Name -AddressFamily IPv4 -AutomaticMetric Disabled -ErrorAction Stop }
        catch { $errors += ('关闭自动度量失败: ' + $_.Exception.Message) }
        try { Set-NetIPInterface -InterfaceAlias $Name -AddressFamily IPv4 -InterfaceMetric $Metric -ErrorAction Stop }
        catch { $errors += ('设置跃点数失败: ' + $_.Exception.Message) }
        # 2) 尽力写入持久存储（重启后仍生效，失败不影响当前生效）
        try {
            Set-NetIPInterface -InterfaceAlias $Name -AddressFamily IPv4 -AutomaticMetric Disabled -PolicyStore PersistentStore -ErrorAction Stop
            Set-NetIPInterface -InterfaceAlias $Name -AddressFamily IPv4 -InterfaceMetric $Metric -PolicyStore PersistentStore -ErrorAction Stop
        } catch { }
        # 3) IPv6 一起调（失败不影响 IPv4）
        try {
            Set-NetIPInterface -InterfaceAlias $Name -AddressFamily IPv6 -AutomaticMetric Disabled -ErrorAction Stop
            Set-NetIPInterface -InterfaceAlias $Name -AddressFamily IPv6 -InterfaceMetric $Metric -ErrorAction Stop
        } catch { }
        $after = Get-InterfaceMetricOf $Name
        if ($after -eq $Metric) {
            Write-Log "跃点数已设置(CIM，已回读校验)：[$Name] = $Metric"
            return $true
        }
        $errors += ("CIM 写入后回读=$after")
    } else {
        $errors += 'Set-NetIPInterface 不可用'
    }

    $r = Invoke-Netsh ('interface ipv4 set interface interface=' + (Quote-NetshArg $Name) + ' metric=' + $Metric + ' store=persistent')
    $null = Invoke-Netsh ('interface ipv6 set interface interface=' + (Quote-NetshArg $Name) + ' metric=' + $Metric + ' store=persistent')
    $after2 = Get-InterfaceMetricOf $Name
    if ($after2 -eq $Metric) {
        Write-Log "跃点数已设置(netsh，已回读校验)：[$Name] = $Metric"
        return $true
    }
    $errors += ('netsh exit={0} out=[{1}] 回读={2}' -f $r.ExitCode, (($r.Output.Trim()) -replace "`r?`n", ' '), $after2)

    $script:LastError = ($errors -join ' ; ')
    Write-Log "跃点数设置失败：[$Name] $script:LastError"
    return $false
}

function Restore-AutomaticMetric {
    param([string[]]$Names)
    $ok = @(); $fail = @()
    foreach ($n in $Names) {
        $done = $false
        try {
            if (Get-Command Set-NetIPInterface -ErrorAction SilentlyContinue) {
                Set-NetIPInterface -InterfaceAlias $n -AddressFamily IPv4 -AutomaticMetric Enabled -ErrorAction Stop
                try { Set-NetIPInterface -InterfaceAlias $n -AddressFamily IPv6 -AutomaticMetric Enabled -ErrorAction Stop } catch { }
                $done = $true
            }
        } catch { $script:LastError = $_.Exception.Message }
        if ($done) { $ok += $n; Write-Log "已恢复自动跃点：[$n]" } else { $fail += $n; Write-Log "恢复自动跃点失败：[$n] $script:LastError" }
    }
    return New-Object psobject -Property @{ Ok = $ok; Failed = $fail }
}

function Wait-AdapterAdminState {
    param([string]$Name, [bool]$Enabled, [int]$TimeoutSec = 30)
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    $unknownSince = $null
    while ((Get-Date) -lt $deadline) {
        $cur = Test-AdminEnabled (Get-InterfaceAdminStateMap -Force) $Name
        if ($null -ne $cur -and $cur -eq $Enabled) { return $true }
        if ($null -eq $cur) {
            # 读不到状态（编码/权限问题）就别干等了，最多等 4 秒
            if (-not $unknownSince) { $unknownSince = Get-Date }
            elseif (((Get-Date) - $unknownSince).TotalSeconds -ge 4) { return $false }
        } else {
            $unknownSince = $null
        }
        try { [System.Windows.Forms.Application]::DoEvents() } catch { }
        Start-Sleep -Milliseconds 600
    }
    return $false
}

function Wait-AdapterGateway {
    param([string]$Name, [int]$TimeoutSec = 30)
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        $a = Get-AdapterByName @(Get-AdapterSnapshot) $Name
        if ($a -and $a.Gateways.Count -gt 0) { return $true }
        try { [System.Windows.Forms.Application]::DoEvents() } catch { }
        Start-Sleep -Milliseconds 700
    }
    return $false
}

function Get-VisibleWifiNetworks {
    <#
      扫描当前可见的无线网络（netsh wlan show networks）。
      注意：Win11 需要开启"位置服务"权限才允许读取无线网络列表，
      权限不足时返回空数组（错误信息放在 .Error 里由调用方提示）。
    #>
    param([string]$Wifi = '')
    $r = Invoke-Netsh 'wlan show networks'
    if ($r.ExitCode -ne 0 -or $r.Output -match 'location permission|位置|elevation|提升') {
        if ($r.Output -notmatch 'SSID\s+\d+') { return @() }
    }
    $list = @()
    $cur = $null
    foreach ($line in ($r.Output -split "`r?`n")) {
        if ($line -match '^\s*SSID\s+\d+\s*[:：]\s*(.*?)\s*$') {
            if ($cur -and $cur.SSID) { $list += $cur }
            $cur = @{ SSID = $Matches[1]; SignalPercent = -1; Signal = ''; Auth = '' }
        } elseif ($cur) {
            if ($line -match '^\s*(Signal|信号)\s*[:：]\s*(.+?)\s*$') {
                $cur.Signal = $Matches[2]
                if ($Matches[2] -match '(\d+)') { $cur.SignalPercent = [int]$Matches[1] }
            } elseif ($line -match '^\s*(Authentication|身份验证|认证)\s*[:：]\s*(.+?)\s*$') {
                $cur.Auth = $Matches[2]
            }
        }
    }
    if ($cur -and $cur.SSID) { $list += $cur }
    return @($list | Sort-Object -Property SignalPercent -Descending)
}

function Select-BestWifiSsid {
    # 没配置 SSID 时：从「可见网络 ∩ 已保存配置文件」里挑信号最强的
    param([string]$Wifi = '')
    $profiles = @(Get-WifiProfiles)
    $nets = @(Get-VisibleWifiNetworks -Wifi $Wifi)
    $best = $null
    foreach ($n in $nets) {
        if ($profiles -contains $n.SSID) {
            if (-not $best -or $n.SignalPercent -gt $best.SignalPercent) { $best = $n }
        }
    }
    if ($best) { return [string]$best.SSID }
    return ''
}

function Connect-WifiNetworkEx {
    <#
      一次 WiFi 连接尝试：
        · 每 5 秒重发一次连接请求（不主动断开，避免打断正在进行的关联）
        · 到点就返回，不再自己"断开重连"（那一步交给上层：失败后重置网卡再试）
        · 单个网络默认只等 15 秒（可在设置里改），失败要快
    #>
    param([string]$InterfaceName, [string]$Ssid, [int]$TimeoutSec = 15)
    if ($TimeoutSec -lt 8) { $TimeoutSec = 8 }
    $start = Get-Date
    $deadline = $start.AddSeconds($TimeoutSec)
    $lastOut = ''
    $lastCmd = [datetime]::MinValue

    while ((Get-Date) -lt $deadline) {
        $st = Get-WlanStatus -InterfaceName $InterfaceName
        if (Test-WlanConnected $st) { return $true }

        if (((Get-Date) - $lastCmd).TotalSeconds -ge 5) {
            $r = Invoke-Netsh ('wlan connect name=' + (Quote-NetshArg $Ssid) + ' ssid=' + (Quote-NetshArg $Ssid) + ' interface=' + (Quote-NetshArg $InterfaceName))
            if ($r.ExitCode -ne 0) {
                $r = Invoke-Netsh ('wlan connect name=' + (Quote-NetshArg $Ssid) + ' interface=' + (Quote-NetshArg $InterfaceName))
            }
            $lastOut = $r.Output
            $lastCmd = Get-Date
        }

        try { [System.Windows.Forms.Application]::DoEvents() } catch { }
        Start-Sleep -Milliseconds 500
    }

    $st = Get-WlanStatus -InterfaceName $InterfaceName
    $script:LastError = ('等待{0}秒未连上；状态={1}；信号={2}；命令输出={3}' -f $TimeoutSec, $st.State, $(if ($st.Signal) { $st.Signal } else { '-' }), (($lastOut.Trim()) -replace "`r?`n", ' '))
    return $false
}

function Reset-WifiAdapter {
    <#
      无线网卡有时会卡在"正在连接 / 正在断开连接"状态，之后所有连接请求都无效。
      禁用再启用一次是最有效的恢复手段。
    #>
    param([string]$Name)
    Write-Log ("[WiFi] 重置无线网卡 [{0}]（禁用再启用）" -f $Name)
    $null = Set-AdapterAdminState -Name $Name -Enabled $false
    $null = Wait-AdapterAdminState -Name $Name -Enabled $false -TimeoutSec 10
    Start-Sleep -Milliseconds 800
    $null = Set-AdapterAdminState -Name $Name -Enabled $true
    $null = Wait-AdapterAdminState -Name $Name -Enabled $true -TimeoutSec 20
    Start-Sleep -Milliseconds 1200
}

function Get-AdapterForRouteIp {
    param($Adapters, [string]$Ip)
    if (-not $Ip) { return $null }
    $a = @($Adapters | Where-Object { $_.IpAddresses -contains $Ip } | Select-Object -First 1)
    if ($a.Count -eq 0) { return $null }
    return $a[0]
}

function Set-DefaultRouteMetric {
    <#
      ★ 关键修复 ★
      Windows 里默认路由自身也有一个"路由跃点"，它是路由创建时写死的，
      改接口跃点 **不会** 同步更新它。所以"只对调两块网卡的接口跃点"会出现：
         以太网   接口跃点 10   + 路由跃点 1000（旧值） = 1010
         以太网 3 接口跃点 1000 + 路由跃点 10（旧值）   = 1010
      两边正好抵消 → 路由表打平 → 实际流量不变。
      这里用 route change 直接显式改写默认路由自身的跃点。
    #>
    param([string]$Gateway, $IfIndex, [int]$Metric)
    if (-not $Gateway) { return $false }
    $cmd = 'change 0.0.0.0 mask 0.0.0.0 ' + $Gateway + ' metric ' + $Metric
    if ($IfIndex) { $cmd += ' if ' + $IfIndex }
    $r = Invoke-Route $cmd
    if ($r.ExitCode -ne 0) { $script:LastError = ($r.Output.Trim() -replace "`r?`n", ' ') }
    return ($r.ExitCode -eq 0)
}

function Test-ExitIsTarget {
    # 路由表里生效跃点最小的那条默认路由，是否属于目标网卡
    param([string]$Target, $Adapters)
    $defs = @(Get-DefaultRoutes)
    if ($defs.Count -eq 0) { return $false }
    $ad = Get-AdapterForRouteIp $Adapters $defs[0].InterfaceIp
    if (-not $ad) { return $false }
    return ($ad.Name -eq $Target)
}

function Set-ExitRoute {
    <#
      把流量出口真正切到 $Target：
        1) 接口跃点：目标 = PreferMetric，其它网卡 = OtherMetric
        2) 路由跃点：用 route change 给每条默认路由显式设定跃点（低/高）★核心
        3) 校验路由表；若目标仍不是最小，则删掉其它默认路由，只留目标那条
      返回：Verified / Messages / Target
    #>
    param([string]$Target, $Adapters, [int]$PreferMetric, [int]$OtherMetric)
    $msgs = @()

    # --- 1) 接口跃点
    # 注意要连"当前断开但参与切换"的网卡（尤其是无线）一起设：
    # 否则它断开时仍保留着低跃点，WiFi 一恢复就会把流量抢回去，看门狗也会反复触发。
    $cfgR = Get-Config
    $named = @([string]$cfgR.localA, [string]$cfgR.localB, [string]$cfgR.wifiAdapter) | Where-Object { $_ }
    foreach ($a in @($Adapters | Where-Object {
                (-not $_.Disabled) -and (
                    $_.Name -eq $Target -or
                    ($named -contains $_.Name) -or
                    ($_.Status -eq 'Up' -and $_.Gateways.Count -gt 0)
                )
            })) {
        $m = if ($a.Name -eq $Target) { $PreferMetric } else { $OtherMetric }
        if (Set-InterfaceMetric -Name $a.Name -Metric $m) {
            $msgs += ('接口跃点 [{0}] = {1}' -f $a.Name, (Get-InterfaceMetricOf $a.Name))
        } else {
            $msgs += ('接口跃点设置失败 [{0}] {1}' -f $a.Name, $script:LastError)
        }
    }

    # --- 2) 确保目标网卡有默认路由（之前可能被删掉过）
    $hasTargetRoute = $false
    foreach ($d in @(Get-DefaultRoutes)) {
        $ad = Get-AdapterForRouteIp $Adapters $d.InterfaceIp
        if ($ad -and $ad.Name -eq $Target) { $hasTargetRoute = $true; break }
    }
    if (-not $hasTargetRoute) {
        $tgtAd = Get-AdapterByName $Adapters $Target
        if ($tgtAd -and $tgtAd.Gateways.Count -gt 0) {
            $cmd = 'add 0.0.0.0 mask 0.0.0.0 ' + $tgtAd.Gateways[0] + ' metric ' + $PreferMetric
            if ($tgtAd.Index) { $cmd += ' if ' + $tgtAd.Index }
            $r = Invoke-Route $cmd
            $msgs += ('补充目标默认路由 exit={0} {1}' -f $r.ExitCode, (($r.Output.Trim()) -replace "`r?`n", ' '))
        }
    }

    # --- 3) 显式改写每条默认路由的跃点
    foreach ($d in @(Get-DefaultRoutes)) {
        $ad = Get-AdapterForRouteIp $Adapters $d.InterfaceIp
        if (-not $ad) { continue }
        $m = if ($ad.Name -eq $Target) { $PreferMetric } else { $OtherMetric }
        $ok = Set-DefaultRouteMetric -Gateway $d.Gateway -IfIndex $ad.Index -Metric $m
        $msgs += ('路由跃点 [{0}] 网关 {1} -> {2} : {3}' -f $ad.Name, $d.Gateway, $m, $(if ($ok) { 'OK' } else { '失败 ' + $script:LastError }))
    }

    # --- 4) 校验 + 兜底（兜底会删路由，所以必须保证删完仍有一条默认路由，并记录删除条数）
    $deleted = 0
    $verified = Test-ExitIsTarget -Target $Target -Adapters $Adapters
    if (-not $verified) {
        $msgs += '校验未通过：目标仍不是跃点最小的默认路由，改为删除其它默认路由'
        # 反查网卡用新鲜快照，避免 DHCP 换址后认错网卡（认错就可能删到目标那条）
        $fresh = @(Get-AdapterSnapshot)
        foreach ($d in @(Get-DefaultRoutes)) {
            $ad = Get-AdapterForRouteIp $fresh $d.InterfaceIp
            if (-not $ad) { continue }
            if ($ad.Name -eq $Target) { continue }
            $cmd = 'delete 0.0.0.0 mask 0.0.0.0 ' + $d.Gateway
            if ($ad.Index) { $cmd += ' if ' + $ad.Index }   # 带接口索引，避免同网关的其它接口路由被一起删掉
            $r = Invoke-Route $cmd
            $msgs += ('删除默认路由 [{0}] 网关 {1} if={2} exit={3}' -f $ad.Name, $d.Gateway, $ad.Index, $r.ExitCode)
            if ($r.ExitCode -eq 0) { $deleted++ }
        }
        # 万一删过头（一条默认路由都不剩），立刻把目标那条补回来，保证不断网
        if (@(Get-DefaultRoutes).Count -eq 0) {
            $tgtAd2 = Get-AdapterByName $fresh $Target
            if ($tgtAd2 -and $tgtAd2.Gateways.Count -gt 0) {
                $cmd = 'add 0.0.0.0 mask 0.0.0.0 ' + $tgtAd2.Gateways[0] + ' metric ' + $PreferMetric
                if ($tgtAd2.Index) { $cmd += ' if ' + $tgtAd2.Index }
                $r2 = Invoke-Route $cmd
                $msgs += ('【兜底】默认路由被删空，已补回目标路由 exit={0} {1}' -f $r2.ExitCode, (($r2.Output.Trim()) -replace "`r?`n", ' '))
                Write-Log ('[Set-ExitRoute] 默认路由被删空，已补回目标路由：' + $Target)
            }
        }
        $verified = Test-ExitIsTarget -Target $Target -Adapters $fresh
    }

    return New-Object psobject -Property @{ Verified = $verified; Messages = $msgs; Target = $Target; DeletedRoutes = $deleted }
}

# ---------------------------------------------------------------------
# 界面辅助
# ---------------------------------------------------------------------
function Show-Msg {
    param([string]$Text, [string]$Title = '网络切换工具', [string]$Kind = 'Info')
    $icon = [System.Windows.Forms.MessageBoxIcon]::Information
    if ($Kind -eq 'Warn')  { $icon = [System.Windows.Forms.MessageBoxIcon]::Warning }
    if ($Kind -eq 'Error') { $icon = [System.Windows.Forms.MessageBoxIcon]::Error }
    if ($script:MainForm) {
        return [System.Windows.Forms.MessageBox]::Show($script:MainForm, $Text, $Title, [System.Windows.Forms.MessageBoxButtons]::OK, $icon)
    }
    return [System.Windows.Forms.MessageBox]::Show($Text, $Title, [System.Windows.Forms.MessageBoxButtons]::OK, $icon)
}

function Confirm-Msg {
    param([string]$Text, [string]$Title = '请确认')
    if ($script:MainForm) {
        return ([System.Windows.Forms.MessageBox]::Show($script:MainForm, $Text, $Title, [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question) -eq [System.Windows.Forms.DialogResult]::Yes)
    }
    return ([System.Windows.Forms.MessageBox]::Show($Text, $Title, [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question) -eq [System.Windows.Forms.DialogResult]::Yes)
}

function New-UiFont {
    foreach ($name in @('Microsoft YaHei UI', 'Microsoft YaHei', 'Segoe UI')) {
        try { return (New-Object System.Drawing.Font($name, 9)) } catch { }
    }
    return [System.Drawing.SystemFonts]::DefaultFont
}

# ---------------------------------------------------------------------
# 提权
# ---------------------------------------------------------------------
function Ensure-Admin {
    if (Test-IsAdmin) { return $true }
    $yes = Confirm-Msg "调整跃点数、连接无线网络需要管理员权限。`r`n是否现在以管理员身份重新启动本工具？" '需要管理员权限'
    if ($yes) { Restart-Elevated }
    return $false
}

function Restart-Elevated {
    try {
        $exe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
        $argLine = '-NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $script:SelfPath + '"'
        Start-Process -FilePath $exe -ArgumentList $argLine -Verb RunAs
        Write-Log '已请求以管理员身份重启'
        if ($script:MainForm) { $script:MainForm.Close() }
        exit 0
    } catch {
        Show-Msg "提升管理员权限失败：$($_.Exception.Message)" '错误' 'Error'
        return $false
    }
}

# ---------------------------------------------------------------------
# 业务动作
# ---------------------------------------------------------------------
function Set-Busy {
    param([bool]$Busy, [string]$StatusText)
    $script:Busy = $Busy
    # 忙碌时把按钮面板上"所有"按钮都禁用：等待循环里有 DoEvents，
    # 否则"刷新信息/诊断/打开日志"能在切换中途插进来，读到中间态并回写 $script:LastSnapshot。
    $disabled = @($script:ActionButtons)
    if ($script:ButtonPanel) { $disabled = @($script:ButtonPanel.Controls) }
    foreach ($b in $disabled) {
        if ($b -is [System.Windows.Forms.Control]) { $b.Enabled = (-not $Busy) }
    }
    if ($StatusText -and $script:StatusLabel) { $script:StatusLabel.Text = $StatusText }
    if ($script:MainForm) {
        $script:MainForm.Cursor = if ($Busy) { [System.Windows.Forms.Cursors]::WaitCursor } else { [System.Windows.Forms.Cursors]::Default }
        try { [System.Windows.Forms.Application]::DoEvents() } catch { }
    }
}

function Get-ExitSummaryText {
    # 一行摘要：默认路由各自的有效跃点，最小者标为出口
    param($Adapters)
    $defs = @(Get-DefaultRoutes)
    if ($defs.Count -eq 0) { return '（无默认路由）' }
    $parts = @()
    for ($i = 0; $i -lt $defs.Count; $i++) {
        $ad = Get-AdapterForRouteIp $Adapters $defs[$i].InterfaceIp
        $nm = if ($ad) { $ad.Name } else { '?' }
        $mark = if ($i -eq 0) { '←出口' } else { '' }
        $parts += ('{0}={1}{2}' -f $nm, $defs[$i].Metric, $mark)
    }
    return ('跃点 ' + ($parts -join '，'))
}

function Invoke-LocalSwitch {
    $cfg = Get-Config
    $A = [string]$cfg.localA
    $B = [string]$cfg.localB

    if (-not $A -or -not $B) {
        $names = @(Get-LocalAdapters | ForEach-Object { $_.Name })
        $txt = '还没有配置两块本地网卡。' + "`r`n" + '请在「设置」中选择「本地网络 A」和「本地网络 B」。'
        if ($names.Count -gt 0) { $txt += "`r`n`r`n当前检测到的本地网卡：`r`n  " + ($names -join "`r`n  ") }
        else { $txt += "`r`n`r`n当前没有检测到任何有线网卡。" }
        Show-Msg $txt '未配置本地网络' 'Warn'
        Show-SettingsDialog
        return
    }
    if ($A -eq $B) { Show-Msg '「本地网络 A」和「本地网络 B」不能是同一块网卡。' '配置有误' 'Warn'; return }
    if (-not (Ensure-Admin)) { return }

    Set-Busy $true '正在读取网络状态…'
    $snap = Get-NetworkSnapshot
    $adA = Get-AdapterByName $snap.Adapters $A
    $adB = Get-AdapterByName $snap.Adapters $B

    $missing = @()
    if (-not $adA) { $missing += $A }
    if (-not $adB) { $missing += $B }
    if ($missing.Count -gt 0) {
        Set-Busy $false
        Show-Msg ("以下网卡当前不存在：`r`n  " + ($missing -join "`r`n") + "`r`n`r`n可能是设备被拔出，请在「设置」中重新选择。") '网卡不存在' 'Warn'
        return
    }

    # 目标选择：优先用「上次切换的目标」做严格 A<->B 交替（不受出口判定误差影响），
    # 没有记录时再按当前出口推断。
    $exitName = ''
    if ($snap.Exit -and $snap.Exit.Name) { $exitName = $snap.Exit.Name }
    $last = [string]$cfg.lastTarget
    if ($last -eq $A) { $tgtName = $B }
    elseif ($last -eq $B) { $tgtName = $A }
    elseif ($exitName -eq $A) { $tgtName = $B }
    elseif ($exitName -eq $B) { $tgtName = $A }
    elseif ($adA.Status -eq 'Up' -and $adA.Gateways.Count -gt 0) { $tgtName = $A }
    elseif ($adB.Status -eq 'Up' -and $adB.Gateways.Count -gt 0) { $tgtName = $B }
    else { $tgtName = $A }
    $curName = $exitName
    $tgt = Get-AdapterByName $snap.Adapters $tgtName

    Write-Log "流量出口切换：目标=[$tgtName] 原出口=[$curName]"
    $st = [System.Diagnostics.Stopwatch]::StartNew()

    try {
        # 只有目标网卡被禁用时才需要等待（启用 + 等地址）
        if (($tgt.AdminEnabled -eq $false) -or ($tgt.Disabled -eq $true)) {
            Set-Busy $true "正在启用网卡 [$tgtName]…"
            if (-not (Set-AdapterAdminState -Name $tgtName -Enabled $true)) {
                Set-Busy $false
                Update-Ui
                Show-Msg "启用网卡 [$tgtName] 失败：`r`n$script:LastError" '切换失败' 'Error'
                return
            }
            $null = Wait-AdapterAdminState -Name $tgtName -Enabled $true -TimeoutSec 15
            Set-Busy $true "已启用 [$tgtName]，正在等待获取网络地址…"
            if (-not (Wait-AdapterGateway -Name $tgtName -TimeoutSec ([int]$cfg.switchTimeoutSeconds))) {
                Set-Busy $false
                Update-Ui
                Show-Msg "[$tgtName] 已启用，但没有获取到默认网关（网线未插好或该网络未联网）。`r`n未做任何修改，流量出口保持不变。" '目标网络不可用' 'Warn'
                return
            }
            $snap = Get-NetworkSnapshot
            $tgt = Get-AdapterByName $snap.Adapters $tgtName
        }

        if (-not $tgt -or $tgt.Gateways.Count -eq 0) {
            Set-Busy $false
            Update-Ui
            Show-Msg "[$tgtName] 当前没有默认网关（未联网），无法把流量切过去。`r`n请检查网线 / 该网络是否需要认证。" '目标网络不可用' 'Warn'
            return
        }

        Set-Busy $true "正在把流量出口切到 [$tgtName]…"
        $res = Set-ExitRoute -Target $tgtName -Adapters $snap.Adapters -PreferMetric ([int]$cfg.preferMetric) -OtherMetric ([int]$cfg.deprioritizeMetric)
        foreach ($m in @($res.Messages)) { Write-Log ('  ' + $m) }
        if ($res.Verified) {
            # 记住本次目标，保证下次点击严格切到另一块（只写配置，不弹窗）
            $c2 = Get-Config
            $c2.lastTarget = $tgtName
            $null = Save-Config $c2
        }
        $st.Stop()
        $sum = Get-ExitSummaryText -Adapters $snap.Adapters

        Set-Busy $false
        if ($res.Verified) {
            $msg = ('已切换流量出口到 [{0}]  ({1:N1}s)  {2}' -f $tgtName, $st.Elapsed.TotalSeconds, $sum)
        } else {
            $msg = ('切换未生效：[{0}] 仍不是跃点最小的默认路由  {1}' -f $tgtName, $sum)
        }
        if ($script:StatusLabel) { $script:StatusLabel.Text = $msg }
        Write-Log $msg

        Update-Ui

        if (-not $res.Verified) {
            $tailMsg = if ([int]$res.DeletedRoutes -gt 0) {
                "`r`n`r`n注意：为强制切换已删除其它默认路由 {0} 条（已确认目标默认路由存在）。`r`n详细过程见 switch-log.txt。" -f $res.DeletedRoutes
            } else {
                "`r`n`r`n本地网卡都没有被动过，网络保持可用。`r`n详细过程见 switch-log.txt（每条命令的结果都有记录）。"
            }
            Show-Msg ($msg + $tailMsg) '切换失败' 'Error'
        }
    } catch {
        Set-Busy $false
        Show-Msg "切换失败：$($_.Exception.Message)" '切换失败' 'Error'
    }
}

function Invoke-WifiToggle {
    <#
      有线 <-> WiFi 双向切换：
        · 当前出口是无线  -> 切回有线（优先切到上次那块有线网卡；被禁用会自动启用）
        · 当前出口是有线  -> 连上 WiFi 再把流量出口切过去
      成功不弹窗，结果写状态栏 + 日志；只有做不了/失败才弹窗。
    #>
    $cfg = Get-Config
    $wifi = [string]$cfg.wifiAdapter
    if (-not $wifi) {
        $cards = @(Get-WifiAdapters)
        if ($cards.Count -gt 0) { $wifi = $cards[0].Name }
    }
    if (-not $wifi) {
        Show-Msg '没有检测到无线网卡，请在「设置」中手动指定。' '未找到无线网卡' 'Warn'
        Show-SettingsDialog
        return
    }
    if (-not (Test-AdapterExists $wifi)) {
        Show-Msg "无线网卡 [$wifi] 当前不存在，请在「设置」中重新选择。" '网卡不存在' 'Warn'
        return
    }
    if (-not (Ensure-Admin)) { return }

    Set-Busy $true '正在读取网络状态…'
    $snap = Get-NetworkSnapshot
    $exitName = ''
    if ($snap.Exit -and $snap.Exit.Name) { $exitName = $snap.Exit.Name }
    $st = [System.Diagnostics.Stopwatch]::StartNew()

    # ==================== 方向一：当前出口是无线 -> 切回有线 ====================
    if ($exitName -eq $wifi) {
        $cands = @()
        foreach ($n in @([string]$cfg.lastWired, [string]$cfg.localA, [string]$cfg.localB)) {
            if (-not $n -or $n -eq $wifi) { continue }
            $a = Get-AdapterByName $snap.Adapters $n
            if ($a -and ($cands -notcontains $a.Name)) { $cands += $a.Name }
        }
        foreach ($a in @(Get-LocalAdapters $snap.Adapters | Where-Object { $_.Status -eq 'Up' -and $_.Gateways.Count -gt 0 })) {
            if ($cands -notcontains $a.Name) { $cands += $a.Name }
        }

        $target = ''
        foreach ($n in $cands) {
            $a = Get-AdapterByName $snap.Adapters $n
            if ($a -and $a.Status -eq 'Up' -and $a.Gateways.Count -gt 0) { $target = $n; break }
        }
        if (-not $target) {
            # 之前若勾选过"切到 WiFi 就禁用本地网卡"，这里把有线网卡重新启用
            foreach ($n in $cands) {
                $a = Get-AdapterByName $snap.Adapters $n
                if (-not $a) { continue }
                if (($a.AdminEnabled -eq $false) -or $a.Disabled) {
                    Set-Busy $true "正在启用有线网卡 [$n]…"
                    if (Set-AdapterAdminState -Name $n -Enabled $true) {
                        $null = Wait-AdapterAdminState -Name $n -Enabled $true -TimeoutSec 15
                        if (Wait-AdapterGateway -Name $n -TimeoutSec ([int]$cfg.switchTimeoutSeconds)) { $target = $n; break }
                    }
                }
            }
        }
        if (-not $target) {
            Set-Busy $false
            Show-Msg '当前没有可用的有线网络（网卡未连接，或没有默认网关），无法切回有线。' '无法切回有线' 'Warn'
            return
        }

        $snap = Get-NetworkSnapshot
        Write-Log "WiFi -> 有线：目标=[$target]"
        Set-Busy $true "正在把流量出口切回 [$target]…"
        $res = Set-ExitRoute -Target $target -Adapters $snap.Adapters -PreferMetric ([int]$cfg.preferMetric) -OtherMetric ([int]$cfg.deprioritizeMetric)
        foreach ($m in @($res.Messages)) { Write-Log ('  ' + $m) }
        $st.Stop()
        $sum = Get-ExitSummaryText -Adapters $snap.Adapters
        Set-Busy $false

        if ($res.Verified) {
            $c2 = Get-Config
            $c2.lastWired = $target
            $c2.lastTarget = ''
            $null = Save-Config $c2
            $msg = ('已从 WiFi 切回有线 [{0}]  ({1:N1}s)  {2}' -f $target, $st.Elapsed.TotalSeconds, $sum)
            if ($script:StatusLabel) { $script:StatusLabel.Text = $msg }
            Write-Log $msg
            Update-Ui
        } else {
            $msg = ('切回有线未生效：[{0}] 仍不是跃点最小的默认路由  {1}' -f $target, $sum)
            if ($script:StatusLabel) { $script:StatusLabel.Text = $msg }
            Write-Log $msg
            Update-Ui
            Show-Msg ($msg + "`r`n`r`n详见 switch-log.txt（每条命令的结果都有记录）。") '切换失败' 'Error'
        }
        return
    }

    # ==================== 方向二：当前出口不是无线 -> 连上 WiFi 并切过去 ====================
    Write-Log "有线 -> WiFi：网卡=[$wifi] 原出口=[$exitName]"
    try {
        # 无线服务没起来的话，netsh 连接会"成功"但永远连不上
        try {
            $svc = Get-Service -Name 'WlanSvc' -ErrorAction SilentlyContinue
            if ($svc -and $svc.Status -ne 'Running') {
                Write-Log '无线服务 WlanSvc 未运行，尝试启动'
                Start-Service -Name 'WlanSvc' -ErrorAction SilentlyContinue
                Start-Sleep -Milliseconds 1500
            }
        } catch { }

        $wlAd = Get-AdapterByName $snap.Adapters $wifi
        if ($wlAd -and (($wlAd.AdminEnabled -eq $false) -or $wlAd.Disabled)) {
            Set-Busy $true "正在启用无线网卡 [$wifi]…"
            if (-not (Set-AdapterAdminState -Name $wifi -Enabled $true)) {
                Set-Busy $false
                Update-Ui
                Show-Msg "启用无线网卡 [$wifi] 失败：`r`n$script:LastError" '切换失败' 'Error'
                return
            }
            $null = Wait-AdapterAdminState -Name $wifi -Enabled $true -TimeoutSec 20
            Write-Log "[$wifi] 已启用"
            Start-Sleep -Milliseconds 1500     # 给射频初始化留点时间
        }

        $wl = Get-WlanStatus -InterfaceName $wifi
        if ($wl.RadioOff) {
            Set-Busy $false
            Update-Ui
            Show-Msg "无线网卡的无线电被关闭了（飞行模式 / 硬件开关 / 无线服务未启动）。`r`n请先打开无线开关再试。" '无线已关闭' 'Warn'
            return
        }

        $connected = Test-WlanConnected $wl
        $wifiWait = [int]$cfg.wifiConnectSeconds
        if ($wifiWait -lt 8) { $wifiWait = 8 }
        if ($wifiWait -gt 120) { $wifiWait = 120 }
        if (-not $connected) {
            $wantSsid = [string]$cfg.wifiSsid

            # ---- 先触发一次扫描，并取到"当前可见网络 + 已保存配置" ----
            Set-Busy $true '正在扫描无线网络…'
            $null = Invoke-Netsh 'wlan show networks'
            Start-Sleep -Milliseconds 1500
            $visible = @(Get-VisibleWifiNetworks -Wifi $wifi)
            $profiles = @(Get-WifiProfiles)
            $visNames = @($visible | ForEach-Object { $_.SSID })

            # ---- 候选列表：配置的 SSID 优先，其次"可见且已保存"的其它网络（按信号强弱） ----
            $cand = @()
            if ($wantSsid) { $cand += $wantSsid }
            elseif ($visNames.Count -gt 0) {
                $auto = Select-BestWifiSsid -Wifi $wifi
                if ($auto) { $cand += $auto }
            }
            if ([bool]$cfg.wifiAltFallback) {
                foreach ($n in $visible) {
                    if (($profiles -contains $n.SSID) -and ($cand -notcontains $n.SSID)) { $cand += $n.SSID }
                }
            }

            if ($cand.Count -eq 0) {
                Set-Busy $false
                Update-Ui
                $tip = ''
                if ($visible.Count -eq 0) {
                    $tip = "`r`n`r`n读不到可见网络列表：Win11 需要在「设置 -> 隐私和安全性 -> 位置」打开位置服务，`r`nnetsh 才允许列出无线网络。"
                } else {
                    $tip = "`r`n`r`n当前可见：`r`n  " + (($visible | Select-Object -First 10 | ForEach-Object { '{0} ({1}){2}' -f $_.SSID, $_.Signal, $(if ($profiles -contains $_.SSID) { ' 已保存' } else { '' }) }) -join "`r`n  ")
                }
                Show-Msg ('没有可连接的无线网络。' + "`r`n" + '请在「设置」的「WiFi 名称」里选择要连接的网络。' + $tip) '未选择 WiFi' 'Warn'
                return
            }

            # 首选网络扫不到时先记一笔（不阻止尝试，可能刚好在扫描后出现）
            if ($wantSsid -and $visNames.Count -gt 0 -and ($visNames -notcontains $wantSsid)) {
                Write-Log ("[WiFi] 当前扫描不到「{0}」；可见且已保存的网络：{1}" -f $wantSsid, (($cand | Select-Object -Skip 1) -join '、'))
            }

            $totalBudget = [Math]::Max(40, ($wifiWait * 3))   # 所有尝试加起来的总上限，避免无限等
            $startAll = Get-Date
            $didReset = $false
            $connectedSsid = ''
            foreach ($c in $cand) {
                if (((Get-Date) - $startAll).TotalSeconds -ge $totalBudget) {
                    Write-Log ('[WiFi] 已达总等待上限 {0} 秒，停止尝试其它网络' -f $totalBudget)
                    break
                }
                # 快速失败：如果这个网络当前根本不在可见列表里，就别等满 wifiWait
                $trySec = $wifiWait
                if ($visNames.Count -gt 0 -and ($visNames -notcontains $c)) {
                    $trySec = [Math]::Min($wifiWait, 6)
                    Write-Log ('[WiFi] 「{0}」当前不在可见列表里，只快速尝试 {1} 秒' -f $c, $trySec)
                }
                Set-Busy $true ("正在连接无线网络 [{0}]（最多 {1} 秒）…" -f $c, $trySec)
                $connected = Connect-WifiNetworkEx -InterfaceName $wifi -Ssid $c -TimeoutSec $trySec
                if ($connected) { $connectedSsid = $c; break }
                Write-Log ("[WiFi] 连接「{0}」失败：{1}" -f $c, $script:LastError)

                if (-not $didReset) {
                    if (((Get-Date) - $startAll).TotalSeconds -ge ($totalBudget - 10)) {
                        Write-Log '[WiFi] 剩余时间不足，跳过"重置无线网卡"这一步'
                    } else {
                        # 网卡可能卡在"正在连接/正在断开连接"：重置一次再重试同一个网络
                        $didReset = $true
                        Set-Busy $true '正在重置无线网卡后重试…'
                        Reset-WifiAdapter -Name $wifi
                        $null = Invoke-Netsh 'wlan show networks'
                        Start-Sleep -Milliseconds 1000
                        if (((Get-Date) - $startAll).TotalSeconds -lt $totalBudget) {
                            $connected = Connect-WifiNetworkEx -InterfaceName $wifi -Ssid $c -TimeoutSec $trySec
                            if ($connected) { $connectedSsid = $c; break }
                            Write-Log ("[WiFi] 重置网卡后连接「{0}」仍失败：{1}" -f $c, $script:LastError)
                        }
                    }
                }
            }

            if ($connected) {
                $ssid = $connectedSsid
                Write-Log "已连接无线网络 [$ssid]"
            } else {
                $visible2 = @(Get-VisibleWifiNetworks -Wifi $wifi)
                $visTxt = ''
                if ($visible2.Count -gt 0) {
                    $visTxt = "`r`n`r`n当前可见网络：`r`n  " + (($visible2 | Select-Object -First 10 | ForEach-Object { '{0} ({1}){2}' -f $_.SSID, $_.Signal, $(if ($profiles -contains $_.SSID) { ' 已保存' } else { '' }) }) -join "`r`n  ")
                } else {
                    $visTxt = "`r`n`r`n（读不到可见网络列表：请在「设置 → 隐私和安全性 → 位置」打开位置服务）"
                }
                $msg = ("连接无线网络失败（已尝试：{0}；每个最多等 {1} 秒，并重置过无线网卡）。`r`n最后一次状态：{2}{3}`r`n`r`n请在「设置 → WiFi 名称」里改选一个当前可见的网络。`r`n本地网络未被改动，仍可正常上网。" -f ($cand -join '、'), $wifiWait, $script:LastError, $visTxt)
                Write-Log ('切换 WiFi 失败：' + ($msg -replace "`r?`n", ' '))
                Set-Busy $false
                Update-Ui
                Show-Msg $msg 'WiFi 连接失败' 'Error'
                return
            }
        }

        Set-Busy $true '正在等待无线网卡获取地址…'
        $null = Wait-AdapterGateway -Name $wifi -TimeoutSec ([Math]::Max(10, $wifiWait))
        $snap2 = Get-NetworkSnapshot
        $wlAd2 = Get-AdapterByName $snap2.Adapters $wifi
        if (-not $wlAd2 -or $wlAd2.Gateways.Count -eq 0) {
            Set-Busy $false
            Update-Ui
            Show-Msg "[$wifi] 已连接但没拿到默认网关（IP 未就绪），未修改任何路由，流量仍在原出口。" '无线未就绪' 'Warn'
            return
        }

        if ($cfg.disableLocalWhenWifi) {
            foreach ($a in @(Get-LocalAdapters $snap2.Adapters | Where-Object { $_.Name -ne $wifi })) {
                Set-Busy $true "正在禁用本地网卡 [$($a.Name)]…"
                if (Set-AdapterAdminState -Name $a.Name -Enabled $false) {
                    $null = Wait-AdapterAdminState -Name $a.Name -Enabled $false -TimeoutSec 8
                    Write-Log "已禁用本地网卡 [$($a.Name)]"
                } else {
                    Write-Log "禁用本地网卡 [$($a.Name)] 失败：$script:LastError"
                }
            }
        }

        Set-Busy $true "正在把流量出口切到 [$wifi]…"
        $res = Set-ExitRoute -Target $wifi -Adapters $snap2.Adapters -PreferMetric ([int]$cfg.preferMetric) -OtherMetric ([int]$cfg.deprioritizeMetric)
        foreach ($m in @($res.Messages)) { Write-Log ('  ' + $m) }
        $st.Stop()
        $sum = Get-ExitSummaryText -Adapters $snap2.Adapters
        Set-Busy $false

        if ($res.Verified) {
            $c2 = Get-Config
            # 只在确实是一块"非无线"的出口时才记，避免记成空串或无线网卡本身
            if ($exitName -and $exitName -ne $wifi) { $c2.lastWired = $exitName }
            $null = Save-Config $c2
            $ws = Get-WlanStatus -InterfaceName $wifi
            $msg = ('已切到 WiFi [{0}] 信号 {1}  ({2:N1}s)  {3}' -f $ws.SSID, $(if ($ws.Signal) { $ws.Signal } else { '-' }), $st.Elapsed.TotalSeconds, $sum)
            if ($script:StatusLabel) { $script:StatusLabel.Text = $msg }
            Write-Log $msg
            Update-Ui
        } else {
            $msg = ('切到 WiFi 未生效：[{0}] 仍不是跃点最小的默认路由  {1}' -f $wifi, $sum)
            if ($script:StatusLabel) { $script:StatusLabel.Text = $msg }
            Write-Log $msg
            Update-Ui
            Show-Msg ($msg + "`r`n`r`n详见 switch-log.txt（每条命令的结果都有记录）。") '切换失败' 'Error'
        }
    } catch {
        Set-Busy $false
        Show-Msg "切换到 WiFi 失败：$($_.Exception.Message)" '切换失败' 'Error'
    }
}

# ---------------------------------------------------------------------
# 设置窗口
# ---------------------------------------------------------------------
function Format-AdapterItem {
    param($Adapter)
    if ($Adapter.Disabled -or $Adapter.AdminEnabled -eq $false) { $state = '已禁用' }
    else { $state = $Adapter.StatusText }
    $tag = if ($Adapter.IsVirtual) { '虚拟设备' } else { '物理设备' }
    $met = Format-MetricText $Adapter
    return ('{0}  [{1} | {2} | {3} | 跃点 {4}]' -f $Adapter.Name, $Adapter.Description, $state, $tag, $met)
}

function Get-AdapterItemName {
    param([string]$Item)
    if (-not $Item) { return '' }
    if ($Item -like '*（未选择）*') { return '' }
    return ($Item -split '  \[')[0].Trim()
}

function Fill-SettingsCombos {
    param($CboA, $CboB, $CboW, $CboS, $Cfg)
    $snap = @(Get-AdapterSnapshot)
    # 本地网络候选：所有非无线网卡（含被禁用的、USB / 虚拟网卡，列表中会标注）
    $loc = @($snap | Where-Object { -not $_.IsWifi } | Sort-Object -Property @{ Expression = { if ($_.IsVirtual) { 1 } else { 0 } } }, Name)
    $wf  = @($snap | Where-Object { $_.IsWifi -and -not $_.IsVirtual } | Sort-Object Name)

    $CboA.Items.Clear(); $CboB.Items.Clear(); $CboW.Items.Clear()
    $null = $CboA.Items.Add('（未选择）')
    $null = $CboB.Items.Add('（未选择）')
    $null = $CboW.Items.Add('（未选择）')
    foreach ($a in $loc) {
        $null = $CboA.Items.Add((Format-AdapterItem $a))
        $null = $CboB.Items.Add((Format-AdapterItem $a))
    }
    foreach ($a in $wf) { $null = $CboW.Items.Add((Format-AdapterItem $a)) }

    # WiFi 名称候选：当前可见的网络（标注信号/是否已保存）+ 已保存但当前不可见的配置文件
    $CboS.Items.Clear()
    $profiles = @(Get-WifiProfiles)
    $seenSsid = @()
    foreach ($n in @(Get-VisibleWifiNetworks -Wifi '')) {
        $has = if ($profiles -contains $n.SSID) { '已保存' } else { '未保存' }
        $sig = if ($n.Signal) { $n.Signal } else { '?' }
        $null = $CboS.Items.Add(('{0}  [可见 {1} · {2}]' -f $n.SSID, $sig, $has))
        $seenSsid += $n.SSID
    }
    foreach ($p in $profiles) {
        if ($seenSsid -contains $p) { continue }
        $null = $CboS.Items.Add(('{0}  [已保存]' -f $p))
    }

    $selA = 0; $selB = 0; $selW = 0
    for ($i = 0; $i -lt $CboA.Items.Count; $i++) { if ((Get-AdapterItemName ([string]$CboA.Items[$i])) -eq [string]$Cfg.localA) { $selA = $i } }
    for ($i = 0; $i -lt $CboB.Items.Count; $i++) { if ((Get-AdapterItemName ([string]$CboB.Items[$i])) -eq [string]$Cfg.localB) { $selB = $i } }
    for ($i = 0; $i -lt $CboW.Items.Count; $i++) { if ((Get-AdapterItemName ([string]$CboW.Items[$i])) -eq [string]$Cfg.wifiAdapter) { $selW = $i } }

    if (-not $Cfg.localA -and $loc.Count -ge 1) { $selA = 1 }
    if (-not $Cfg.localB -and $loc.Count -ge 2) { $selB = 2 }
    if (-not $Cfg.wifiAdapter -and $wf.Count -ge 1) { $selW = 1 }

    $CboA.SelectedIndex = $selA
    $CboB.SelectedIndex = $selB
    $CboW.SelectedIndex = $selW

    $CboS.Text = [string]$Cfg.wifiSsid
    for ($i = 0; $i -lt $CboS.Items.Count; $i++) {
        if ((Get-AdapterItemName ([string]$CboS.Items[$i])) -eq [string]$Cfg.wifiSsid) { $CboS.SelectedIndex = $i; break }
    }
}

function Show-SettingsDialog {
    $cfg = Get-Config

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = '设置 — 网络切换工具'
    $dlg.ClientSize = New-Object System.Drawing.Size(620, 450)
    $dlg.StartPosition = 'CenterParent'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.Font = $script:UiFont

    $labels = @(
        @('本地网络 A：', 20), @('本地网络 B：', 58), @('无线网卡：', 96), @('WiFi 名称：', 134)
    )
    foreach ($l in $labels) {
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $l[0]
        $lbl.Location = New-Object System.Drawing.Point(18, ([int]$l[1] + 3))
        $lbl.Size = New-Object System.Drawing.Size(150, 22)
        $dlg.Controls.Add($lbl)
    }

    $cboA = New-Object System.Windows.Forms.ComboBox
    $cboA.Location = New-Object System.Drawing.Point(175, 18); $cboA.Size = New-Object System.Drawing.Size(420, 24); $cboA.DropDownStyle = 'DropDownList'
    $dlg.Controls.Add($cboA)

    $cboB = New-Object System.Windows.Forms.ComboBox
    $cboB.Location = New-Object System.Drawing.Point(175, 56); $cboB.Size = New-Object System.Drawing.Size(420, 24); $cboB.DropDownStyle = 'DropDownList'
    $dlg.Controls.Add($cboB)

    $cboW = New-Object System.Windows.Forms.ComboBox
    $cboW.Location = New-Object System.Drawing.Point(175, 94); $cboW.Size = New-Object System.Drawing.Size(420, 24); $cboW.DropDownStyle = 'DropDownList'
    $dlg.Controls.Add($cboW)

    $cboS = New-Object System.Windows.Forms.ComboBox
    $cboS.Location = New-Object System.Drawing.Point(175, 132); $cboS.Size = New-Object System.Drawing.Size(420, 24); $cboS.DropDownStyle = 'DropDown'
    $dlg.Controls.Add($cboS)

    $chkDisable = New-Object System.Windows.Forms.CheckBox
    $chkDisable.Text = '切换到 WiFi 时改为「禁用」本地有线网卡（默认不勾选：只切流量，不断开网卡）'
    $chkDisable.Location = New-Object System.Drawing.Point(175, 166)
    $chkDisable.Size = New-Object System.Drawing.Size(430, 24)
    $chkDisable.Checked = [bool]$cfg.disableLocalWhenWifi
    $dlg.Controls.Add($chkDisable)

    $lblRefresh = New-Object System.Windows.Forms.Label
    $lblRefresh.Text = '自动刷新(秒)：'
    $lblRefresh.Location = New-Object System.Drawing.Point(18, 200)
    $lblRefresh.Size = New-Object System.Drawing.Size(150, 22)
    $dlg.Controls.Add($lblRefresh)

    $numRefresh = New-Object System.Windows.Forms.NumericUpDown
    $numRefresh.Location = New-Object System.Drawing.Point(175, 198)
    $numRefresh.Size = New-Object System.Drawing.Size(70, 24)
    $numRefresh.Minimum = 0
    $numRefresh.Maximum = 3600
    $numRefresh.Value = [decimal][int]$cfg.autoRefreshSeconds
    $dlg.Controls.Add($numRefresh)

    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Text = '（0=不刷新）'
    $lblHint.Location = New-Object System.Drawing.Point(250, 203)
    $lblHint.Size = New-Object System.Drawing.Size(92, 22)
    $lblHint.ForeColor = [System.Drawing.Color]::Gray
    $dlg.Controls.Add($lblHint)

    $lblWifiWait = New-Object System.Windows.Forms.Label
    $lblWifiWait.Text = 'WiFi 等待(秒)：'
    $lblWifiWait.Location = New-Object System.Drawing.Point(348, 203)
    $lblWifiWait.Size = New-Object System.Drawing.Size(105, 22)
    $dlg.Controls.Add($lblWifiWait)

    $numWifiWait = New-Object System.Windows.Forms.NumericUpDown
    $numWifiWait.Location = New-Object System.Drawing.Point(455, 198)
    $numWifiWait.Size = New-Object System.Drawing.Size(70, 24)
    $numWifiWait.Minimum = 8
    $numWifiWait.Maximum = 120
    $numWifiWait.Value = [decimal][int]$cfg.wifiConnectSeconds
    $dlg.Controls.Add($numWifiWait)

    $lblWifiWaitHint = New-Object System.Windows.Forms.Label
    $lblWifiWaitHint.Text = '（连单个 WiFi 的最长等待）'
    $lblWifiWaitHint.Location = New-Object System.Drawing.Point(530, 203)
    $lblWifiWaitHint.Size = New-Object System.Drawing.Size(80, 22)
    $lblWifiWaitHint.ForeColor = [System.Drawing.Color]::Gray
    $dlg.Controls.Add($lblWifiWaitHint)

    # --- 跃点数设置
    $lblMetric = New-Object System.Windows.Forms.Label
    $lblMetric.Text = '跃点数(首选/其它)：'
    $lblMetric.Location = New-Object System.Drawing.Point(18, 236)
    $lblMetric.Size = New-Object System.Drawing.Size(150, 22)
    $dlg.Controls.Add($lblMetric)

    $numPrefer = New-Object System.Windows.Forms.NumericUpDown
    $numPrefer.Location = New-Object System.Drawing.Point(175, 234)
    $numPrefer.Size = New-Object System.Drawing.Size(78, 24)
    $numPrefer.Minimum = 1
    $numPrefer.Maximum = 9999
    $numPrefer.Value = [decimal][int]$cfg.preferMetric
    $dlg.Controls.Add($numPrefer)

    $lblSlash = New-Object System.Windows.Forms.Label
    $lblSlash.Text = '/'
    $lblSlash.Location = New-Object System.Drawing.Point(258, 237)
    $lblSlash.Size = New-Object System.Drawing.Size(14, 22)
    $dlg.Controls.Add($lblSlash)

    $numDeprior = New-Object System.Windows.Forms.NumericUpDown
    $numDeprior.Location = New-Object System.Drawing.Point(276, 234)
    $numDeprior.Size = New-Object System.Drawing.Size(80, 24)
    $numDeprior.Minimum = 1
    $numDeprior.Maximum = 9999
    $numDeprior.Value = [decimal][int]$cfg.deprioritizeMetric
    $dlg.Controls.Add($numDeprior)

    $lblMetricHint = New-Object System.Windows.Forms.Label
    $lblMetricHint.Text = '数值越小越优先；首选必须小于其它值'
    $lblMetricHint.Location = New-Object System.Drawing.Point(364, 237)
    $lblMetricHint.Size = New-Object System.Drawing.Size(240, 22)
    $lblMetricHint.ForeColor = [System.Drawing.Color]::Gray
    $dlg.Controls.Add($lblMetricHint)

    $chkAutoFallback = New-Object System.Windows.Forms.CheckBox
    $chkAutoFallback.Text = '网卡断开后自动把流量切到另一块可用网卡（WiFi↔有线、有线↔有线）'
    $chkAutoFallback.Location = New-Object System.Drawing.Point(175, 268)
    $chkAutoFallback.Size = New-Object System.Drawing.Size(430, 24)
    $chkAutoFallback.Checked = [bool]$cfg.autoFallback
    $dlg.Controls.Add($chkAutoFallback)

    $chkWifiAlt = New-Object System.Windows.Forms.CheckBox
    $chkWifiAlt.Text = '指定的 WiFi 连不上时，自动尝试其它可见且已保存的网络（推荐）'
    $chkWifiAlt.Location = New-Object System.Drawing.Point(175, 298)
    $chkWifiAlt.Size = New-Object System.Drawing.Size(430, 24)
    $chkWifiAlt.Checked = [bool]$cfg.wifiAltFallback
    $dlg.Controls.Add($chkWifiAlt)

    $lblTip = New-Object System.Windows.Forms.Label
    $lblTip.Text = "说明：切换流量出口只改跃点数与默认路由，网卡都保持启用，可随时来回切换，不会断网。`r`n已禁用的网卡也会出现在下拉框里（标注为已禁用），切到它时会自动先启用。`r`n「WiFi 等待(秒)」是连接单个无线网络的最长等待，默认 15 秒，失败会很快返回，不会长时间干等。"
    $lblTip.Location = New-Object System.Drawing.Point(18, 330)
    $lblTip.Size = New-Object System.Drawing.Size(590, 48)
    $lblTip.ForeColor = [System.Drawing.Color]::DimGray
    $dlg.Controls.Add($lblTip)

    $btnReload = New-Object System.Windows.Forms.Button
    $btnReload.Text = '重新检测网卡/WiFi'
    $btnReload.Location = New-Object System.Drawing.Point(18, 388)
    $btnReload.Size = New-Object System.Drawing.Size(128, 30)
    $btnReload.Add_Click({ Fill-SettingsCombos -CboA $cboA -CboB $cboB -CboW $cboW -CboS $cboS -Cfg (Get-Config) })
    $dlg.Controls.Add($btnReload)

    $btnAuto = New-Object System.Windows.Forms.Button
    $btnAuto.Text = '恢复自动跃点'
    $btnAuto.Location = New-Object System.Drawing.Point(152, 388)
    $btnAuto.Size = New-Object System.Drawing.Size(128, 30)
    $btnAuto.Add_Click({
        if (-not (Confirm-Msg "将把所有网卡的「接口跃点数」恢复为 Windows 自动度量（取消手动设定），是否继续？" '恢复自动跃点')) { return }
        if (-not (Ensure-Admin)) { return }
        $names = @(Get-AdapterSnapshot | Where-Object { -not $_.Disabled } | ForEach-Object { $_.Name })
        $res = Restore-AutomaticMetric -Names $names
        $msg = ''
        if ($res.Ok.Count -gt 0) { $msg += "已恢复自动跃点：`r`n  " + ($res.Ok -join "`r`n  ") }
        if ($res.Failed.Count -gt 0) { $msg += "`r`n`r`n失败：`r`n  " + ($res.Failed -join "`r`n  ") + "`r`n$script:LastError" }
        Show-Msg $msg '恢复自动跃点' $(if ($res.Failed.Count -gt 0) { 'Warn' } else { 'Info' })
        Fill-SettingsCombos -CboA $cboA -CboB $cboB -CboW $cboW -CboS $cboS -Cfg (Get-Config)
    })
    $dlg.Controls.Add($btnAuto)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = '保存'
    $btnOk.Location = New-Object System.Drawing.Point(375, 388)
    $btnOk.Size = New-Object System.Drawing.Size(100, 30)
    $btnOk.DialogResult = 'OK'
    $dlg.Controls.Add($btnOk)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = '取消'
    $btnCancel.Location = New-Object System.Drawing.Point(490, 388)
    $btnCancel.Size = New-Object System.Drawing.Size(100, 30)
    $btnCancel.DialogResult = 'Cancel'
    $dlg.Controls.Add($btnCancel)

    $dlg.AcceptButton = $btnOk
    $dlg.CancelButton = $btnCancel

    Fill-SettingsCombos -CboA $cboA -CboB $cboB -CboW $cboW -CboS $cboS -Cfg $cfg

    $result = $dlg.ShowDialog($script:MainForm)
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        if ([int]$numPrefer.Value -ge [int]$numDeprior.Value) {
            Show-Msg '「首选跃点数」必须小于「其它跃点数」，例如 10 和 1000。请修改后重试。' '设置无效' 'Warn'
            $dlg.Dispose()
            return
        }
        $newCfg = Get-Config
        $oldA = [string]$newCfg.localA
        $oldB = [string]$newCfg.localB
        $newCfg.localA = Get-AdapterItemName ([string]$cboA.SelectedItem)
        $newCfg.localB = Get-AdapterItemName ([string]$cboB.SelectedItem)
        if ($newCfg.localA -ne $oldA -or $newCfg.localB -ne $oldB) { $newCfg.lastTarget = '' }
        $newCfg.wifiAdapter = Get-AdapterItemName ([string]$cboW.SelectedItem)
        $newCfg.wifiSsid = Get-AdapterItemName ([string]$cboS.Text)
        $newCfg.disableLocalWhenWifi = [bool]$chkDisable.Checked
        $newCfg.autoFallback = [bool]$chkAutoFallback.Checked
        $newCfg.wifiAltFallback = [bool]$chkWifiAlt.Checked
        $newCfg.wifiConnectSeconds = [int]$numWifiWait.Value
        $newCfg.autoRefreshSeconds = [int]$numRefresh.Value
        $newCfg.preferMetric = [int]$numPrefer.Value
        $newCfg.deprioritizeMetric = [int]$numDeprior.Value
        if (Save-Config $newCfg) {
            Apply-RefreshInterval
            Update-Ui
            if ($script:StatusLabel) { $script:StatusLabel.Text = '设置已保存' }
        } else {
            Show-Msg '保存配置失败，请检查程序目录是否可写。' '保存失败' 'Error'
        }
    }
    $dlg.Dispose()
}

# ---------------------------------------------------------------------
# 诊断报告
# ---------------------------------------------------------------------
function Write-Diagnostics {
    $L = @()
    $L += '===== 网络切换工具 诊断报告 ====='
    $L += ('时间        : {0}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
    $L += ('管理员权限  : {0}' -f (Test-IsAdmin))
    $L += ('脚本路径    : {0}' -f $script:SelfPath)
    $cfg = Get-Config
    $L += ('配置        : localA={0} localB={1} wifi={2} ssid={3} 首选跃点={4} 其它跃点={5} 上次目标={6}' -f $cfg.localA, $cfg.localB, $cfg.wifiAdapter, $cfg.wifiSsid, $cfg.preferMetric, $cfg.deprioritizeMetric, $cfg.lastTarget)
    $L += ''

    $L += '--- 1) Get-NetIPInterface -AddressFamily IPv4（接口跃点数/自动度量）---'
    $got = $false
    try {
        if (Get-Command Get-NetIPInterface -ErrorAction SilentlyContinue) {
            foreach ($o in @(Get-NetIPInterface -AddressFamily IPv4 -ErrorAction Stop)) {
                if (-not $o.InterfaceAlias) { continue }
                $got = $true
                $L += ('  {0,-16} idx={1,-4} 接口跃点={2,-6} 自动度量={3,-9} 状态={4,-12} DHCP={5}' -f $o.InterfaceAlias, $o.InterfaceIndex, $o.InterfaceMetric, $o.AutomaticMetric, $o.ConnectionState, $o.Dhcp)
            }
        } else { $L += '  (Get-NetIPInterface 不可用)' }
    } catch { $L += ('  错误: ' + $_.Exception.Message) }
    if (-not $got) { $L += '  (无数据：CIM 可能受限)' }

    $L += ''
    $L += '--- 2) Get-NetRoute 默认路由（按总跃点排序）---'
    try {
        if (Get-Command Get-NetRoute -ErrorAction SilentlyContinue) {
            $rs = @(Get-NetRoute -AddressFamily IPv4 -DestinationPrefix '0.0.0.0/0' -ErrorAction Stop)
            if ($rs.Count -eq 0) { $L += '  (无默认路由)' }
            foreach ($r in $rs) {
                $tot = $null
                try { $tot = [int]$r.InterfaceMetric + [int]$r.RouteMetric } catch { }
                $L += ('  {0,-16} 下一跳={1,-16} 接口跃点={2,-6} 路由跃点={3,-6} 总={4}' -f $r.InterfaceAlias, $r.NextHop, $r.InterfaceMetric, $r.RouteMetric, $tot)
            }
        } else { $L += '  (Get-NetRoute 不可用)' }
    } catch { $L += ('  错误: ' + $_.Exception.Message) }

    $L += ''
    $L += '--- 3) Find-NetRoute 8.8.8.8（原始返回，注意可能有多条）---'
    try {
        if (Get-Command Find-NetRoute -ErrorAction SilentlyContinue) {
            $n = 0
            foreach ($o in @(Find-NetRoute -RemoteIPAddress '8.8.8.8' -ErrorAction Stop)) {
                $n++
                $L += ('  [{0}] 类型={1} 接口={2} IP={3} 下一跳={4} 接口跃点={5} 路由跃点={6}' -f $n, $o.CimClass.CimClassName, $o.InterfaceAlias, $o.IPAddress, $o.NextHop, $o.InterfaceMetric, $o.RouteMetric)
            }
            if ($n -eq 0) { $L += '  (无返回)' }
        } else { $L += '  (Find-NetRoute 不可用)' }
    } catch { $L += ('  错误: ' + $_.Exception.Message) }

    $L += ''
    $L += '--- 4) route print -4 的默认路由行（顺序不可信，仅参考）---'
    foreach ($line in ((Invoke-Route 'print -4').Output -split "`r?`n")) {
        if ($line -match '^\s*0\.0\.0\.0\s+0\.0\.0\.0') { $L += ('  ' + $line.Trim()) }
    }
    $L += ''

    $L += '--- 5) netsh interface show interface ---'
    foreach ($line in ((Invoke-Netsh 'interface show interface').Output -split "`r?`n")) { $L += ('  ' + $line) }
    $L += '--- 6) netsh interface ipv4 show interfaces ---'
    foreach ($line in ((Invoke-Netsh 'interface ipv4 show interfaces').Output -split "`r?`n")) { $L += ('  ' + $line) }
    $L += ''
    $L += '--- 6b) netsh wlan show interfaces（原始输出，排查 WiFi 用）---'
    foreach ($line in ((Invoke-Netsh 'wlan show interfaces').Output -split "`r?`n")) { $L += ('  ' + $line) }
    $L += '--- 6c) 可见无线网络 ---'
    $vis = @(Get-VisibleWifiNetworks)
    if ($vis.Count -eq 0) { $L += '  (读不到：可能未开位置服务，或不在范围内)' }
    foreach ($n in $vis) { $L += ('  {0}  信号={1}  认证={2}' -f $n.SSID, $n.Signal, $n.Auth) }
    $L += '--- 6d) 无线服务 ---'
    try {
        $svc = Get-Service -Name 'WlanSvc' -ErrorAction Stop
        $L += ('  WlanSvc = {0} / {1}' -f $svc.Status, $svc.StartType)
    } catch { $L += ('  读取失败: ' + $_.Exception.Message) }
    $L += ''

    $L += '--- 7) 程序采集的网卡快照 ---'
    $snap = Get-NetworkSnapshot
    foreach ($a in @($snap.Adapters | Sort-Object Name)) {
        $L += ('  {0,-16} 类型={1,-4} 状态={2,-8} 管理={3,-8} 跃点={4,-8} 自动={5,-6} 已禁用={6,-6} IP={7} 网关={8}' -f $a.Name, $(if ($a.IsWifi) { '无线' } elseif ($a.IsVirtual) { '虚拟' } else { '有线' }), $a.StatusText, $a.AdminState, (Format-Metric $a.Metric), $a.MetricAuto, $a.Disabled, ($a.IpAddresses -join ','), ($a.Gateways -join ','))
    }
    $L += ''
    $L += '--- 8) 程序判定的当前出口 ---'
    if ($snap.Exit) {
        $L += ('  出口网卡 = {0}   网关 = {1}   生效跃点 = {2}   判定来源 = {3}' -f $snap.Exit.Name, $snap.Exit.Gateway, $snap.Exit.Metric, $snap.Exit.Source)
    } else { $L += '  (未判定出出口)' }
    if ($snap.ActualRoute) {
        $L += ('  最佳默认路由 = {0} 下一跳 {1} 总跃点 {2} 来源 {3}' -f $snap.ActualRoute.InterfaceAlias, $snap.ActualRoute.NextHop, $snap.ActualRoute.Total, $snap.ActualRoute.Source)
    }

    $path = Join-Path $script:ScriptDir '诊断报告.txt'
    try {
        Set-Content -LiteralPath $path -Value $L -Encoding UTF8
        Write-Log ('已生成诊断报告：' + $path)
        return $path
    } catch {
        Write-Log ('生成诊断报告失败：' + $_.Exception.Message)
        return ''
    }
}

# ---------------------------------------------------------------------
# 主界面
# ---------------------------------------------------------------------
function Update-Ui {
    if ($script:Busy) { return }
    if (-not $script:Grid) { return }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $snap = Get-NetworkSnapshot
        $script:LastSnapshot = $snap
        $exitName = ''
        if ($snap.Exit -and $snap.Exit.Name) { $exitName = $snap.Exit.Name }

        $script:Grid.SuspendLayout()
        try {
            $script:Grid.Rows.Clear()
            $sorted = @($snap.Adapters | Sort-Object -Property @{ Expression = { if ($_.Disabled) { 3 } elseif ($_.Name -eq $exitName) { 0 } elseif ($_.Status -eq 'Up') { 1 } else { 2 } } }, Name)
            foreach ($a in $sorted) {
                $kind = if ($a.IsWifi) { '无线' } elseif ($a.IsVirtual) { '虚拟' } else { '有线' }
                $isExit = ($a.Name -eq $exitName)
                $mark = if ($isExit) { '●' } else { '' }
                $speed = if ($a.SpeedMbps -gt 0) { "$($a.SpeedMbps) Mbps" } else { '-' }
                $ip = if ($a.IpAddresses.Count -gt 0) { $a.IpAddresses -join ', ' } else { '-' }
                $gw = if ($a.Gateways.Count -gt 0) { $a.Gateways -join ', ' } else { '-' }
                $dns = if ($a.Dns.Count -gt 0) { $a.Dns -join ', ' } else { '-' }
                $mac = if ($a.Mac) { $a.Mac } else { '-' }
                $admin = switch ($a.AdminEnabled) { $true { '已启用' } $false { '已禁用' } default { '未知' } }
                $idx = $script:Grid.Rows.Add($a.Name, $mark, $kind, $a.StatusText, (Format-MetricText $a), $speed, $ip, $a.SubnetMask, $gw, $dns, $mac, $admin)
                if ($isExit) { $script:Grid.Rows[$idx].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(206, 240, 206) }
                elseif ($a.Disabled -or $a.AdminEnabled -eq $false) { $script:Grid.Rows[$idx].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(238, 238, 238) }
                elseif ($a.Status -eq 'Up') { $script:Grid.Rows[$idx].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(232, 248, 232) }
            }
        } finally {
            # 必须放在 finally：一旦中间抛异常而没恢复布局，表格会永久停止重绘（看起来像假死）
            $script:Grid.ResumeLayout()
        }

        $script:TextBox.Text = Format-NetworkReport $snap
        $script:AdminLabel.Text = if ($script:IsAdmin) { '管理员模式：是' } else { '管理员模式：否（切换功能不可用）' }
        $script:AdminLabel.ForeColor = if ($script:IsAdmin) { [System.Drawing.Color]::FromArgb(0, 120, 0) } else { [System.Drawing.Color]::FromArgb(180, 60, 0) }
        $script:TimeLabel.Text = ('上次刷新：{0}（耗时 {1:N1}s）' -f (Get-Date).ToString('HH:mm:ss'), ($script:LastRefreshMs / 1000.0))
    } catch {
        Write-Log ('界面刷新失败：' + $_.Exception.Message)
        if ($script:StatusLabel) { $script:StatusLabel.Text = "刷新失败：$($_.Exception.Message)" }
    }
    $sw.Stop()
    $script:LastRefreshMs = [int]$sw.Elapsed.TotalMilliseconds
}

function Invoke-AutoFallbackCheck {
    <#
      ★ 自动回切看门狗 ★
      后台每刷新一轮检查一次：如果"无线出口已经不可用"（断开 / 没网关 / 没连上），
      而存在可用的有线网卡，就把流量出口自动切回有线，并把无线的跃点压下去
      （这样 WiFi 稍后自己重连也不会把流量抢回去）。
      触发条件很严格，避免误动：
        · 开关 autoFallback 打开（设置里可关）
        · 无线当前不可用
        · 有可用的有线网卡（在线 + 有默认网关）
        · 且此刻"无线仍然比有线优先"（出口是无线 / 没有出口 / 无线跃点更小）
        · 距上次自动回切至少 20 秒
    #>
    $cfg = Get-Config
    if (-not [bool]$cfg.autoFallback) { return }
    if ($script:Busy) { return }
    if (-not $script:IsAdmin) { return }
    $wifi = [string]$cfg.wifiAdapter
    if (-not $wifi) { return }
    $snap = $script:LastSnapshot
    if (-not $snap) { return }

    $wlAd = Get-AdapterByName $snap.Adapters $wifi
    if (-not $wlAd) { return }

    # 无线是否可用：网卡在线 + 有默认网关 + 无线状态是已连接
    $wlanUsable = ($wlAd.Status -eq 'Up') -and ($wlAd.Gateways.Count -gt 0)
    if ($snap.Wlan -and $snap.Wlan.Available -and (-not (Test-WlanConnected $snap.Wlan))) { $wlanUsable = $false }
    if ($wlanUsable) { return }

    $exitName = ''
    if ($snap.Exit -and $snap.Exit.Name) { $exitName = $snap.Exit.Name }

    # 找一块可用的有线网卡
    $cands = @()
    foreach ($n in @([string]$cfg.lastWired, [string]$cfg.localA, [string]$cfg.localB)) {
        if (-not $n -or $n -eq $wifi) { continue }
        $a = Get-AdapterByName $snap.Adapters $n
        if ($a -and ($cands -notcontains $a.Name)) { $cands += $a.Name }
    }
    foreach ($a in @(Get-LocalAdapters $snap.Adapters | Where-Object { $_.Status -eq 'Up' -and $_.Gateways.Count -gt 0 })) {
        if ($cands -notcontains $a.Name) { $cands += $a.Name }
    }
    $target = ''
    foreach ($n in $cands) {
        $a = Get-AdapterByName $snap.Adapters $n
        if ($a -and $a.Status -eq 'Up' -and $a.Gateways.Count -gt 0) { $target = $n; break }
    }

    # 冷却检查放在动作之前，避免在冷却期内白折腾
    if (((Get-Date) - $script:LastFallback).TotalSeconds -lt 20) { return }

    # 有线网卡可能被"切到 WiFi 就禁用本地网卡"选项禁用了：这里自动把它启用回来
    if (-not $target) {
        foreach ($n in $cands) {
            $a = Get-AdapterByName $snap.Adapters $n
            if (-not $a) { continue }
            if (($a.AdminEnabled -eq $false) -or $a.Disabled) {
                Set-Busy $true ("无线已断开，正在启用有线网卡 [{0}]…" -f $n)
                Write-Log ("[自动回切] 有线网卡 [{0}] 处于禁用状态，先启用它" -f $n)
                if (Set-AdapterAdminState -Name $n -Enabled $true) {
                    $null = Wait-AdapterAdminState -Name $n -Enabled $true -TimeoutSec 15
                    if (Wait-AdapterGateway -Name $n -TimeoutSec 15) { $target = $n; break }
                }
            }
        }
    }
    if (-not $target) { Set-Busy $false; return }

    # 无线还比有线优先才需要动作（已经是"有线优先"就别重复折腾）
    $wiredAd = Get-AdapterByName $snap.Adapters $target
    $wifiRank = if ($null -ne $wlAd.Metric) { [int]$wlAd.Metric } else { 999999 }
    $wiredRank = if ($wiredAd -and $null -ne $wiredAd.Metric) { [int]$wiredAd.Metric } else { 999999 }
    $needAction = ($exitName -eq $wifi) -or (-not $exitName) -or ($wifiRank -lt $wiredRank)
    if (-not $needAction) { return }

    # 动手前用新鲜快照复核一次：避免基于几秒前的旧状态误切
    $chk = Get-NetworkSnapshot
    $wlChk = Get-AdapterByName $chk.Adapters $wifi
    $wlanOk2 = ($wlChk -and $wlChk.Status -eq 'Up' -and $wlChk.Gateways.Count -gt 0)
    if ($chk.Wlan -and $chk.Wlan.Available -and (-not (Test-WlanConnected $chk.Wlan))) { $wlanOk2 = $false }
    if ($wlanOk2) {
        Write-Log '[自动回切] 复核时无线已恢复，取消本次回切'
        if ($script:Busy) { Set-Busy $false }
        return
    }
    $tgtChk = Get-AdapterByName $chk.Adapters $target
    if (-not $tgtChk -or $tgtChk.Status -ne 'Up' -or $tgtChk.Gateways.Count -eq 0) {
        Write-Log ('[自动回切] 复核时有线 [{0}] 不可用，取消本次回切' -f $target)
        if ($script:Busy) { Set-Busy $false }
        return
    }

    $script:LastFallback = Get-Date

    Write-Log ("[自动回切] 无线不可用（状态={0} 网关数={1}），自动把流量切回有线 [{2}]" -f $wlAd.Status, $wlAd.Gateways.Count, $target)
    Set-Busy $true '无线已断开，正在自动切回有线…'
    try {
        # 直接复用刚复核过的新鲜快照，不再多采一次
        $fresh = $chk
        $res = Set-ExitRoute -Target $target -Adapters $fresh.Adapters -PreferMetric ([int]$cfg.preferMetric) -OtherMetric ([int]$cfg.deprioritizeMetric)
        foreach ($m in @($res.Messages)) { Write-Log ('  ' + $m) }
        if ($res.Verified) {
            $c2 = Get-Config
            $c2.lastWired = $target
            $null = Save-Config $c2
            $msg = ('WiFi 已断开，已自动切回有线 [{0}]  {1}' -f $target, (Get-ExitSummaryText -Adapters $fresh.Adapters))
            if ($script:StatusLabel) { $script:StatusLabel.Text = $msg }
            Write-Log $msg
        } else {
            Write-Log ('[自动回切] 切回有线未生效：' + $target + ' 不是跃点最小的默认路由')
            if ($script:StatusLabel) { $script:StatusLabel.Text = ('WiFi 已断开，自动切回有线 [{0}] 未生效，详见 switch-log.txt' -f $target) }
        }
    } catch {
        Write-Log ('[自动回切] 异常：' + $_.Exception.Message)
    } finally {
        Set-Busy $false
        Update-Ui
    }
}

function Invoke-LocalFallbackCheck {
    <#
      ★ 本地网卡之间的自动切换 ★
      判定依据是"跃点谁更优先"，与当前出口是谁无关（出口会因路由消失自动漂移，
      所以不能拿出口当判据 —— 那正是"断开后跃点不变"的原因）：
        · 目标 = 当前出口（它可用时保持不动），否则配置里的 A/B，再否则任意可用本地网卡
        · 只要还有另一块本地网卡"跃点不比目标差"（含断开但跃点更低、或跃点打平），就动手：
          目标 -> preferMetric，其它本地网卡（含断开那块）-> deprioritizeMetric
      这样断开的那块不会在恢复后抢回流量，看门狗也不会反复触发。
    #>
    param($Snap)
    $cfg = Get-Config
    if (-not [bool]$cfg.autoFallback) { return }
    if ($script:Busy) { return }
    if (-not $script:IsAdmin) { return }
    if (-not $Snap) { return }

    $locals = @(Get-LocalAdapters $Snap.Adapters)
    if ($locals.Count -lt 2) { return }        # 只有一块本地网卡，没有"切到另一块"

    $exitName = ''
    if ($Snap.Exit -and $Snap.Exit.Name) { $exitName = $Snap.Exit.Name }
    $localNames = @($locals | ForEach-Object { $_.Name })
    # 出口不是本地网卡（比如还走在 WiFi 上）就不插手，交给 WiFi 那一路
    if ($exitName -and ($localNames -notcontains $exitName)) { return }

    $usable = @($locals | Where-Object { $_.Status -eq 'Up' -and $_.Gateways.Count -gt 0 })
    if ($usable.Count -eq 0) { return }        # 一块可用的都没有，无从下手

    # 如果本地网卡全都还是"自动度量"（例如用户刚点过「恢复自动跃点」），
    # 就不要去打破这个平局、把用户的选择又改回去
    $explicit = @($locals | Where-Object { $_.MetricAuto -eq $false })
    if ($explicit.Count -eq 0) { return }

    # 目标选择：保持当前出口优先，其次是配置的 A/B，最后是任意可用本地网卡
    $target = ''
    if ($exitName -and (@($usable | Where-Object { $_.Name -eq $exitName }).Count -gt 0)) { $target = $exitName }
    if (-not $target) {
        foreach ($n in @([string]$cfg.localA, [string]$cfg.localB)) {
            if ($n -and (@($usable | Where-Object { $_.Name -eq $n }).Count -gt 0)) { $target = $n; break }
        }
    }
    if (-not $target) { $target = $usable[0].Name }

    $tgtAd = Get-AdapterByName $Snap.Adapters $target
    $tgtRank = if ($tgtAd -and $null -ne $tgtAd.Metric) { [int]$tgtAd.Metric } else { 999999 }

    # 是否还有别的本地网卡"不比目标差"（打平也算）——包括已经断开但跃点仍然更低的那块
    $needAction = $false
    foreach ($a in $locals) {
        if ($a.Name -eq $target) { continue }
        $r = if ($null -ne $a.Metric) { [int]$a.Metric } else { 999999 }
        if ($r -le $tgtRank) { $needAction = $true; break }
    }
    if (-not $needAction) { return }
    if (((Get-Date) - $script:LastFallback).TotalSeconds -lt 20) { return }
    $script:LastFallback = Get-Date

    Write-Log ("[自动切换] 本地网卡之间竞争（目标=[{0}] 跃点={1}），自动把优先出口固定到 [{0}]" -f $target, $tgtRank)
    Set-Busy $true ("正在把本地出口固定到 [{0}]…" -f $target)
    try {
        # 其它本地网卡（含断开那块）先压跃点：避免"断了却仍占优先位"，也避免看门狗反复触发
        foreach ($a in $locals) {
            if ($a.Name -eq $target) { continue }
            if (Set-InterfaceMetric -Name $a.Name -Metric ([int]$cfg.deprioritizeMetric)) {
                Write-Log ('  接口跃点 [{0}] = {1}' -f $a.Name, (Get-InterfaceMetricOf $a.Name))
            } else {
                Write-Log ('  接口跃点设置失败 [{0}] {1}' -f $a.Name, $script:LastError)
            }
        }
        $fresh = Get-NetworkSnapshot
        $res = Set-ExitRoute -Target $target -Adapters $fresh.Adapters -PreferMetric ([int]$cfg.preferMetric) -OtherMetric ([int]$cfg.deprioritizeMetric)
        foreach ($m in @($res.Messages)) { Write-Log ('  ' + $m) }
        if ($res.Verified) {
            $msg = ('已把本地出口固定到 [{0}]  {1}' -f $target, (Get-ExitSummaryText -Adapters $fresh.Adapters))
            if ($script:StatusLabel) { $script:StatusLabel.Text = $msg }
            Write-Log $msg
        } else {
            Write-Log ('[自动切换] 固定到 [{0}] 未生效：不是跃点最小的默认路由' -f $target)
            if ($script:StatusLabel) { $script:StatusLabel.Text = ('把本地出口固定到 [{0}] 未生效，详见 switch-log.txt' -f $target) }
        }
    } catch {
        Write-Log ('[自动切换] 异常：' + $_.Exception.Message)
    } finally {
        Set-Busy $false
        Update-Ui
    }
}

function Apply-RefreshInterval {
    if (-not $script:Timer) { return }
    $sec = [int](Get-Config).autoRefreshSeconds
    if ($sec -le 0) { $script:Timer.Stop() }
    else {
        # 自适应：刷新耗时长得离谱时，保证两次刷新之间至少留 1 秒空闲，否则界面会被采集占满而假死
        $script:Timer.Interval = [int]([Math]::Max(($sec * 1000), ($script:LastRefreshMs + 1000)))
        $script:Timer.Start()
    }
}

function New-ActionButton {
    param($Panel, [string]$Text, [int]$Width, [scriptblock]$OnClick, [bool]$Track = $false)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text
    $b.Size = New-Object System.Drawing.Size($Width, 38)
    $b.Margin = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
    $b.Add_Click($OnClick)
    $Panel.Controls.Add($b)
    if ($Track) { $script:ActionButtons += $b }
    return $b
}

function Start-Gui {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    try { [System.Windows.Forms.Application]::EnableVisualStyles() } catch { }

    $script:UiFont = New-UiFont
    $script:IsAdmin = Test-IsAdmin
    $script:ActionButtons = @()

    $form = New-Object System.Windows.Forms.Form
    $script:MainForm = $form
    $form.Text = '网络连接查看与切换工具'
    $form.Size = New-Object System.Drawing.Size(1050, 730)
    $form.MinimumSize = New-Object System.Drawing.Size(880, 600)
    $form.StartPosition = 'CenterScreen'
    $form.Font = $script:UiFont

    $root = New-Object System.Windows.Forms.TableLayoutPanel
    $root.Dock = 'Fill'
    $root.ColumnCount = 1
    $root.RowCount = 5
    $null = $root.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $null = $root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 52)))
    $null = $root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 205)))
    $null = $root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $null = $root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 58)))
    $null = $root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 26)))
    $form.Controls.Add($root)

    # --- 第 1 行：标题栏
    $head = New-Object System.Windows.Forms.Panel
    $head.Dock = 'Fill'
    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = '网络连接查看与切换'
    $lblTitle.Font = New-Object System.Drawing.Font($script:UiFont.FontFamily, 14, [System.Drawing.FontStyle]::Bold)
    $lblTitle.Location = New-Object System.Drawing.Point(12, 8)
    $lblTitle.Size = New-Object System.Drawing.Size(330, 34)
    $head.Controls.Add($lblTitle)

    $script:AdminLabel = New-Object System.Windows.Forms.Label
    $script:AdminLabel.Location = New-Object System.Drawing.Point(350, 16)
    $script:AdminLabel.Size = New-Object System.Drawing.Size(340, 22)
    $head.Controls.Add($script:AdminLabel)

    $script:TimeLabel = New-Object System.Windows.Forms.Label
    $script:TimeLabel.Location = New-Object System.Drawing.Point(750, 16)
    $script:TimeLabel.Size = New-Object System.Drawing.Size(270, 22)
    $script:TimeLabel.TextAlign = 'MiddleRight'
    $script:TimeLabel.Anchor = 'Top,Right'
    $head.Controls.Add($script:TimeLabel)
    $root.Controls.Add($head, 0, 0)

    # --- 第 2 行：适配器列表
    $grpGrid = New-Object System.Windows.Forms.GroupBox
    $grpGrid.Text = '网络适配器（● = 当前流量出口）'
    $grpGrid.Dock = 'Fill'
    $grpGrid.Padding = New-Object System.Windows.Forms.Padding(8)
    $grid = New-Object System.Windows.Forms.DataGridView
    $script:Grid = $grid
    $grid.Dock = 'Fill'
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AllowUserToResizeRows = $false
    $grid.RowHeadersVisible = $false
    $grid.SelectionMode = 'FullRowSelect'
    $grid.MultiSelect = $false
    $grid.AutoSizeColumnsMode = 'Fill'
    $grid.BackgroundColor = [System.Drawing.Color]::White
    $grid.BorderStyle = 'FixedSingle'
    $grid.EnableHeadersVisualStyles = $false
    $grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    $colSpecs = @(
        @('名称', 80), @('出口', 32), @('类型', 40), @('状态', 60), @('跃点数', 46), @('速率', 58),
        @('IPv4 地址', 95), @('子网掩码', 85), @('默认网关', 85),
        @('DNS 服务器', 85), @('MAC 地址', 100), @('管理状态', 52)
    )
    foreach ($c in $colSpecs) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.HeaderText = $c[0]
        $col.FillWeight = [single]$c[1]
        $col.SortMode = 'NotSortable'
        $null = $grid.Columns.Add($col)
    }
    $grpGrid.Controls.Add($grid)
    $root.Controls.Add($grpGrid, 0, 1)

    # --- 第 3 行：详情
    $grpDetail = New-Object System.Windows.Forms.GroupBox
    $grpDetail.Text = '当前连接详情'
    $grpDetail.Dock = 'Fill'
    $grpDetail.Padding = New-Object System.Windows.Forms.Padding(8)
    $txt = New-Object System.Windows.Forms.TextBox
    $script:TextBox = $txt
    $txt.Dock = 'Fill'
    $txt.Multiline = $true
    $txt.ReadOnly = $true
    $txt.ScrollBars = 'Both'
    $txt.WordWrap = $false
    $txt.BackColor = [System.Drawing.Color]::White
    try { $txt.Font = New-Object System.Drawing.Font('Consolas', 9.5) } catch { }
    $grpDetail.Controls.Add($txt)
    $root.Controls.Add($grpDetail, 0, 2)

    # --- 第 4 行：按钮
    $panelBtn = New-Object System.Windows.Forms.FlowLayoutPanel
    $script:ButtonPanel = $panelBtn
    $panelBtn.Dock = 'Fill'
    $panelBtn.FlowDirection = 'RightToLeft'
    $panelBtn.WrapContents = $false
    $panelBtn.Padding = New-Object System.Windows.Forms.Padding(8)
    $root.Controls.Add($panelBtn, 0, 3)

    $null = New-ActionButton -Panel $panelBtn -Text '设置…' -Width 90 -Track $true -OnClick { Show-SettingsDialog }
    $null = New-ActionButton -Panel $panelBtn -Text '诊断' -Width 70 -OnClick {
        $dp = Write-Diagnostics
        if ($dp) {
            try { Start-Process notepad.exe -ArgumentList ('"' + $dp + '"') } catch { }
            if ($script:StatusLabel) { $script:StatusLabel.Text = ('诊断报告：' + $dp) }
        } else {
            Show-Msg '生成诊断报告失败，请检查程序目录是否可写。' '诊断' 'Error'
        }
    }
    $null = New-ActionButton -Panel $panelBtn -Text '打开日志' -Width 90 -OnClick { try { Start-Process notepad.exe -ArgumentList ('"' + $script:LogPath + '"') } catch { } }
    $null = New-ActionButton -Panel $panelBtn -Text '刷新信息' -Width 100 -OnClick { Update-Ui; if ($script:StatusLabel) { $script:StatusLabel.Text = '信息已刷新' } }
    $btnWifi = New-ActionButton -Panel $panelBtn -Text '有线 ⇄ WiFi' -Width 150 -Track $true -OnClick { Invoke-WifiToggle }
    $btnLocal = New-ActionButton -Panel $panelBtn -Text '本地网络切换' -Width 150 -Track $true -OnClick { Invoke-LocalSwitch }
    $btnLocal.Font = New-Object System.Drawing.Font($script:UiFont, [System.Drawing.FontStyle]::Bold)
    $btnWifi.Font = New-Object System.Drawing.Font($script:UiFont, [System.Drawing.FontStyle]::Bold)

    # --- 第 5 行：状态栏
    $script:StatusLabel = New-Object System.Windows.Forms.Label
    $script:StatusLabel.Dock = 'Fill'
    $script:StatusLabel.TextAlign = 'MiddleLeft'
    $script:StatusLabel.Padding = New-Object System.Windows.Forms.Padding(10, 0, 0, 0)
    $script:StatusLabel.Text = '就绪'
    $root.Controls.Add($script:StatusLabel, 0, 4)

    $tip = New-Object System.Windows.Forms.ToolTip
    $tip.SetToolTip($btnLocal, '把实际流量出口在「本地网络 A」和「本地网络 B」之间来回切换：只调整接口跃点数，两块网卡都保持启用，不断网')
    $tip.SetToolTip($btnWifi, '有线 ⇄ WiFi 双向切换：当前走有线就切到 WiFi（自动连接）；当前走 WiFi 就切回有线。默认都不断开任何网卡')

    # --- 定时刷新
    $script:Timer = New-Object System.Windows.Forms.Timer
    $script:Timer.Interval = 3000
    $script:Timer.Add_Tick({
        if ($script:Busy) { return }
        Update-Ui
        if (-not $script:Busy) { Invoke-AutoFallbackCheck }
        if (-not $script:Busy) { Invoke-LocalFallbackCheck -Snap $script:LastSnapshot }
        # 自适应刷新间隔：采集耗时 + 1 秒空闲，避免 UI 线程被采集占满
        $sec = [int](Get-Config).autoRefreshSeconds
        if ($sec -gt 0) {
            $script:Timer.Interval = [int]([Math]::Max(($sec * 1000), ($script:LastRefreshMs + 1000)))
        }
    })

    $form.Add_Shown({
        $script:StatusLabel.Text = '正在读取网络信息…'
        try { [System.Windows.Forms.Application]::DoEvents() } catch { }
        Update-Ui
        Apply-RefreshInterval
        if (-not $script:IsAdmin) { $script:StatusLabel.Text = '当前为普通权限：可以查看信息，切换流量出口需要管理员权限' }
        else { $script:StatusLabel.Text = '就绪' }
    })
    $form.Add_FormClosing({ try { $script:Timer.Stop() } catch { } })

    Write-Log '界面已启动'
    [void]$form.ShowDialog()
    $form.Dispose()
}

# ---------------------------------------------------------------------
# Probe 模式：只输出信息，用于排查
# ---------------------------------------------------------------------
function Invoke-Probe {
    $snap = Get-NetworkSnapshot
    Write-Output (Format-NetworkReport $snap)
    Write-Output ''
    Write-Output '--- 配置与权限 ---'
    $cfg = Get-Config
    Write-Output ("管理员权限 : {0}" -f (Test-IsAdmin))
    Write-Output ("配置文件   : {0}" -f $script:ConfigPath)
    Write-Output ("本地网络 A : {0}" -f $cfg.localA)
    Write-Output ("本地网络 B : {0}" -f $cfg.localB)
    Write-Output ("无线网卡   : {0}" -f $cfg.wifiAdapter)
    Write-Output ("WiFi 名称  : {0}" -f $cfg.wifiSsid)
    Write-Output ("跃点数     : 首选={0} 其它={1}" -f $cfg.preferMetric, $cfg.deprioritizeMetric)
    Write-Output ''
    Write-Output '--- 接口跃点数明细（自动度量 = 会被 Windows 覆盖手动值）---'
    $mm = Get-InterfaceMetricMap
    foreach ($k in @($mm.Keys | Sort-Object)) {
        Write-Output ("  {0,-16} 跃点={1,-6} 自动度量={2}" -f $k, $mm[$k].Metric, $mm[$k].Auto)
    }
    Write-Output '--- Windows 实际会走的路由（最优默认路由）---'
    $act = Get-ActualExitRoute $snap
    if ($act) {
        Write-Output ("  网卡={0}  下一跳={1}  接口跃点={2}  路由跃点={3}  总跃点={4}  来源={5}" -f $act.InterfaceAlias, $act.NextHop, $act.InterfaceMetric, $act.RouteMetric, $act.Total, $act.Source)
    } else {
        Write-Output '  （Find-NetRoute 不可用或查询失败，已回退到路由表推断）'
    }
    Write-Output '--- netsh interface show interface ---'
    Write-Output (Invoke-Netsh 'interface show interface').Output
    Write-Output '--- netsh interface ipv4 show interfaces ---'
    Write-Output (Invoke-Netsh 'interface ipv4 show interfaces').Output
    Write-Output ''
    $dp = Write-Diagnostics
    if ($dp) { Write-Output ('诊断报告已写入：' + $dp) } else { Write-Output '诊断报告写入失败' }
}

# ---------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------
if ($Probe) {
    Invoke-Probe
    return
}

# 首次运行时根据系统实际情况给出默认配置
if (-not (Test-Path -LiteralPath $script:ConfigPath)) {
    try {
        $snapAdapters = @(Get-AdapterSnapshot)
        $loc = @($snapAdapters | Where-Object { -not $_.IsWifi -and -not $_.IsVirtual -and -not $_.Disabled })
        $wf = @($snapAdapters | Where-Object { $_.IsWifi -and -not $_.IsVirtual })
        $cfg = Get-Config
        if ($loc.Count -ge 1) { $cfg.localA = $loc[0].Name }
        if ($loc.Count -ge 2) { $cfg.localB = $loc[1].Name }
        if ($wf.Count -ge 1) { $cfg.wifiAdapter = $wf[0].Name }
        $script:ConfigCache = $cfg
        $null = Save-Config $cfg
    } catch { }
}

# 启动时提示提升权限
if (-not $NoElevate -and -not (Test-IsAdmin)) {
    Add-Type -AssemblyName System.Windows.Forms
    $r = [System.Windows.Forms.MessageBox]::Show(
        "本工具切换流量出口、读取无线信息需要管理员权限。`r`n`r`n是否以管理员身份重新启动？`r`n（选择「否」可以只用普通权限查看连接信息）",
        '网络切换工具', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($r -eq [System.Windows.Forms.DialogResult]::Yes) {
        Restart-Elevated
        exit 0
    }
}

Start-Gui
