<#
    名称: IP 地理监控 (IpGeoMonitor)
    作用: 在桌面浮起一条精简的通知条, 实时显示本机公网 IP 与其所在国家/地区。
    交互: 左键单击立即刷新; 不操作则按设定间隔(默认 60 秒)自动刷新; 右键弹出菜单与设置。
    运行: Windows 10/11 自带 PowerShell 5.1 + .NET Framework, 无需安装任何依赖。
    作者: DSH

    参数:
      -SelfTest            只做一次联网查询并把结果打到控制台(不显示界面), 用于排错
      -PreviewPath <dir>   把主界面与设置窗口渲染成 PNG 预览图后退出
      -ResetConfig         删除本地配置并退出
#>
[CmdletBinding()]
param(
    [switch]$SelfTest,
    [switch]$SmokeTest,
    [string]$PreviewPath,
    [switch]$ResetConfig,
    [string]$DataDir,
    [string]$Source
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# 0. 基础常量 / 路径
# ---------------------------------------------------------------------------
$script:AppName  = 'IP 地理监控'
$script:AppVer   = '1.0.0'
$script:AppDir   = Split-Path -Parent $MyInvocation.MyCommand.Path

function Test-DirWritable {
    param([string]$Dir)
    try {
        if (-not (Test-Path $Dir)) { New-Item -ItemType Directory -Path $Dir -Force -ErrorAction Stop | Out-Null }
        $probe = Join-Path $Dir '.write-probe'
        Set-Content -Path $probe -Value 'ok' -Encoding ASCII -ErrorAction Stop
        Remove-Item $probe -Force -ErrorAction SilentlyContinue
        return $true
    } catch { return $false }
}

$script:DataDir = ''
foreach ($candidate in @($DataDir, (Join-Path $env:APPDATA 'IpGeoMonitor'), (Join-Path $script:AppDir 'data'), (Join-Path $env:TEMP 'IpGeoMonitor'))) {
    if (-not $candidate) { continue }
    if (Test-DirWritable -Dir $candidate) { $script:DataDir = $candidate; break }
}
if (-not $script:DataDir) { $script:DataDir = $env:TEMP }

$script:CfgPath  = Join-Path $script:DataDir 'config.json'
$script:CachePath= Join-Path $script:DataDir 'last.json'
$script:LogPath  = Join-Path $script:DataDir 'monitor.log'
$script:MutexName= 'Local\IpGeoMonitor_SingleInstance'

function Write-Log {
    param([string]$Message)
    try {
        $line = '{0}  {1}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $Message
        Add-Content -Path $script:LogPath -Value $line -Encoding UTF8
        $fi = Get-Item $script:LogPath -ErrorAction SilentlyContinue
        if ($fi -and $fi.Length -gt 256KB) {
            $tail = Get-Content $script:LogPath -Tail 300 -Encoding UTF8
            Set-Content -Path $script:LogPath -Value $tail -Encoding UTF8
        }
    } catch { }
}

# 事件处理器里抛出的异常会被 PowerShell 静默吞掉, 统一在这里捕获并记日志
function Invoke-Guarded {
    param([string]$Name, [scriptblock]$Body)
    try { & $Body }
    catch {
        if (-not $script:GuardErrors) { $script:GuardErrors = @{} }
        $n = 0
        if ($script:GuardErrors.ContainsKey($Name)) { $n = [int]$script:GuardErrors[$Name] }
        $script:GuardErrors[$Name] = $n + 1
        if ($n -lt 3) {
            $pos = ''
            try { $pos = ($_.InvocationInfo.PositionMessage -replace "`r?`n", ' ') } catch { }
            Write-Log ("[$Name] 异常(第{0}次): {1} :: {2}" -f ($n + 1), $_.Exception.Message, $pos)
        }
    }
}

if ($ResetConfig) {
    Remove-Item $script:CfgPath -Force -ErrorAction SilentlyContinue
    Remove-Item $script:CachePath -Force -ErrorAction SilentlyContinue
    Write-Host '配置已重置。' -ForegroundColor Green
    exit 0
}

# ---------------------------------------------------------------------------
# 1. 国家/地区代码 -> 中文名 映射表
# ---------------------------------------------------------------------------
$script:IsoText = @'
AD=安道尔;AE=阿联酋;AF=阿富汗;AG=安提瓜和巴布达;AI=安圭拉;AL=阿尔巴尼亚;AM=亚美尼亚;AO=安哥拉;AQ=南极洲;AR=阿根廷;AS=美属萨摩亚;AT=奥地利;AU=澳大利亚;AW=阿鲁巴;AX=奥兰群岛;AZ=阿塞拜疆;
BA=波黑;BB=巴巴多斯;BD=孟加拉国;BE=比利时;BF=布基纳法索;BG=保加利亚;BH=巴林;BI=布隆迪;BJ=贝宁;BL=圣巴泰勒米;BM=百慕大;BN=文莱;BO=玻利维亚;BQ=荷兰加勒比区;BR=巴西;BS=巴哈马;BT=不丹;BV=布韦岛;BW=博茨瓦纳;BY=白俄罗斯;BZ=伯利兹;
CA=加拿大;CC=科科斯群岛;CD=刚果(金);CF=中非;CG=刚果(布);CH=瑞士;CI=科特迪瓦;CK=库克群岛;CL=智利;CM=喀麦隆;CN=中国;CO=哥伦比亚;CR=哥斯达黎加;CU=古巴;CV=佛得角;CW=库拉索;CX=圣诞岛;CY=塞浦路斯;CZ=捷克;
DE=德国;DJ=吉布提;DK=丹麦;DM=多米尼克;DO=多米尼加;DZ=阿尔及利亚;
EC=厄瓜多尔;EE=爱沙尼亚;EG=埃及;EH=西撒哈拉;ER=厄立特里亚;ES=西班牙;ET=埃塞俄比亚;
FI=芬兰;FJ=斐济;FK=福克兰群岛;FM=密克罗尼西亚;FO=法罗群岛;FR=法国;
GA=加蓬;GB=英国;GD=格林纳达;GE=格鲁吉亚;GF=法属圭亚那;GG=根西岛;GH=加纳;GI=直布罗陀;GL=格陵兰;GM=冈比亚;GN=几内亚;GP=瓜德罗普;GQ=赤道几内亚;GR=希腊;GS=南乔治亚岛;GT=危地马拉;GU=关岛;GW=几内亚比绍;GY=圭亚那;
HK=中国香港;HM=赫德岛;HN=洪都拉斯;HR=克罗地亚;HT=海地;HU=匈牙利;
ID=印度尼西亚;IE=爱尔兰;IL=以色列;IM=马恩岛;IN=印度;IO=英属印度洋领地;IQ=伊拉克;IR=伊朗;IS=冰岛;IT=意大利;
JE=泽西岛;JM=牙买加;JO=约旦;JP=日本;
KE=肯尼亚;KG=吉尔吉斯斯坦;KH=柬埔寨;KI=基里巴斯;KM=科摩罗;KN=圣基茨和尼维斯;KP=朝鲜;KR=韩国;KW=科威特;KY=开曼群岛;KZ=哈萨克斯坦;
LA=老挝;LB=黎巴嫩;LC=圣卢西亚;LI=列支敦士登;LK=斯里兰卡;LR=利比里亚;LS=莱索托;LT=立陶宛;LU=卢森堡;LV=拉脱维亚;LY=利比亚;
MA=摩洛哥;MC=摩纳哥;MD=摩尔多瓦;ME=黑山;MF=法属圣马丁;MG=马达加斯加;MH=马绍尔群岛;MK=北马其顿;ML=马里;MM=缅甸;MN=蒙古;MO=中国澳门;MP=北马里亚纳群岛;MQ=马提尼克;MR=毛里塔尼亚;MS=蒙特塞拉特;MT=马耳他;MU=毛里求斯;MV=马尔代夫;MW=马拉维;MX=墨西哥;MY=马来西亚;MZ=莫桑比克;
NA=纳米比亚;NC=新喀里多尼亚;NE=尼日尔;NF=诺福克岛;NG=尼日利亚;NI=尼加拉瓜;NL=荷兰;NO=挪威;NP=尼泊尔;NR=瑙鲁;NU=纽埃;NZ=新西兰;
OM=阿曼;
PA=巴拿马;PE=秘鲁;PF=法属波利尼西亚;PG=巴布亚新几内亚;PH=菲律宾;PK=巴基斯坦;PL=波兰;PM=圣皮埃尔和密克隆;PN=皮特凯恩群岛;PR=波多黎各;PS=巴勒斯坦;PT=葡萄牙;PW=帕劳;PY=巴拉圭;
QA=卡塔尔;
RE=留尼汪;RO=罗马尼亚;RS=塞尔维亚;RU=俄罗斯;RW=卢旺达;
SA=沙特阿拉伯;SB=所罗门群岛;SC=塞舌尔;SD=苏丹;SE=瑞典;SG=新加坡;SH=圣赫勒拿;SI=斯洛文尼亚;SJ=斯瓦尔巴和扬马延;SK=斯洛伐克;SL=塞拉利昂;SM=圣马力诺;SN=塞内加尔;SO=索马里;SR=苏里南;SS=南苏丹;ST=圣多美和普林西比;SV=萨尔瓦多;SX=荷属圣马丁;SY=叙利亚;SZ=斯威士兰;
TC=特克斯和凯科斯群岛;TD=乍得;TF=法属南方领地;TG=多哥;TH=泰国;TJ=塔吉克斯坦;TK=托克劳;TL=东帝汶;TM=土库曼斯坦;TN=突尼斯;TO=汤加;TR=土耳其;TT=特立尼达和多巴哥;TV=图瓦卢;TW=中国台湾;TZ=坦桑尼亚;
UA=乌克兰;UG=乌干达;UM=美国本土外小岛屿;US=美国;UY=乌拉圭;UZ=乌兹别克斯坦;
VA=梵蒂冈;VC=圣文森特和格林纳丁斯;VE=委内瑞拉;VG=英属维尔京群岛;VI=美属维尔京群岛;VN=越南;VU=瓦努阿图;
WF=瓦利斯和富图纳;WS=萨摩亚;
YE=也门;YT=马约特;
ZA=南非;ZM=赞比亚;ZW=津巴布韦;
EU=欧盟;AP=亚太地区;A1=匿名代理;A2=卫星链路;O1=其他;T1=Tor 网络
'@

function Convert-IsoText {
    param([string]$Text)
    $map = @{}
    foreach ($pair in ($Text -split ';')) {
        $p = $pair.Trim()
        if (-not $p) { continue }
        $kv = $p -split '=', 2
        if ($kv.Count -eq 2) { $map[$kv[0].Trim().ToUpper()] = $kv[1].Trim() }
    }
    return $map
}
$script:IsoMap = Convert-IsoText $script:IsoText

function Get-CountryName {
    param([string]$Code, [hashtable]$Map, [string]$Lang = 'zh')
    if (-not $Code) { return '' }
    $Code = $Code.ToUpper()
    if ($Lang -eq 'zh' -and $Map.ContainsKey($Code)) { return $Map[$Code] }
    return $Code
}

function Get-CodeFromName {
    param([string]$Name, [hashtable]$Map)
    if (-not $Name) { return '' }
    foreach ($k in $Map.Keys) { if ($Map[$k] -eq $Name) { return $k } }
    return ''
}

# ---------------------------------------------------------------------------
# 2. 数据源定义
#    kind : json = 返回 JSON ; text-ip = 只返回纯 IP 文本(需要二次查询归属地)
#    ok   : 判定成功的字段路径, okValue 为期望值
#    |n   : 字段值按制表符切分后取第 n 段
# ---------------------------------------------------------------------------
$script:Endpoints = @(
    @{ id = 'ip-api'; name = 'ip-api.com'; kind = 'json'
       url = 'http://ip-api.com/json/?lang={lang}&fields=status,message,country,countryCode,regionName,city,isp,org,as,timezone,query'
       ok = 'status'; okValue = 'success'
       ip = 'query'; country = 'country'; cc = 'countryCode'; region = 'regionName'
       city = 'city'; isp = 'isp'; org = 'org'; tz = 'timezone'
       geoUrl = 'http://ip-api.com/json/{ip}?lang={lang}&fields=status,message,country,countryCode,regionName,city,isp,org,as,timezone,query' }

    @{ id = 'ipwho'; name = 'ipwho.is'; kind = 'json'
       url = 'http://ipwho.is/'
       ok = 'success'; okValue = 'True'
       ip = 'ip'; country = 'country'; cc = 'country_code'; region = 'region'
       city = 'city'; isp = 'connection.isp'; org = 'connection.org'; tz = 'timezone.id'
       geoUrl = 'http://ipwho.is/{ip}' }

    @{ id = 'ipinfo'; name = 'ipinfo.io'; kind = 'json'
       url = 'http://ipinfo.io/json'
       ip = 'ip'; country = 'country'; cc = 'country'; region = 'region'
       city = 'city'; isp = 'org'; org = 'org'; tz = 'timezone'
       geoUrl = 'http://ipinfo.io/{ip}/json' }

    @{ id = 'zxinc'; name = 'zxinc.org'; kind = 'json'
       url = 'http://ip.zxinc.org/api.php?type=json'
       ip = 'data.myip'; country = 'data.country|0'; cc = ''; region = 'data.country|1'
       city = 'data.location|1'; isp = 'data.local|0'; org = 'data.local|1'; tz = ''
       geoUrl = 'http://ip.zxinc.org/api.php?type=json&ip={ip}' }

    @{ id = 'ipsb'; name = 'ip.sb'; kind = 'text-ip'
       url = 'http://ip.sb'
       geoUrl = 'http://ip-api.com/json/{ip}?lang={lang}&fields=status,message,country,countryCode,regionName,city,isp,org,as,timezone,query' }

    @{ id = '3322'; name = 'ip.3322.net'; kind = 'text-ip'
       url = 'http://ip.3322.net/'
       geoUrl = 'http://ip-api.com/json/{ip}?lang={lang}&fields=status,message,country,countryCode,regionName,city,isp,org,as,timezone,query' }
)

# ---------------------------------------------------------------------------
# 3. 配置
# ---------------------------------------------------------------------------
function Get-DefaultConfig {
    return @{
        intervalSeconds   = 60
        timeoutSeconds    = 8
        refreshOnStart    = $true
        autoFallbackProxy = $true
        source            = 'auto'
        customUrl         = ''
        proxyMode         = 'system'
        proxyAddress      = ''
        showCountry       = $true
        showIp            = $true
        showCity          = $false
        showIsp           = $false
        badgeMode         = 'badge'
        showCountdown     = $true
        showUpdatedAt     = $false
        theme             = 'dark'
        countryLanguage   = 'zh'
        alwaysOnTop       = $true
        locked            = $false
        snapToEdge        = $false
        rememberPosition  = $true
        autoStart         = $false
        mouseThrough      = $false
        opacity           = 0.95
        fontScale         = 1.0
        size              = 'medium'
        doubleClickAction = 'refresh'
        posX              = -1
        posY              = -1
    }
}

function Import-Config {
    $cfg = Get-DefaultConfig
    if (Test-Path $script:CfgPath) {
        try {
            $raw = Get-Content -Path $script:CfgPath -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($k in @($cfg.Keys)) {
                $prop = $raw.PSObject.Properties[$k]
                if ($prop -and $null -ne $prop.Value) { $cfg[$k] = $prop.Value }
            }
        } catch { Write-Log "配置读取失败: $($_.Exception.Message)" }
    }
    # 兜底取值范围
    $cfg.intervalSeconds = [int][Math]::Min(3600, [Math]::Max(10, [int]$cfg.intervalSeconds))
    $cfg.timeoutSeconds  = [int][Math]::Min(60,   [Math]::Max(3,  [int]$cfg.timeoutSeconds))
    $cfg.opacity         = [double][Math]::Min(1.0, [Math]::Max(0.30, [double]$cfg.opacity))
    $cfg.fontScale       = [double][Math]::Min(1.6, [Math]::Max(0.8,  [double]$cfg.fontScale))
    return $cfg
}

function Save-Config {
    try {
        $script:Cfg | ConvertTo-Json -Depth 5 | Set-Content -Path $script:CfgPath -Encoding UTF8
    } catch { Write-Log "配置保存失败: $($_.Exception.Message)" }
}

function Save-Cache {
    try {
        @{ ip = $script:Result.Ip; cc = $script:Result.Cc; country = $script:Result.Country
           region = $script:Result.Region; city = $script:Result.City; isp = $script:Result.Isp
           tz = $script:Result.Tz; source = $script:Result.Source; at = $script:Result.At } |
            ConvertTo-Json -Depth 3 | Set-Content -Path $script:CachePath -Encoding UTF8
    } catch { }
}

function Load-Cache {
    if (-not (Test-Path $script:CachePath)) { return $null }
    try {
        $c = Get-Content -Path $script:CachePath -Raw -Encoding UTF8 | ConvertFrom-Json
        return @{ Ok = $true; Ip = $c.ip; Cc = $c.cc; Country = $c.country; Region = $c.region
                  City = $c.city; Isp = $c.isp; Tz = $c.tz; Source = $c.source; At = $c.at; Cached = $true }
    } catch { return $null }
}

$script:Cfg = Import-Config

# ---------------------------------------------------------------------------
# 4. 网络查询(纯函数, 会被复制到后台运行空间执行)
# ---------------------------------------------------------------------------
function Get-Prop {
    param($Object, [string]$Path, $Default = '')
    if (-not $Path -or $null -eq $Object) { return $Default }
    $cur = $Object
    foreach ($part in $Path.Split('.')) {
        if ($null -eq $cur) { return $Default }
        if ($cur -is [System.Collections.IDictionary]) {
            if ($cur.Contains($part)) { $cur = $cur[$part] } else { return $Default }
        } else {
            $p = $cur.PSObject.Properties[$part]
            if ($p) { $cur = $p.Value } else { return $Default }
        }
    }
    if ($null -eq $cur) { return $Default }
    return $cur
}

function Get-PropSplit {
    param($Object, [string]$Path)
    $idx = -1
    if ($Path -match '\|(-?\d+)$') { $idx = [int]$Matches[1]; $Path = $Path -replace '\|(-?\d+)$', '' }
    $v = Get-Prop -Object $Object -Path $Path -Default ''
    if ($idx -ge 0 -and $v -is [string]) {
        $parts = $v -split "`t"
        if ($idx -lt $parts.Count) { return $parts[$idx].Trim() }
        return ''
    }
    return ([string]$v).Trim()
}

function Invoke-RawRequest {
    param([string]$Url, [int]$Timeout, [string[]]$Modes, [string]$ProxyAddress)
    $errors = @()
    foreach ($m in $Modes) {
        $prev = $null
        try { $prev = [System.Net.WebRequest]::DefaultWebProxy } catch { }
        try {
            switch ($m) {
                'direct' { [System.Net.WebRequest]::DefaultWebProxy = $null }
                'custom' {
                    if ($ProxyAddress) {
                        [System.Net.WebRequest]::DefaultWebProxy = New-Object System.Net.WebProxy($ProxyAddress, $true)
                    } else {
                        [System.Net.WebRequest]::DefaultWebProxy = [System.Net.WebRequest]::GetSystemWebProxy()
                    }
                }
                default  { [System.Net.WebRequest]::DefaultWebProxy = [System.Net.WebRequest]::GetSystemWebProxy() }
            }
            $resp = Invoke-WebRequest -Uri $Url -TimeoutSec $Timeout -UseBasicParsing `
                        -Headers @{ 'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) IpGeoMonitor/1.0'
                                    'Accept' = 'application/json, text/plain, */*'
                                    'Cache-Control' = 'no-cache' } -ErrorAction Stop
            if ($resp.StatusCode -ge 200 -and $resp.StatusCode -lt 300) { return [string]$resp.Content }
            $errors += "$m HTTP $($resp.StatusCode)"
        } catch {
            $errors += "$m $($_.Exception.Message)"
        } finally {
            try { [System.Net.WebRequest]::DefaultWebProxy = $prev } catch { }
        }
    }
    throw ('请求失败: ' + ($errors -join ' ; '))
}

function Resolve-CountryName {
    param([string]$Code, [string]$Raw, [string]$Lang, [hashtable]$Map)
    $hasCjk = $false
    if ($Raw) { $hasCjk = [bool]($Raw -match '[\u4e00-\u9fff]') }
    if ($Lang -eq 'zh') {
        if ($hasCjk) { return $Raw }
        if ($Code -and $Map.ContainsKey($Code.ToUpper())) { return $Map[$Code.ToUpper()] }
        if ($Raw) { return $Raw }
        return $Code
    }
    if ($hasCjk) { if ($Code) { return $Code.ToUpper() } else { return $Raw } }
    if ($Raw) { return $Raw }
    return $Code
}

function ConvertFrom-Endpoint {
    param($Ep, [string]$Content, [hashtable]$Map, [string]$Lang)
    $fail = @{ ok = $false }
    if ($Ep.kind -eq 'text-ip') {
        $t = ($Content -replace '\s+', ' ').Trim()
        if ($t -match '([0-9]{1,3}(?:\.[0-9]{1,3}){3})') { return @{ ok = $true; ip = $Matches[1]; cc = ''; country = '' } }
        if ($t -match '([0-9a-fA-F]{0,4}(?::[0-9a-fA-F]{0,4}){3,7})') { return @{ ok = $true; ip = $Matches[1]; cc = ''; country = '' } }
        return $fail
    }
    $j = $null
    try { $j = $Content | ConvertFrom-Json } catch { return $fail }
    if ($Ep.ok) {
        $v = Get-Prop -Object $j -Path $Ep.ok -Default ''
        if ("$v" -ne "$($Ep.okValue)") { return $fail }
    }
    $ip = [string](Get-Prop -Object $j -Path $Ep.ip -Default '')
    if (-not $ip) { return $fail }
    $cc      = [string](Get-PropSplit -Object $j -Path $Ep.cc)
    $country = [string](Get-PropSplit -Object $j -Path $Ep.country)
    $region  = [string](Get-PropSplit -Object $j -Path $Ep.region)
    $city    = [string](Get-PropSplit -Object $j -Path $Ep.city)
    $isp     = [string](Get-PropSplit -Object $j -Path $Ep.isp)
    $org     = [string](Get-PropSplit -Object $j -Path $Ep.org)
    $tz      = [string](Get-PropSplit -Object $j -Path $Ep.tz)

    if (-not $cc -and $country -match '^[A-Za-z]{2}$') { $cc = $country.ToUpper(); $country = '' }
    if (-not $cc -and $country) { $cc = Get-CodeFromName -Name $country -Map $Map }
    $disp = Resolve-CountryName -Code $cc -Raw $country -Lang $Lang -Map $Map
    return @{ ok = $true; ip = $ip; cc = $cc.ToUpper(); country = $disp; region = $region
              city = $city; isp = $isp; org = $org; tz = $tz }
}

function Invoke-GeoLookup {
    param([hashtable]$Job)
    $cfg  = $Job.Cfg
    $map  = Convert-IsoText -Text $Job.IsoText
    $lang = [string]$cfg.countryLanguage
    if ($lang -ne 'en') { $lang = 'zh' }
    $apiLang = if ($lang -eq 'zh') { 'zh-CN' } else { 'en' }

    try {
        $t = [Net.SecurityProtocolType]::Tls12
        try { $t = $t -bor [Net.SecurityProtocolType]::Tls13 } catch { }
        [Net.ServicePointManager]::SecurityProtocol = $t
    } catch { }

    $modes = @()
    switch ([string]$cfg.proxyMode) {
        'direct' { $modes = @('direct'); if ($cfg.autoFallbackProxy) { $modes += 'system' } }
        'custom' { $modes = @('custom'); if ($cfg.autoFallbackProxy) { $modes += @('direct', 'system') } }
        default  { $modes = @('system'); if ($cfg.autoFallbackProxy) { $modes += 'direct' } }
    }
    if (-not $modes -or $modes.Count -eq 0) { $modes = @('system') }

    $all = @($Job.Endpoints)
    if ($cfg.source -eq 'custom' -and $cfg.customUrl) {
        $all = @(@{ id = 'custom'; name = '自定义'; kind = 'json'; url = [string]$cfg.customUrl
                    ok = ''; ip = 'ip'; country = 'country'; cc = 'country_code'
                    region = 'region'; city = 'city'; isp = 'isp'; org = 'org'; tz = 'timezone' })
    } elseif ($cfg.source -ne 'auto') {
        $first = @($all | Where-Object { $_.id -eq $cfg.source })
        if ($first.Count -gt 0) { $all = $first + @($all | Where-Object { $_.id -ne $cfg.source }) }
    }

    $ipApi = @($Job.Endpoints | Where-Object { $_.id -eq 'ip-api' })[0]
    $sw    = [System.Diagnostics.Stopwatch]::StartNew()
    $errs  = @()
    $partial = $null

    foreach ($ep in $all) {
        $url = ([string]$ep.url).Replace('{lang}', $apiLang)
        try {
            $content = Invoke-RawRequest -Url $url -Timeout ([int]$cfg.timeoutSeconds) -Modes $modes -ProxyAddress ([string]$cfg.proxyAddress)
            $res = ConvertFrom-Endpoint -Ep $ep -Content $content -Map $map -Lang $lang
            if ($res.ok) {
                $final = $res
                if (-not $res.cc -and -not $res.country -and $ep.geoUrl) {
                    # 只拿到 IP, 需要二次查询归属地
                    $geoUrl = ([string]$ep.geoUrl).Replace('{ip}', [string]$res.ip).Replace('{lang}', $apiLang)
                    try {
                        $c2 = Invoke-RawRequest -Url $geoUrl -Timeout ([int]$cfg.timeoutSeconds) -Modes $modes -ProxyAddress ([string]$cfg.proxyAddress)
                        $r2 = ConvertFrom-Endpoint -Ep $ipApi -Content $c2 -Map $map -Lang $lang
                        if ($r2.ok) { $final = $r2; $final.ip = $res.ip }
                    } catch { $errs += "二次查询 $($ep.name): $($_.Exception.Message)" }
                }
                if ($final.cc -or $final.country) {
                    $sw.Stop()
                    return @{ Ok = $true; Ip = $final.ip; Cc = $final.cc; Country = $final.country
                              Region = $final.region; City = $final.city; Isp = $final.isp
                              Org = $final.org; Tz = $final.tz; Source = $ep.name
                              Elapsed = $sw.ElapsedMilliseconds; At = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') }
                }
                if (-not $partial) { $partial = @{ res = $final; src = $ep.name } }
            }
        } catch {
            $errs += "$($ep.name): $($_.Exception.Message)"
        }
        if ($partial) { break }
    }

    $sw.Stop()
    if ($partial) {
        $r = $partial.res
        return @{ Ok = $true; Ip = $r.ip; Cc = ''; Country = '未知地区'
                  Region = ''; City = ''; Isp = ''; Org = ''; Tz = ''
                  Source = $partial.src; Elapsed = $sw.ElapsedMilliseconds; At = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') }
    }
    return @{ Ok = $false; Error = ($errs -join ' | '); Elapsed = $sw.ElapsedMilliseconds }
}

# 生成后台运行空间要执行的脚本文本(复用上面的纯函数, 避免重复实现)
function Get-WorkerText {
    $names = @('Get-Prop', 'Get-PropSplit', 'Invoke-RawRequest', 'Convert-IsoText', 'Get-CountryName',
               'Get-CodeFromName', 'Resolve-CountryName', 'ConvertFrom-Endpoint', 'Invoke-GeoLookup')
    $sb = New-Object System.Text.StringBuilder
    foreach ($n in $names) {
        $def = (Get-Item "function:$n").Definition
        [void]$sb.AppendLine("function $n { $def }")
    }
    [void]$sb.AppendLine('$Job = $args[0]')
    [void]$sb.AppendLine('Invoke-GeoLookup -Job $Job')
    return $sb.ToString()
}

# ---------------------------------------------------------------------------
# 5. Win32 / GDI 辅助类
# ---------------------------------------------------------------------------
$script:CsSource = @'
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public class IpmBufferedForm : Form {
    public IpmBufferedForm() {
        this.SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                      ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        this.UpdateStyles();
    }
}

public class IpmNative {
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] public static extern int GetWindowLong(IntPtr hWnd, int nIndex);
    [DllImport("user32.dll")] public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
    [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr hWnd, IntPtr after, int X, int Y, int cx, int cy, uint flags);
}

public class IpmColorTable : ProfessionalColorTable {
    public Color Bg, Sel, BorderCol;
    public override Color ToolStripDropDownBackground { get { return Bg; } }
    public override Color MenuBorder { get { return BorderCol; } }
    public override Color MenuItemBorder { get { return Sel; } }
    public override Color MenuItemSelected { get { return Sel; } }
    public override Color MenuItemSelectedGradientBegin { get { return Sel; } }
    public override Color MenuItemSelectedGradientEnd { get { return Sel; } }
    public override Color MenuItemPressedGradientBegin { get { return Sel; } }
    public override Color MenuItemPressedGradientEnd { get { return Sel; } }
    public override Color ImageMarginGradientBegin { get { return Bg; } }
    public override Color ImageMarginGradientMiddle { get { return Bg; } }
    public override Color ImageMarginGradientEnd { get { return Bg; } }
    public override Color SeparatorDark { get { return BorderCol; } }
    public override Color SeparatorLight { get { return BorderCol; } }
    public override Color CheckBackground { get { return Sel; } }
    public override Color CheckSelectedBackground { get { return Sel; } }
    public override Color CheckPressedBackground { get { return Sel; } }
}

public class IpmMenuRenderer : ToolStripProfessionalRenderer {
    public Color TextColor = Color.FromArgb(235, 238, 242);
    public Color DisabledColor = Color.FromArgb(120, 126, 136);
    public IpmMenuRenderer(IpmColorTable t) : base(t) { this.RoundedEdges = false; }
    protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e) {
        e.TextColor = e.Item.Enabled ? TextColor : DisabledColor;
        base.OnRenderItemText(e);
    }
    protected override void OnRenderArrow(ToolStripArrowRenderEventArgs e) {
        e.ArrowColor = TextColor;
        base.OnRenderArrow(e);
    }
    protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e) {
        using (Pen p = new Pen(((IpmColorTable)this.ColorTable).BorderCol)) {
            e.Graphics.DrawRectangle(p, 0, 0, e.ToolStrip.Width - 1, e.ToolStrip.Height - 1);
        }
    }
}
'@

try { [void]([System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')) } catch { }
try { [void]([System.Reflection.Assembly]::LoadWithPartialName('System.Drawing')) } catch { }
Add-Type -TypeDefinition $script:CsSource -ReferencedAssemblies @('System.dll', 'System.Drawing.dll', 'System.Windows.Forms.dll') -ErrorAction Stop
try { [void][IpmNative]::SetProcessDPIAware() } catch { }

# 常量
$script:WS_EX_TRANSPARENT = 0x00000020
$script:WS_EX_LAYERED     = 0x00080000
$script:GWL_EXSTYLE       = -20

# ---------------------------------------------------------------------------
# 6. 主题 / 字体 / 绘制工具
# ---------------------------------------------------------------------------
function Get-Palette {
    param([string]$Theme)
    if ($Theme -eq 'light') {
        return @{
            Bg = '#FFFFFF'; BgAlt = '#F2F4F7'; Border = '#D7DBE1'; Text = '#14171C'
            Sub = '#525A66'; Muted = '#8B939F'; Accent = '#2F6FED'; Hover = '#EDF1F7'; Err = '#D93A2B'
        }
    }
    return @{
        Bg = '#1B1D21'; BgAlt = '#23262B'; Border = '#33373E'; Text = '#F2F4F7'
        Sub = '#A9B0BA'; Muted = '#767E8A'; Accent = '#5B9BFF'; Hover = '#2B2F36'; Err = '#FF6B5E'
    }
}

$script:BadgeColors = @('#E5484D', '#F76808', '#C08A00', '#3E9B4F', '#12A594', '#0090FF', '#3E63DD', '#8E4EC6', '#E93D82', '#6E56CF')

function Get-BadgeColor {
    param([string]$Code)
    if (-not $Code) { return [System.Drawing.ColorTranslator]::FromHtml('#5A6472') }
    $sum = 0
    foreach ($ch in $Code.ToCharArray()) { $sum += [int]$ch }
    $hex = $script:BadgeColors[$sum % $script:BadgeColors.Count]
    return [System.Drawing.ColorTranslator]::FromHtml($hex)
}

function C { param([string]$Hex) return [System.Drawing.ColorTranslator]::FromHtml($Hex) }

function New-RoundPath {
    param([System.Drawing.RectangleF]$R, [single]$Radius)
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = $Radius * 2
    if ($d -le 0.5) { $path.AddRectangle($R); return $path }
    if ($d -gt $R.Width)  { $d = $R.Width }
    if ($d -gt $R.Height) { $d = $R.Height }
    $path.AddArc($R.X, $R.Y, $d, $d, 180, 90)
    $path.AddArc($R.Right - $d, $R.Y, $d, $d, 270, 90)
    $path.AddArc($R.Right - $d, $R.Bottom - $d, $d, $d, 0, 90)
    $path.AddArc($R.X, $R.Bottom - $d, $d, $d, 90, 90)
    $path.CloseFigure()
    return $path
}

function Fill-RoundRect {
    param($G, [System.Drawing.RectangleF]$R, [single]$Radius, $Brush)
    $p = New-RoundPath -R $R -Radius $Radius
    $G.FillPath($Brush, $p)
    $p.Dispose()
}

function Draw-RoundRect {
    param($G, [System.Drawing.RectangleF]$R, [single]$Radius, $Pen)
    $p = New-RoundPath -R $R -Radius $Radius
    $G.DrawPath($Pen, $p)
    $p.Dispose()
}

function Draw-Line {
    param($G, [string]$Text, $Font, $Color, [System.Drawing.RectangleF]$Rect, [string]$Align = 'Near')
    if (-not $Text) { return }
    $sf = New-Object System.Drawing.StringFormat
    $sf.Trimming = [System.Drawing.StringTrimming]::EllipsisCharacter
    $sf.FormatFlags = [System.Drawing.StringFormatFlags]::NoWrap
    $sf.Alignment = [System.Drawing.StringAlignment]::$Align
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
    $b = New-Object System.Drawing.SolidBrush $Color
    $G.DrawString($Text, $Font, $b, $Rect, $sf)
    $b.Dispose(); $sf.Dispose()
}

function New-Font {
    param([string]$Family, [single]$Px, [System.Drawing.FontStyle]$Style = [System.Drawing.FontStyle]::Regular)
    try { return New-Object System.Drawing.Font($Family, $Px, $Style, [System.Drawing.GraphicsUnit]::Pixel) }
    catch { return New-Object System.Drawing.Font('Segoe UI', $Px, $Style, [System.Drawing.GraphicsUnit]::Pixel) }
}

function New-MenuRenderer {
    param([bool]$Dark)
    $ct = New-Object IpmColorTable
    $r  = New-Object IpmMenuRenderer -ArgumentList $ct
    if ($Dark) {
        $ct.Bg = [System.Drawing.Color]::FromArgb(35, 38, 43)
        $ct.Sel = [System.Drawing.Color]::FromArgb(58, 64, 74)
        $ct.BorderCol = [System.Drawing.Color]::FromArgb(62, 68, 77)
        $r.TextColor = [System.Drawing.Color]::FromArgb(235, 238, 242)
        $r.DisabledColor = [System.Drawing.Color]::FromArgb(118, 126, 138)
    } else {
        $ct.Bg = [System.Drawing.Color]::FromArgb(252, 252, 253)
        $ct.Sel = [System.Drawing.Color]::FromArgb(228, 236, 250)
        $ct.BorderCol = [System.Drawing.Color]::FromArgb(214, 219, 226)
        $r.TextColor = [System.Drawing.Color]::FromArgb(20, 23, 28)
        $r.DisabledColor = [System.Drawing.Color]::FromArgb(150, 156, 166)
    }
    return $r
}

# ---------------------------------------------------------------------------
# 7. 运行时状态
# ---------------------------------------------------------------------------
$script:Busy       = $false
$script:DpiScale   = 1.0
$script:SpinAngle  = 0
$script:HoverBtn   = $false
$script:DragMoved  = $false
$script:DragOrigin = $null
$script:FormOrigin = $null
$script:NextAt     = $null
$script:LastError  = ''
$script:Result = @{ Ok = $false; Ip = ''; Cc = ''; Country = ''; Region = ''; City = ''
                    Isp = ''; Org = ''; Tz = ''; Source = ''; At = ''; Cached = $false; Elapsed = 0 }

# ---------------------------------------------------------------------------
# 8. 异步查询(后台运行空间 + 轮询, 保证界面不卡)
# ---------------------------------------------------------------------------
function Initialize-Worker {
    $script:Runspace = [runspacefactory]::CreateRunspace()
    $script:Runspace.ApartmentState = 'MTA'
    $script:Runspace.ThreadOptions = 'ReuseThread'
    $script:Runspace.Open()
    $script:WorkerText = Get-WorkerText
}

function Start-LookupAsync {
    if ($script:AsyncPs) { return $null }
    $job = @{ Cfg = $script:Cfg.Clone(); Endpoints = $script:Endpoints; IsoText = $script:IsoText }
    $ps = [PowerShell]::Create()
    $ps.Runspace = $script:Runspace
    [void]$ps.AddScript($script:WorkerText).AddArgument($job)
    $handle = $ps.BeginInvoke()
    return @{ Ps = $ps; Handle = $handle }
}

function Complete-LookupAsync {
    param($Async)
    if (-not $Async) { return $null }
    $out = $null
    try {
        $raw = $Async.Ps.EndInvoke($Async.Handle)
        foreach ($item in @($raw)) {
            if ($item -is [hashtable] -and $item.ContainsKey('Ok')) { $out = $item; break }
        }
        if (-not $out) {
            $err = ($Async.Ps.Streams.Error | ForEach-Object { $_.ToString() }) -join ' | '
            $out = @{ Ok = $false; Error = ($err -replace '\s+', ' ') }
        }
    } catch {
        $out = @{ Ok = $false; Error = $_.Exception.Message }
    } finally {
        try { $Async.Ps.Dispose() } catch { }
    }
    return $out
}

# ---------------------------------------------------------------------------
# 9. 主窗口
# ---------------------------------------------------------------------------
function Update-Layout {
    $c = $script:Cfg
    $p = $script:Palette
    $k = $script:DpiScale * [double]$c.fontScale

    $w = switch ([string]$c.size) { 'small' { 268 } 'large' { 432 } default { 336 } }
    $pad = [int](10 * $k)

    $lines = @()
    if ($c.showCountry) { $lines += @{ h = [int](19 * $k) } }
    if ($c.showIp)      { $lines += @{ h = [int](17 * $k) } }
    if ($c.showCity -or $c.showIsp) { $lines += @{ h = [int](14 * $k) } }
    if ($lines.Count -eq 0) { $lines += @{ h = [int](19 * $k) } }
    if ($c.showUpdatedAt -and $c.showIp) { }

    $gap = [int](3 * $k)
    $contentH = 0
    foreach ($l in $lines) { $contentH += $l.h }
    $contentH += $gap * ($lines.Count - 1)

    $badgeSize = if ($c.badgeMode -eq 'emoji') { [int](30 * $k) } elseif ($c.badgeMode -eq 'badge') { [int](34 * $k) } else { 0 }
    $btnSize = [int](28 * $k)
    $showCount = [bool]$c.showCountdown
    $colNeeded = if ($showCount) { $btnSize + [int](4 * $k) + [int](11 * $k) } else { $btnSize }

    $h = [Math]::Max($contentH + 2 * $pad, $badgeSize + 2 * $pad)
    $h = [Math]::Max($h, $colNeeded + 2 * $pad)
    $h = [int]$h

    $textX = $pad + $(if ($badgeSize -gt 0) { $badgeSize + [int](10 * $k) } else { 0 })
    $rightW = $btnSize + $pad + [int](2 * $k)
    $textW = [Math]::Max(40, $w - $textX - $rightW)

    $lineRects = @()
    $y = [int](($h - $contentH) / 2)
    foreach ($l in $lines) {
        $lineRects += [System.Drawing.RectangleF]::new([single]$textX, [single]$y, [single]$textW, [single]$l.h)
        $y += $l.h + $gap
    }

    $btnRect = [System.Drawing.RectangleF]::new([single]($w - $pad - $btnSize), [single](($h - $colNeeded) / 2), [single]$btnSize, [single]$btnSize)
    $cntRect = [System.Drawing.RectangleF]::new([single]($w - $pad - $btnSize - [int](10 * $k)), [single]($btnRect.Bottom + [int](2 * $k)), [single]($btnSize + [int](10 * $k)), [single](11 * $k))
    $badgeRect = [System.Drawing.RectangleF]::new([single]$pad, [single](($h - $badgeSize) / 2), [single]$badgeSize, [single]$badgeSize)

    $script:Layout = @{ W = $w; H = $h; Pad = $pad; K = $k; Btn = $btnRect; Cnt = $cntRect
                        Badge = $badgeRect; BadgeSize = $badgeSize; Lines = $lineRects; TextW = $textW
                        ShowCount = $showCount }

    $f = $script:UI.Form
    if ($f.ClientSize.Width -ne $w -or $f.ClientSize.Height -ne $h) {
        $f.ClientSize = New-Object System.Drawing.Size($w, $h)
    }
    $radius = [single](12 * $k)
    $path = New-RoundPath -R ([System.Drawing.RectangleF]::new(0, 0, $w, $h)) -Radius $radius
    $old = $f.Region
    $f.Region = New-Object System.Drawing.Region($path)
    if ($old) { $old.Dispose() }
    $path.Dispose()
}

function Rebuild-Fonts {
    $k = $script:DpiScale * [double]$script:Cfg.fontScale
    if ($script:Fonts) {
        foreach ($fnt in $script:Fonts.Values) { try { $fnt.Dispose() } catch { } }
    }
    $script:Fonts = @{
        Country  = New-Font 'Segoe UI' ([single](13.5 * $k)) ([System.Drawing.FontStyle]::Bold)
        Ip       = New-Font 'Consolas' ([single](11.5 * $k)) ([System.Drawing.FontStyle]::Regular)
        Detail   = New-Font 'Segoe UI' ([single](9.8 * $k))  ([System.Drawing.FontStyle]::Regular)
        Badge    = New-Font 'Segoe UI' ([single](12 * $k))   ([System.Drawing.FontStyle]::Bold)
        Emoji    = New-Font 'Segoe UI Emoji' ([single](19 * $k)) ([System.Drawing.FontStyle]::Regular)
        Glyph    = New-Font 'Segoe UI' ([single](14 * $k))   ([System.Drawing.FontStyle]::Regular)
        Count    = New-Font 'Segoe UI' ([single](8.6 * $k))  ([System.Drawing.FontStyle]::Regular)
    }
}

function Get-DetailLine {
    $r = $script:Result
    $parts = @()
    if ($script:Cfg.showCity -and $r.City) {
        if ($r.Region -and $r.Region -ne $r.City) { $parts += ("$($r.Region) · $($r.City)") } else { $parts += $r.City }
    }
    if ($script:Cfg.showIsp -and $r.Isp) { $parts += $r.Isp }
    return ($parts -join ' · ')
}

function Get-FlagEmoji {
    param([string]$Code)
    if (-not $Code -or $Code.Length -lt 2) { return '' }
    $a = [int][char]$Code.ToUpper()[0]
    $b = [int][char]$Code.ToUpper()[1]
    if ($a -lt 65 -or $a -gt 90 -or $b -lt 65 -or $b -gt 90) { return '' }
    return ([char](0x1F1E6 + $a - 65)).ToString() + ([char](0x1F1E6 + $b - 65)).ToString()
}

function Invoke-BarPaint {
    param($G)
    $p = $script:Palette
    $L = $script:Layout
    $r = $script:Result
    $G.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $G.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit

    $bg = New-Object System.Drawing.SolidBrush (C $p.Bg)
    $G.FillRectangle($bg, 0, 0, $L.W, $L.H)
    $bg.Dispose()

    $penBorder = New-Object System.Drawing.Pen ((C $p.Border), 1)
    $G.DrawRectangle($penBorder, 0, 0, $L.W - 1, $L.H - 1)
    $penBorder.Dispose()

    # --- 徽标 / 旗帜 ---
    if ($L.BadgeSize -gt 0) {
        if ($script:Cfg.badgeMode -eq 'emoji') {
            $emo = Get-FlagEmoji $r.Cc
            if (-not $emo) { $emo = if ($r.Cc) { $r.Cc } else { '··' } }
            Draw-Line -G $G -Text $emo -Font $script:Fonts.Emoji -Color (C $p.Text) -Rect $L.Badge -Align 'Center'
        } else {
            $bc = Get-BadgeColor $r.Cc
            $bb = New-Object System.Drawing.SolidBrush $bc
            $G.FillEllipse($bb, $L.Badge)
            $bb.Dispose()
            $code = if ($r.Cc) { $r.Cc.ToUpper() } else { '?' }
            Draw-Line -G $G -Text $code -Font $script:Fonts.Badge -Color ([System.Drawing.Color]::White) -Rect $L.Badge -Align 'Center'
        }
    }

    # --- 文本行 ---
    $i = 0
    if ($script:Cfg.showCountry) {
        $txt = if ($r.Country) { $r.Country } else { '正在获取…' }
        $col = if ($r.Country) { C $p.Text } else { C $p.Muted }
        Draw-Line -G $G -Text $txt -Font $script:Fonts.Country -Color $col -Rect $L.Lines[$i]
        $i++
    }
    if ($script:Cfg.showIp) {
        if ($script:Busy -and -not $r.Ip) { $txt = '查询中…'; $col = C $p.Accent }
        elseif ($r.Ip) { $txt = $r.Ip; $col = C $p.Sub }
        else { $txt = '—'; $col = C $p.Muted }
        Draw-Line -G $G -Text $txt -Font $script:Fonts.Ip -Color $col -Rect $L.Lines[$i]
        $i++
    }
    if ($script:Cfg.showCity -or $script:Cfg.showIsp) {
        $d = Get-DetailLine
        if (-not $d -and $script:Cfg.showUpdatedAt -and $r.At) { $d = "更新于 $($r.At.Substring(11))" }
        Draw-Line -G $G -Text $d -Font $script:Fonts.Detail -Color (C $p.Muted) -Rect $L.Lines[$i]
        $i++
    }

    # --- 刷新按钮 ---
    $btn = $L.Btn
    if ($script:HoverBtn -and -not $script:Busy) {
        $hb = New-Object System.Drawing.SolidBrush (C $p.Hover)
        Fill-RoundRect -G $G -R $btn -Radius ([single](8 * $L.K)) -Brush $hb
        $hb.Dispose()
    }
    if ($script:Busy) {
        $penArc = New-Object System.Drawing.Pen ((C $p.Accent), [single][Math]::Max(1.6, 2.0 * $L.K))
        $penArc.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
        $penArc.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
        # 注意: 不能写成 New-Object RectangleF(a + b, c + d, ...)
        # PowerShell 中逗号优先级高于加减号, 那样会被解析成数组相加而报 op_Addition 错误
        $inset = [single](6 * $L.K)
        $ax = [single]([single]$btn.X + $inset)
        $ay = [single]([single]$btn.Y + $inset)
        $aw = [single]([single]$btn.Width - (2 * $inset))
        $ah = [single]([single]$btn.Height - (2 * $inset))
        $arcRect = [System.Drawing.RectangleF]::new($ax, $ay, $aw, $ah)
        $G.DrawArc($penArc, $arcRect, [single]$script:SpinAngle, [single]260)
        $penArc.Dispose()
    } else {
        Draw-Line -G $G -Text ([char]0x21BB) -Font $script:Fonts.Glyph -Color (C $(if ($script:HoverBtn) { $p.Accent } else { $p.Sub })) -Rect $btn -Align 'Center'
    }

    # --- 倒计时 / 状态 ---
    if ($L.ShowCount) {
        $txt = ''
        $col = C $p.Muted
        if ($script:Busy) { $txt = '···'; $col = C $p.Accent }
        elseif ($script:LastError) { $txt = '失败'; $col = C $p.Err }
        elseif ($script:NextAt) {
            $sec = [int][Math]::Ceiling(([datetime]$script:NextAt - (Get-Date)).TotalSeconds)
            if ($sec -lt 0) { $sec = 0 }
            $txt = "${sec}s"
        }
        Draw-Line -G $G -Text $txt -Font $script:Fonts.Count -Color $col -Rect $L.Cnt -Align 'Center'
    }
}

function Update-Tooltip {
    $r = $script:Result
    $lines = @()
    $lines += "$($script:AppName) v$($script:AppVer)"
    $lines += '──────────────'
    $lines += "IP 地址 : " + $(if ($r.Ip) { $r.Ip } else { '—' })
    $lines += "国家/地区: " + $(if ($r.Country) { $r.Country + $(if ($r.Cc) { " ($($r.Cc))" } else { '' }) } else { '—' })
    if ($r.Region -or $r.City) {
        $cityParts = @(@($r.Region, $r.City) | Where-Object { $_ } | Select-Object -Unique)
        $lines += "城市 : " + ($cityParts -join ' · ')
    }
    if ($r.Isp) { $lines += "运营商 : $($r.Isp)" }
    if ($r.Tz)  { $lines += "时区 : $($r.Tz)" }
    if ($r.Source) { $lines += "数据源 : $($r.Source)  ($($r.Elapsed) ms)" }
    if ($r.At)  { $lines += "更新时间: $($r.At)" + $(if ($r.Cached) { ' (缓存)' } else { '' }) }
    if ($script:LastError) { $lines += "上次错误: $($script:LastError)" }
    $lines += '──────────────'
    $lines += '左键点击 = 立即刷新'
    $lines += '右键 = 菜单 / 设置'
    $lines += '拖动 = 移动位置'
    try { $script:UI.Tip.SetToolTip($script:UI.Form, ($lines -join "`r`n")) } catch { }
}

# ---------------------------------------------------------------------------
# 10. 配置应用 / 刷新流程
# ---------------------------------------------------------------------------
function Apply-Config {
    param([switch]$KeepPosition)
    $script:Palette = Get-Palette -Theme ([string]$script:Cfg.theme)
    $f = $script:UI.Form
    $f.BackColor = C $script:Palette.Bg
    $f.TopMost = [bool]$script:Cfg.alwaysOnTop
    $f.Opacity = [double]$script:Cfg.opacity
    Rebuild-Fonts
    Update-Layout
    Set-MouseThrough -Enabled ([bool]$script:Cfg.mouseThrough)
    Update-MenuRenderer
    Sync-Menu
    $script:NextAt = (Get-Date).AddSeconds([int]$script:Cfg.intervalSeconds)
    Update-Tooltip
    $f.Invalidate()
}

function Set-MouseThrough {
    param([bool]$Enabled)
    try {
        $h = $script:UI.Form.Handle
        $ex = [IpmNative]::GetWindowLong($h, $script:GWL_EXSTYLE)
        if ($Enabled) { $ex = $ex -bor $script:WS_EX_TRANSPARENT -bor $script:WS_EX_LAYERED }
        else { $ex = $ex -band (-bnot $script:WS_EX_TRANSPARENT) }
        [void][IpmNative]::SetWindowLong($h, $script:GWL_EXSTYLE, $ex)
    } catch { Write-Log "设置鼠标穿透失败: $($_.Exception.Message)" }
}

function Start-Refresh {
    if ($script:Busy) { return }
    $script:Busy = $true
    $script:LastError = ''
    $script:SpinAngle = 0
    $script:SpinTimer.Start()
    $script:UI.Form.Invalidate()
    try {
        $script:Async = Start-LookupAsync
        if ($script:Async) { $script:PollTimer.Start() }
        else { $script:Busy = $false; $script:SpinTimer.Stop() }
    } catch {
        $script:Busy = $false
        $script:SpinTimer.Stop()
        $script:LastError = $_.Exception.Message
        Write-Log "查询启动失败: $($_.Exception.Message)"
    }
    Update-Tooltip
}

function Apply-Result {
    param($Res)
    $script:Busy = $false
    $script:SpinTimer.Stop()
    if ($Res -and $Res.Ok) {
        $script:Result = @{ Ok = $true; Ip = $Res.Ip; Cc = $Res.Cc; Country = $Res.Country
                            Region = $Res.Region; City = $Res.City; Isp = $Res.Isp; Org = $Res.Org
                            Tz = $Res.Tz; Source = $Res.Source; At = $Res.At; Elapsed = $Res.Elapsed; Cached = $false }
        $script:LastError = ''
        Save-Cache
        Write-Log "OK  $($Res.Ip)  $($Res.Country)  [$($Res.Source)]  $($Res.Elapsed)ms"
    } else {
        $msg = ''
        if ($Res) { $msg = [string]$Res.Error }
        if (-not $msg) { $msg = '未知错误' }
        $script:LastError = $msg
        Write-Log "FAIL $msg"
    }
    $script:NextAt = (Get-Date).AddSeconds([int]$script:Cfg.intervalSeconds)
    Update-Tooltip
    try { $script:UI.Tray.Text = Get-TrayText } catch { }
    $script:UI.Form.Invalidate()
}

function Get-TrayText {
    $r = $script:Result
    $s = "$($script:AppName)"
    if ($r.Ip) { $s += "`r`n$($r.Ip)" }
    if ($r.Country) { $s += "`r`n$($r.Country)" }
    if ($r.Ip.Length -gt 60) { $s = $s.Substring(0, 60) }
    if ($s.Length -gt 63) { $s = $s.Substring(0, 63) }
    return $s
}

# ---------------------------------------------------------------------------
# 11. 菜单
# ---------------------------------------------------------------------------
function Set-CfgValue {
    param([string]$Key, $Value, [switch]$NoSave)
    $script:Cfg[$Key] = $Value
    if (-not $NoSave) { Save-Config }
    switch ($Key) {
        'intervalSeconds' { $script:NextAt = (Get-Date).AddSeconds([int]$script:Cfg.intervalSeconds) }
        'autoStart'       { Set-AutoStart -Enabled ([bool]$Value) | Out-Null }
    }
    Apply-Config
}

function New-MenuItem {
    param([string]$Text, $OnClick, [bool]$Checkable = $false, [bool]$Checked = $false)
    $mi = New-Object System.Windows.Forms.ToolStripMenuItem $Text
    if ($Checkable) { $mi.CheckOnClick = $false; $mi.Checked = $Checked }
    if ($OnClick) { $mi.Add_Click($OnClick) }
    return $mi
}

function New-ChoiceItem {
    param([string]$Text, [string]$Key, $Value, [bool]$Checkable = $true)
    $mi = New-Object System.Windows.Forms.ToolStripMenuItem $Text
    $mi.Tag = @{ kind = 'choice'; key = $Key; value = $Value }
    if ($Checkable) { $mi.CheckOnClick = $false; $mi.Checked = ($script:Cfg[$Key] -eq $Value) }
    $mi.Add_Click({ param($s, $e) Set-CfgValue -Key $s.Tag.key -Value $s.Tag.value })
    return $mi
}

function New-ToggleItem {
    param([string]$Text, [string]$Key)
    $mi = New-Object System.Windows.Forms.ToolStripMenuItem $Text
    $mi.Tag = @{ kind = 'toggle'; key = $Key }
    $mi.CheckOnClick = $false
    $mi.Checked = [bool]$script:Cfg[$Key]
    $mi.Add_Click({ param($s, $e) Set-CfgValue -Key $s.Tag.key -Value (-not [bool]$script:Cfg[$s.Tag.key]) })
    return $mi
}

function New-ActionItem {
    param([string]$Text, [scriptblock]$Action)
    $mi = New-Object System.Windows.Forms.ToolStripMenuItem $Text
    if ($Action) { $mi.Add_Click($Action) }
    return $mi
}

function Build-Menu {
    $m = New-Object System.Windows.Forms.ContextMenuStrip
    $m.ShowImageMargin = $false
    $m.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)

    [void]$m.Items.Add((New-ActionItem '立即刷新 (F5)' { Start-Refresh }))
    [void]$m.Items.Add((New-ActionItem '复制 IP 地址' {
        if ($script:Result.Ip) { try { [System.Windows.Forms.Clipboard]::SetText($script:Result.Ip) } catch { } }
    }))
    [void]$m.Items.Add((New-ActionItem '复制完整信息' {
        $r = $script:Result
        $t = @("IP: $($r.Ip)", "国家/地区: $($r.Country)", "代码: $($r.Cc)", "城市: $($r.City)", "运营商: $($r.Isp)", "时区: $($r.Tz)", "更新时间: $($r.At)") -join "`r`n"
        try { [System.Windows.Forms.Clipboard]::SetText($t) } catch { }
    }))
    [void]$m.Items.Add((New-ActionItem '搜索该 IP 归属' {
        if ($script:Result.Ip) { Start-Process ("https://www.bing.com/search?q=" + $script:Result.Ip) }
    }))

    [void]$m.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))

    $miDisp = New-Object System.Windows.Forms.ToolStripMenuItem '显示内容'
    [void]$miDisp.DropDownItems.Add((New-ToggleItem '国家 / 地区' 'showCountry'))
    [void]$miDisp.DropDownItems.Add((New-ToggleItem 'IP 地址' 'showIp'))
    [void]$miDisp.DropDownItems.Add((New-ToggleItem '城市 / 地区' 'showCity'))
    [void]$miDisp.DropDownItems.Add((New-ToggleItem '运营商' 'showIsp'))
    [void]$miDisp.DropDownItems.Add((New-ToggleItem '刷新倒计时' 'showCountdown'))
    [void]$miDisp.DropDownItems.Add((New-ToggleItem '更新时间' 'showUpdatedAt'))
    [void]$miDisp.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    [void]$miDisp.DropDownItems.Add((New-ChoiceItem '色块徽标' 'badgeMode' 'badge'))
    [void]$miDisp.DropDownItems.Add((New-ChoiceItem 'Emoji 旗帜' 'badgeMode' 'emoji'))
    [void]$miDisp.DropDownItems.Add((New-ChoiceItem '不显示徽标' 'badgeMode' 'none'))
    [void]$m.Items.Add($miDisp)

    $miInt = New-Object System.Windows.Forms.ToolStripMenuItem '刷新间隔'
    foreach ($pair in @(@('30 秒', 30), @('1 分钟', 60), @('2 分钟', 120), @('5 分钟', 300), @('15 分钟', 900), @('30 分钟', 1800))) {
        [void]$miInt.DropDownItems.Add((New-ChoiceItem $pair[0] 'intervalSeconds' $pair[1]))
    }
    [void]$miInt.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    [void]$miInt.DropDownItems.Add((New-ActionItem '自定义…' { Show-SettingsDialog -Tab 0 }))
    [void]$m.Items.Add($miInt)

    $miSrc = New-Object System.Windows.Forms.ToolStripMenuItem '数据源'
    [void]$miSrc.DropDownItems.Add((New-ChoiceItem '自动检测' 'source' 'auto'))
    [void]$miSrc.DropDownItems.Add((New-ChoiceItem 'ip-api.com' 'source' 'ip-api'))
    [void]$miSrc.DropDownItems.Add((New-ChoiceItem 'ipwho.is' 'source' 'ipwho'))
    [void]$miSrc.DropDownItems.Add((New-ChoiceItem 'ipinfo.io' 'source' 'ipinfo'))
    [void]$miSrc.DropDownItems.Add((New-ChoiceItem 'zxinc.org' 'source' 'zxinc'))
    [void]$miSrc.DropDownItems.Add((New-ChoiceItem 'ip.sb (仅 IP)' 'source' 'ipsb'))
    [void]$m.Items.Add($miSrc)

    $miWin = New-Object System.Windows.Forms.ToolStripMenuItem '窗口'
    [void]$miWin.DropDownItems.Add((New-ToggleItem '总在最前' 'alwaysOnTop'))
    [void]$miWin.DropDownItems.Add((New-ToggleItem '锁定位置' 'locked'))
    [void]$miWin.DropDownItems.Add((New-ToggleItem '贴边吸附' 'snapToEdge'))
    [void]$miWin.DropDownItems.Add((New-ToggleItem '鼠标穿透' 'mouseThrough'))
    [void]$miWin.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    $miSize = New-Object System.Windows.Forms.ToolStripMenuItem '尺寸'
    [void]$miSize.DropDownItems.Add((New-ChoiceItem '小' 'size' 'small'))
    [void]$miSize.DropDownItems.Add((New-ChoiceItem '中' 'size' 'medium'))
    [void]$miSize.DropDownItems.Add((New-ChoiceItem '大' 'size' 'large'))
    [void]$miWin.DropDownItems.Add($miSize)
    $miOpa = New-Object System.Windows.Forms.ToolStripMenuItem '不透明度'
    foreach ($v in @(0.5, 0.7, 0.85, 0.95, 1.0)) {
        [void]$miOpa.DropDownItems.Add((New-ChoiceItem ("{0}%" -f [int]($v * 100)) 'opacity' $v))
    }
    [void]$miWin.DropDownItems.Add($miOpa)
    $miTheme = New-Object System.Windows.Forms.ToolStripMenuItem '主题'
    [void]$miTheme.DropDownItems.Add((New-ChoiceItem '深色' 'theme' 'dark'))
    [void]$miTheme.DropDownItems.Add((New-ChoiceItem '浅色' 'theme' 'light'))
    [void]$miWin.DropDownItems.Add($miTheme)
    [void]$m.Items.Add($miWin)

    [void]$m.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    [void]$m.Items.Add((New-ActionItem '设置…' { Show-SettingsDialog }))
    [void]$m.Items.Add((New-ToggleItem '开机自动启动' 'autoStart'))
    [void]$m.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    [void]$m.Items.Add((New-ActionItem '隐藏到托盘' { $script:UI.Form.Hide() }))
    [void]$m.Items.Add((New-ActionItem '退出' { Exit-App }))

    $m.Add_Opening({ Sync-Menu })
    return $m
}

function Sync-Menu {
    $m = $script:UI.Menu
    if (-not $m) { return }
    foreach ($item in (Get-MenuItemsRecursive $m.Items)) {
        if ($item.Tag -is [hashtable]) {
            $key = $item.Tag.key
            if ($item.Tag.kind -eq 'toggle') { $item.Checked = [bool]$script:Cfg[$key] }
            elseif ($item.Tag.kind -eq 'choice') { $item.Checked = ($script:Cfg[$key] -eq $item.Tag.value) }
        }
    }
}

function Get-MenuItemsRecursive {
    param($Items)
    $out = @()
    foreach ($i in $Items) {
        $out += $i
        if ($i -is [System.Windows.Forms.ToolStripMenuItem] -and $i.DropDownItems.Count -gt 0) {
            $out += Get-MenuItemsRecursive $i.DropDownItems
        }
    }
    return $out
}

function Update-MenuRenderer {
    $dark = ($script:Cfg.theme -ne 'light')
    $r = New-MenuRenderer -Dark $dark
    foreach ($m in @($script:UI.Menu)) { if ($m) { $m.Renderer = $r } }
    if ($script:UI.Tray -and $script:UI.Tray.ContextMenuStrip) { $script:UI.Tray.ContextMenuStrip.Renderer = $r }
    if ($script:UI.DlgMenu) { $script:UI.DlgMenu.Renderer = $r }
}

# ---------------------------------------------------------------------------
# 12. 开机自启
# ---------------------------------------------------------------------------
function Get-LauncherPath {
    return (Join-Path $script:DataDir 'launch.vbs')
}

function Ensure-Launcher {
    $vbs = Get-LauncherPath
    $ps1 = Join-Path $script:AppDir 'IpGeoMonitor.ps1'
    $body = @"
' IP 地理监控 - 无窗口启动器 (自动生成, 请勿手动修改)
Set sh = CreateObject("WScript.Shell")
sh.Run "powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File ""$ps1""", 0, False
"@
    try { Set-Content -Path $vbs -Value $body -Encoding ASCII } catch { Write-Log "启动器写入失败: $($_.Exception.Message)" }
    return $vbs
}

function Test-AutoStart {
    try {
        $startup = [Environment]::GetFolderPath('Startup')
        return (Test-Path (Join-Path $startup 'IP地理监控.lnk'))
    } catch { return $false }
}

function Set-AutoStart {
    param([bool]$Enabled)
    $startup = [Environment]::GetFolderPath('Startup')
    $lnk = Join-Path $startup 'IP地理监控.lnk'
    try {
        if ($Enabled) {
            $vbs = Ensure-Launcher
            $ws = New-Object -ComObject WScript.Shell
            $sc = $ws.CreateShortcut($lnk)
            $sc.TargetPath = "$env:SystemRoot\System32\wscript.exe"
            $sc.Arguments = '"' + $vbs + '"'
            $sc.WorkingDirectory = $script:AppDir
            $sc.Description = $script:AppName
            $sc.IconLocation = "$env:SystemRoot\System32\shell32.dll,13"
            $sc.Save()
            Write-Log '已设置开机自启'
        } else {
            if (Test-Path $lnk) { Remove-Item $lnk -Force }
            Write-Log '已取消开机自启'
        }
        return $true
    } catch {
        Write-Log "开机自启设置失败: $($_.Exception.Message)"
        return $false
    }
}

# ---------------------------------------------------------------------------
# 13. 设置窗口
# ---------------------------------------------------------------------------
function New-DarkCombo {
    param([int]$X, [int]$Y, [int]$W, [string[]]$Items, [int]$Index = 0)
    $cb = New-Object System.Windows.Forms.ComboBox
    $cb.DropDownStyle = 'DropDownList'
    $cb.FlatStyle = 'Flat'
    $cb.Location = New-Object System.Drawing.Point($X, $Y)
    $cb.Size = New-Object System.Drawing.Size($W, 24)
    [void]$cb.Items.AddRange($Items)
    if ($Index -ge 0 -and $Index -lt $Items.Count) { $cb.SelectedIndex = $Index }
    return $cb
}

function New-DarkCheck {
    param([string]$Text, [int]$X, [int]$Y, [int]$W = 180, [bool]$Checked = $false)
    $ck = New-Object System.Windows.Forms.CheckBox
    $ck.Text = $Text
    $ck.FlatStyle = 'Flat'
    $ck.AutoSize = $false
    $ck.Location = New-Object System.Drawing.Point($X, $Y)
    $ck.Size = New-Object System.Drawing.Size($W, 22)
    $ck.Checked = $Checked
    return $ck
}

function New-DarkNum {
    param([int]$X, [int]$Y, [int]$W, [int]$Min, [int]$Max, [int]$Value)
    $n = New-Object System.Windows.Forms.NumericUpDown
    $n.Minimum = $Min; $n.Maximum = $Max; $n.Value = $Value
    $n.Location = New-Object System.Drawing.Point($X, $Y)
    $n.Size = New-Object System.Drawing.Size($W, 24)
    $n.BorderStyle = 'FixedSingle'
    return $n
}

function New-DarkText {
    param([int]$X, [int]$Y, [int]$W, [string]$Text = '')
    $t = New-Object System.Windows.Forms.TextBox
    $t.Text = $Text
    $t.Location = New-Object System.Drawing.Point($X, $Y)
    $t.Size = New-Object System.Drawing.Size($W, 24)
    $t.BorderStyle = 'FixedSingle'
    return $t
}

function New-GroupPanel {
    param([string]$Title, [int]$Y, [int]$Height, $Theme)
    $pnl = New-Object System.Windows.Forms.Panel
    $pnl.Location = New-Object System.Drawing.Point(14, $Y)
    $pnl.Size = New-Object System.Drawing.Size(492, $Height)
    $pnl.BackColor = C $Theme.BgAlt
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Title
    $lbl.Location = New-Object System.Drawing.Point(12, 8)
    $lbl.AutoSize = $true
    $lbl.Font = New-Object System.Drawing.Font('Segoe UI', 9.5, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point)
    $lbl.ForeColor = C $Theme.Accent
    $pnl.Controls.Add($lbl)
    return $pnl
}

function New-FieldLabel {
    param([string]$Text, [int]$X, [int]$Y, [int]$W = 130)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text
    $l.Location = New-Object System.Drawing.Point($X, $Y)
    $l.Size = New-Object System.Drawing.Size($W, 22)
    $l.TextAlign = 'MiddleLeft'
    return $l
}

function Apply-ThemeRecursive {
    param($Control, $Theme)
    foreach ($c in $Control.Controls) {
        if ($c -is [System.Windows.Forms.Panel] -or $c -is [System.Windows.Forms.GroupBox]) {
            $c.BackColor = C $Theme.BgAlt
            $c.ForeColor = C $Theme.Text
        } elseif ($c -is [System.Windows.Forms.Button]) {
            $c.FlatStyle = 'Flat'
            $c.FlatAppearance.BorderColor = C $Theme.Border
            $c.BackColor = C $Theme.BgAlt
            $c.ForeColor = C $Theme.Text
            $c.UseVisualStyleBackColor = $false
        } elseif ($c -is [System.Windows.Forms.ComboBox] -or $c -is [System.Windows.Forms.TextBox] -or $c -is [System.Windows.Forms.NumericUpDown]) {
            $c.BackColor = (C $(if ($Theme.Bg -eq '#FFFFFF') { '#FFFFFF' } else { '#2A2E35' }))
            $c.ForeColor = C $Theme.Text
        } elseif ($c -is [System.Windows.Forms.CheckBox]) {
            $c.ForeColor = C $Theme.Text
            $c.BackColor = [System.Drawing.Color]::Transparent
        } elseif ($c -is [System.Windows.Forms.TrackBar]) {
            $c.BackColor = C $Theme.BgAlt
        } elseif ($c -is [System.Windows.Forms.Label]) {
            $c.ForeColor = C $Theme.Sub
            $c.BackColor = [System.Drawing.Color]::Transparent
        } else {
            $c.ForeColor = C $Theme.Text
        }
        if ($c.HasChildren) { Apply-ThemeRecursive -Control $c -Theme $Theme }
    }
}

function Show-SettingsDialog {
    param([int]$Tab = -1)
    if ($script:UI.Dlg -and -not $script:UI.Dlg.IsDisposed) { $script:UI.Dlg.Activate(); return }

    $theme = $script:Palette
    $c = $script:Cfg

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "$($script:AppName) - 设置"
    $dlg.FormBorderStyle = 'FixedSingle'
    $dlg.StartPosition = 'CenterScreen'
    $dlg.ClientSize = New-Object System.Drawing.Size(520, 700)
    $dlg.BackColor = C $theme.Bg
    $dlg.ForeColor = C $theme.Text
    $dlg.TopMost = $true
    $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)

    $scroll = New-Object System.Windows.Forms.Panel
    $scroll.Location = New-Object System.Drawing.Point(0, 0)
    $scroll.Size = New-Object System.Drawing.Size(520, 646)
    $scroll.AutoScroll = $true
    $scroll.BackColor = C $theme.Bg
    $dlg.Controls.Add($scroll)

    $y = 12

    # --- 刷新 ---
    $g1 = New-GroupPanel -Title '刷新' -Y $y -Height 96 -Theme $theme
    $g1.Controls.Add((New-FieldLabel '刷新间隔（秒）' 14 36))
    $nInterval = New-DarkNum 160 34 90 10 3600 ([int]$c.intervalSeconds)
    $g1.Controls.Add($nInterval)
    $g1.Controls.Add((New-FieldLabel '请求超时（秒）' 268 36))
    $nTimeout = New-DarkNum 410 34 70 3 60 ([int]$c.timeoutSeconds)
    $g1.Controls.Add($nTimeout)
    $ckStart = New-DarkCheck '启动后立即刷新一次' 14 66 200 ([bool]$c.refreshOnStart)
    $g1.Controls.Add($ckStart)
    $ckFallback = New-DarkCheck '失败时自动换其他方式重试' 232 66 250 ([bool]$c.autoFallbackProxy)
    $g1.Controls.Add($ckFallback)
    $scroll.Controls.Add($g1)
    $y += $g1.Height + 10

    # --- 数据源 ---
    $g2 = New-GroupPanel -Title '数据源 / 网络' -Y $y -Height 130 -Theme $theme
    $g2.Controls.Add((New-FieldLabel '优先数据源' 14 36))
    $srcItems = @('自动检测', 'ip-api.com', 'ipwho.is', 'ipinfo.io', 'zxinc.org', 'ip.sb (仅 IP)', 'ip.3322.net (仅 IP)', '自定义 URL')
    $srcKeys  = @('auto', 'ip-api', 'ipwho', 'ipinfo', 'zxinc', 'ipsb', '3322', 'custom')
    $srcIdx = [Math]::Max(0, [array]::IndexOf($srcKeys, [string]$c.source))
    $cbSource = New-DarkCombo 160 34 200 $srcItems $srcIdx
    $g2.Controls.Add($cbSource)
    $cbProxy = New-DarkCombo 380 34 100 @('跟随系统', '直连', '自定义') ([Math]::Max(0, [array]::IndexOf(@('system', 'direct', 'custom'), [string]$c.proxyMode)))
    $g2.Controls.Add($cbProxy)

    $g2.Controls.Add((New-FieldLabel '自定义 URL' 14 68))
    $txUrl = New-DarkText 160 66 320 ([string]$c.customUrl)
    $g2.Controls.Add($txUrl)
    $g2.Controls.Add((New-FieldLabel '代理地址' 14 100))
    $txProxy = New-DarkText 160 98 320 ([string]$c.proxyAddress)
    $g2.Controls.Add($txProxy)
    $scroll.Controls.Add($g2)
    $y += $g2.Height + 10

    # --- 显示 ---
    $g3 = New-GroupPanel -Title '显示内容' -Y $y -Height 160 -Theme $theme
    $ckCountry = New-DarkCheck '国家 / 地区' 14 34 130 ([bool]$c.showCountry)
    $ckIp      = New-DarkCheck 'IP 地址' 150 34 110 ([bool]$c.showIp)
    $ckCity    = New-DarkCheck '城市 / 地区' 266 34 130 ([bool]$c.showCity)
    $ckIsp     = New-DarkCheck '运营商' 400 34 90 ([bool]$c.showIsp)
    foreach ($x in @($ckCountry, $ckIp, $ckCity, $ckIsp)) { $g3.Controls.Add($x) }
    $ckCount = New-DarkCheck '刷新倒计时' 14 60 130 ([bool]$c.showCountdown)
    $ckUpd   = New-DarkCheck '更新时间' 150 60 110 ([bool]$c.showUpdatedAt)
    foreach ($x in @($ckCount, $ckUpd)) { $g3.Controls.Add($x) }

    $g3.Controls.Add((New-FieldLabel '徽标样式' 14 92 80))
    $cbBadge = New-DarkCombo 96 90 120 @('色块徽标', 'Emoji 旗帜', '不显示') ([Math]::Max(0, [array]::IndexOf(@('badge', 'emoji', 'none'), [string]$c.badgeMode)))
    $g3.Controls.Add($cbBadge)
    $g3.Controls.Add((New-FieldLabel '国家名语言' 232 92 80))
    $cbLang = New-DarkCombo 312 90 90 @('中文', '英文') ([Math]::Max(0, [array]::IndexOf(@('zh', 'en'), [string]$c.countryLanguage)))
    $g3.Controls.Add($cbLang)
    $g3.Controls.Add((New-FieldLabel '双击行为' 232 124 80))
    $cbDbl = New-DarkCombo 312 122 170 @('立即刷新', '打开设置', '隐藏到托盘') ([Math]::Max(0, [array]::IndexOf(@('refresh', 'settings', 'hide'), [string]$c.doubleClickAction)))
    $g3.Controls.Add($cbDbl)
    $g3.Controls.Add((New-FieldLabel '主题' 14 124 60))
    $cbTheme = New-DarkCombo 96 122 120 @('深色', '浅色') ([Math]::Max(0, [array]::IndexOf(@('dark', 'light'), [string]$c.theme)))
    $g3.Controls.Add($cbTheme)
    $scroll.Controls.Add($g3)
    $y += $g3.Height + 10

    # --- 窗口 ---
    $g4 = New-GroupPanel -Title '窗口' -Y $y -Height 208 -Theme $theme
    $ckTop  = New-DarkCheck '总在最前' 14 34 110 ([bool]$c.alwaysOnTop)
    $ckLock = New-DarkCheck '锁定位置' 140 34 110 ([bool]$c.locked)
    $ckSnap = New-DarkCheck '贴边吸附' 266 34 110 ([bool]$c.snapToEdge)
    $ckRem  = New-DarkCheck '记住位置' 390 34 100 ([bool]$c.rememberPosition)
    foreach ($x in @($ckTop, $ckLock, $ckSnap, $ckRem)) { $g4.Controls.Add($x) }
    $ckThrough = New-DarkCheck '鼠标穿透（只能从托盘菜单取消）' 14 60 300 ([bool]$c.mouseThrough)
    $g4.Controls.Add($ckThrough)

    $g4.Controls.Add((New-FieldLabel '整体尺寸' 14 92 70))
    $cbSize = New-DarkCombo 96 90 110 @('小', '中', '大') ([Math]::Max(0, [array]::IndexOf(@('small', 'medium', 'large'), [string]$c.size)))
    $g4.Controls.Add($cbSize)
    $ckAuto = New-DarkCheck '开机自动启动' 232 90 180 ([bool]$c.autoStart)
    $g4.Controls.Add($ckAuto)

    $g4.Controls.Add((New-FieldLabel '不透明度' 14 126 70))
    $tbOpacity = New-Object System.Windows.Forms.TrackBar
    $tbOpacity.Minimum = 30; $tbOpacity.Maximum = 100
    $tbOpacity.Value = [int]([double]$c.opacity * 100)
    $tbOpacity.TickFrequency = 10
    $tbOpacity.Location = New-Object System.Drawing.Point(96, 122)
    $tbOpacity.Size = New-Object System.Drawing.Size(280, 30)
    $g4.Controls.Add($tbOpacity)
    $lblOpacity = New-FieldLabel ("{0}%" -f $tbOpacity.Value) 386 126 60
    $g4.Controls.Add($lblOpacity)
    $tbOpacity.Add_Scroll({ param($s, $e) $script:lblOpacity.Text = ("{0}%" -f $s.Value) })

    $g4.Controls.Add((New-FieldLabel '字体缩放' 14 164 70))
    $tbFont = New-Object System.Windows.Forms.TrackBar
    $tbFont.Minimum = 80; $tbFont.Maximum = 160
    $tbFont.Value = [int]([double]$c.fontScale * 100)
    $tbFont.TickFrequency = 10
    $tbFont.Location = New-Object System.Drawing.Point(96, 160)
    $tbFont.Size = New-Object System.Drawing.Size(280, 30)
    $g4.Controls.Add($tbFont)
    $lblFont = New-FieldLabel ("{0}%" -f $tbFont.Value) 386 164 60
    $g4.Controls.Add($lblFont)
    $tbFont.Add_Scroll({ param($s, $e) $script:lblFont.Text = ("{0}%" -f $s.Value) })
    $scroll.Controls.Add($g4)
    $y += $g4.Height + 6

    # --- 底部按钮 ---
    $btnTest = New-Object System.Windows.Forms.Button
    $btnTest.Text = '测试连接'
    $btnTest.Location = New-Object System.Drawing.Point(16, 656)
    $btnTest.Size = New-Object System.Drawing.Size(96, 30)
    $dlg.Controls.Add($btnTest)

    $lblTest = New-Object System.Windows.Forms.Label
    $lblTest.Text = ''
    $lblTest.Location = New-Object System.Drawing.Point(120, 660)
    $lblTest.Size = New-Object System.Drawing.Size(150, 24)
    $lblTest.TextAlign = 'MiddleLeft'
    $dlg.Controls.Add($lblTest)

    $btnReset = New-Object System.Windows.Forms.Button
    $btnReset.Text = '恢复默认'
    $btnReset.Location = New-Object System.Drawing.Point(276, 656)
    $btnReset.Size = New-Object System.Drawing.Size(88, 30)
    $dlg.Controls.Add($btnReset)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = '取消'
    $btnCancel.Location = New-Object System.Drawing.Point(370, 656)
    $btnCancel.Size = New-Object System.Drawing.Size(64, 30)
    $dlg.Controls.Add($btnCancel)

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = '保存并应用'
    $btnSave.Location = New-Object System.Drawing.Point(440, 656)
    $btnSave.Size = New-Object System.Drawing.Size(64, 30)
    $dlg.Controls.Add($btnSave)
    $btnSave.DialogResult = [System.Windows.Forms.DialogResult]::OK

    # 事件处理器只能访问脚本作用域的变量, 把用到的控件引用统一提升到脚本作用域
    $promote = @('dlg', 'srcKeys', 'cbSource', 'txUrl', 'cbProxy', 'txProxy', 'nInterval', 'nTimeout',
                 'ckStart', 'ckFallback', 'ckCountry', 'ckIp', 'ckCity', 'ckIsp', 'ckCount', 'ckUpd',
                 'cbBadge', 'cbLang', 'cbDbl', 'cbTheme', 'ckTop', 'ckLock', 'ckSnap', 'ckRem',
                 'ckThrough', 'cbSize', 'ckAuto', 'tbOpacity', 'tbFont', 'lblOpacity', 'lblFont',
                 'btnTest', 'lblTest', 'btnSave', 'btnReset', 'btnCancel')
    foreach ($vn in $promote) {
        $vv = Get-Variable -Name $vn -ValueOnly -ErrorAction SilentlyContinue
        Set-Variable -Name $vn -Value $vv -Scope Script
    }

    Apply-ThemeRecursive -Control $dlg -Theme $theme
    $lblTest.ForeColor = C $theme.Muted

    # 自定义 URL 只在选中自定义时可用
    $syncUrlState = {
        $isCustom = ($srcKeys[$cbSource.SelectedIndex] -eq 'custom')
        $txUrl.Enabled = $isCustom
        $txProxy.Enabled = ($cbProxy.SelectedIndex -eq 2)
    }
    $cbSource.Add_SelectedIndexChanged($syncUrlState)
    $cbProxy.Add_SelectedIndexChanged($syncUrlState)
    & $syncUrlState

    $script:UI.Dlg = $dlg

    $btnCancel.Add_Click({ $script:UI.Dlg.Close() })

    $btnReset.Add_Click({
        $d = Get-DefaultConfig
        $nInterval.Value = $d.intervalSeconds; $nTimeout.Value = $d.timeoutSeconds
        $ckStart.Checked = $d.refreshOnStart; $ckFallback.Checked = $d.autoFallbackProxy
        $cbSource.SelectedIndex = 0; $txUrl.Text = ''; $cbProxy.SelectedIndex = 0; $txProxy.Text = ''
        $ckCountry.Checked = $d.showCountry; $ckIp.Checked = $d.showIp; $ckCity.Checked = $d.showCity; $ckIsp.Checked = $d.showIsp
        $ckCount.Checked = $d.showCountdown; $ckUpd.Checked = $d.showUpdatedAt
        $cbBadge.SelectedIndex = 0; $cbLang.SelectedIndex = 0; $cbDbl.SelectedIndex = 0; $cbTheme.SelectedIndex = 0
        $ckTop.Checked = $d.alwaysOnTop; $ckLock.Checked = $d.locked; $ckSnap.Checked = $d.snapToEdge; $ckRem.Checked = $d.rememberPosition
        $ckThrough.Checked = $false; $cbSize.SelectedIndex = 1; $ckAuto.Checked = $false
        $tbOpacity.Value = [int]($d.opacity * 100); $tbFont.Value = [int]($d.fontScale * 100)
        $lblOpacity.Text = "$($tbOpacity.Value)%"; $lblFont.Text = "$($tbFont.Value)%"
    })

    $btnTest.Add_Click({
        $btnTest.Enabled = $false
        $lblTest.Text = '测试中…'
        $lblTest.ForeColor = C $script:Palette.Accent
        $snapshot = $script:Cfg.Clone()
        $snapshot.defaultUrl = ''
        $script:DlgTest = @{ Ps = $null; Handle = $null }
        $job = @{ Cfg = $snapshot; Endpoints = $script:Endpoints; IsoText = $script:IsoText }
        $ps = [PowerShell]::Create()
        $ps.Runspace = $script:Runspace
        [void]$ps.AddScript($script:WorkerText).AddArgument($job)
        $script:DlgTest = @{ Ps = $ps; Handle = $ps.BeginInvoke() }
        $timer.Start()
    })

    $btnSave.Add_Click({
        Set-CfgValue -Key 'intervalSeconds' -Value ([int]$nInterval.Value) -NoSave
        Set-CfgValue -Key 'timeoutSeconds' -Value ([int]$nTimeout.Value) -NoSave
        Set-CfgValue -Key 'refreshOnStart' -Value ([bool]$ckStart.Checked) -NoSave
        Set-CfgValue -Key 'autoFallbackProxy' -Value ([bool]$ckFallback.Checked) -NoSave
        Set-CfgValue -Key 'source' -Value $srcKeys[$cbSource.SelectedIndex] -NoSave
        Set-CfgValue -Key 'customUrl' -Value ([string]$txUrl.Text).Trim() -NoSave
        Set-CfgValue -Key 'proxyMode' -Value @('system', 'direct', 'custom')[$cbProxy.SelectedIndex] -NoSave
        Set-CfgValue -Key 'proxyAddress' -Value ([string]$txProxy.Text).Trim() -NoSave
        Set-CfgValue -Key 'showCountry' -Value ([bool]$ckCountry.Checked) -NoSave
        Set-CfgValue -Key 'showIp' -Value ([bool]$ckIp.Checked) -NoSave
        Set-CfgValue -Key 'showCity' -Value ([bool]$ckCity.Checked) -NoSave
        Set-CfgValue -Key 'showIsp' -Value ([bool]$ckIsp.Checked) -NoSave
        Set-CfgValue -Key 'showCountdown' -Value ([bool]$ckCount.Checked) -NoSave
        Set-CfgValue -Key 'showUpdatedAt' -Value ([bool]$ckUpd.Checked) -NoSave
        Set-CfgValue -Key 'badgeMode' -Value @('badge', 'emoji', 'none')[$cbBadge.SelectedIndex] -NoSave
        Set-CfgValue -Key 'countryLanguage' -Value @('zh', 'en')[$cbLang.SelectedIndex] -NoSave
        Set-CfgValue -Key 'doubleClickAction' -Value @('refresh', 'settings', 'hide')[$cbDbl.SelectedIndex] -NoSave
        Set-CfgValue -Key 'theme' -Value @('dark', 'light')[$cbTheme.SelectedIndex] -NoSave
        Set-CfgValue -Key 'alwaysOnTop' -Value ([bool]$ckTop.Checked) -NoSave
        Set-CfgValue -Key 'locked' -Value ([bool]$ckLock.Checked) -NoSave
        Set-CfgValue -Key 'snapToEdge' -Value ([bool]$ckSnap.Checked) -NoSave
        Set-CfgValue -Key 'rememberPosition' -Value ([bool]$ckRem.Checked) -NoSave
        Set-CfgValue -Key 'mouseThrough' -Value ([bool]$ckThrough.Checked) -NoSave
        Set-CfgValue -Key 'size' -Value @('small', 'medium', 'large')[$cbSize.SelectedIndex] -NoSave
        Set-CfgValue -Key 'opacity' -Value ([double]$tbOpacity.Value / 100.0) -NoSave
        Set-CfgValue -Key 'fontScale' -Value ([double]$tbFont.Value / 100.0) -NoSave
        $script:Cfg['autoStart'] = [bool]$ckAuto.Checked
        Set-AutoStart -Enabled ([bool]$ckAuto.Checked) | Out-Null
        Save-Config
        Apply-Config
        Start-Refresh
    })

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 250
    Set-Variable -Name timer -Value $timer -Scope Script
    $timer.Add_Tick({
        if (-not $script:DlgTest -or -not $script:DlgTest.Handle) { $timer.Stop(); return }
        if (-not $script:DlgTest.Handle.IsCompleted) { return }
        $timer.Stop()
        $res = Complete-LookupAsync -Async $script:DlgTest
        $script:DlgTest = $null
        $btnTest.Enabled = $true
        if ($res -and $res.Ok) {
            $lblTest.Text = "成功 · $($res.Ip) $($res.Country)"
            $lblTest.ForeColor = C '#3DD68C'
        } else {
            $msg = if ($res) { [string]$res.Error } else { '未知错误' }
            if ($msg.Length -gt 40) { $msg = $msg.Substring(0, 40) + '…' }
            $lblTest.Text = "失败 · $msg"
            $lblTest.ForeColor = C $script:Palette.Err
        }
    })
    $dlg.Add_FormClosed({
        try { $timer.Stop(); $timer.Dispose() } catch { }
        if ($script:DlgTest -and $script:DlgTest.Ps) { try { $script:DlgTest.Ps.Dispose() } catch { } }
        $script:DlgTest = $null
        $script:UI.Dlg = $null
    })
    $dlg.Add_KeyDown({
        param($s, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) { $dlg.Close() }
    })

    if ($script:PreviewMode -or $script:AutoCloseDlg) {
        $dlg.Show()
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 500
        [System.Windows.Forms.Application]::DoEvents()
        if ($script:DlgHook) { & $script:DlgHook $dlg }
        if ($script:PreviewDir) {
            $pb = New-Object System.Drawing.Bitmap($dlg.Width, $dlg.Height)
            $dlg.DrawToBitmap($pb, (New-Object System.Drawing.Rectangle(0, 0, $dlg.Width, $dlg.Height)))
            $pp = Join-Path $script:PreviewDir 'preview-settings.png'
            $pb.Save($pp, [System.Drawing.Imaging.ImageFormat]::Png)
            Write-Host "已生成: $pp  ($($pb.Width)x$($pb.Height))"
            $pb.Dispose()
        }
        $dlg.Close()
        return
    }

    [void]$dlg.ShowDialog($script:UI.Form)
}

# ---------------------------------------------------------------------------
# 14. 主窗口构建
# ---------------------------------------------------------------------------
function New-AppIcon {
    $bmp = New-Object System.Drawing.Bitmap 32, 32
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.Clear([System.Drawing.Color]::Transparent)
    $bg = New-Object System.Drawing.SolidBrush ([System.Drawing.ColorTranslator]::FromHtml('#2F6FED'))
    $g.FillEllipse($bg, 1, 1, 30, 30)
    $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::White), 2.0
    $g.DrawEllipse($pen, 8, 8, 16, 16)
    $g.DrawLine($pen, 16, 4, 16, 28)
    $g.DrawArc($pen, 4, 11, 24, 12, 0, 180)
    $g.DrawArc($pen, 4, 15, 24, 12, 180, 180)
    $bg.Dispose(); $pen.Dispose(); $g.Dispose()
    $hicon = $bmp.GetHicon()
    $icon = [System.Drawing.Icon]::FromHandle($hicon)
    $bmp.Dispose()
    return $icon
}

function New-MainWindow {
    $script:UI = @{}

    $f = New-Object IpmBufferedForm
    $f.FormBorderStyle = 'None'
    $f.ShowInTaskbar = $false
    $f.StartPosition = 'Manual'
    $f.Text = $script:AppName
    $f.KeyPreview = $true
    $f.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)
    $script:UI.Form = $f
    # 事件处理器只能访问脚本作用域的变量, 这里把局部变量提升上去
    Set-Variable -Name f -Value $f -Scope Script

    $script:UI.Tip = New-Object System.Windows.Forms.ToolTip
    $script:UI.Tip.ShowAlways = $true
    $script:UI.Tip.InitialDelay = 350
    $script:UI.Tip.ReshowDelay = 100
    $script:UI.Tip.AutoPopDelay = 20000

    $script:UI.Menu = Build-Menu

    # 托盘
    $tray = New-Object System.Windows.Forms.NotifyIcon
    $tray.Icon = New-AppIcon
    $tray.Text = $script:AppName
    $tray.ContextMenuStrip = Build-TrayMenu
    $tray.Add_MouseDoubleClick({ param($s, $e) Show-Bar })
    $script:UI.Tray = $tray
    $tray.Visible = $true

    # 计时器
    $script:TickTimer = New-Object System.Windows.Forms.Timer
    $script:TickTimer.Interval = 1000
    $script:TickTimer.Add_Tick({
        Invoke-Guarded -Name 'TickTimer' -Body {
            if (-not $script:Busy -and $script:NextAt -and (Get-Date) -ge $script:NextAt) { Start-Refresh }
            elseif (-not $script:Busy) { $script:UI.Form.Invalidate() }
        }
    })

    $script:SpinTimer = New-Object System.Windows.Forms.Timer
    $script:SpinTimer.Interval = 90
    $script:SpinTimer.Add_Tick({
        Invoke-Guarded -Name 'SpinTimer' -Body {
            $script:SpinAngle = ($script:SpinAngle + 30) % 360
            $script:UI.Form.Invalidate()
        }
    })

    $script:PollTimer = New-Object System.Windows.Forms.Timer
    $script:PollTimer.Interval = 200
    $script:PollTimer.Add_Tick({
        Invoke-Guarded -Name 'PollTimer' -Body {
            if (-not $script:Async -or -not $script:Async.Handle.IsCompleted) { return }
            $script:PollTimer.Stop()
            $async = $script:Async
            $script:Async = $null
            $res = Complete-LookupAsync -Async $async
            Apply-Result -Res $res
        }
    })

    # 事件
    $f.Add_Paint({
        param($s, $e)
        try { Invoke-BarPaint -G $e.Graphics }
        catch {
            if (-not $script:PaintErrorLogged) {
                $script:PaintErrorLogged = $true
                $pos = ''
                try { $pos = ($_.InvocationInfo.PositionMessage -replace "`r?`n", ' ') } catch { }
                Write-Log ("绘制异常: {0} :: {1}" -f $_.Exception.Message, $pos)
                Write-Log ("绘制堆栈: {0}" -f (($_.ScriptStackTrace -replace "`r?`n", ' <- ')))
            }
            try {
                $e.Graphics.Clear([System.Drawing.Color]::FromArgb(27, 29, 33))
                $fnt = New-Font 'Segoe UI' 11 Regular
                $br = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 107, 94))
                $e.Graphics.DrawString('绘制异常, 详见 monitor.log', $fnt, $br, 10, 10)
                $br.Dispose(); $fnt.Dispose()
            } catch { }
        }
    })

    $f.Add_MouseDown({
        param($s, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left -and -not $script:Cfg.locked) {
            $script:DragMoved = $false
            $script:DragOrigin = [System.Windows.Forms.Cursor]::Position
            $script:FormOrigin = $f.Location
        }
    })
    $f.Add_MouseMove({
        param($s, $e)
        Invoke-Guarded -Name 'MouseMove' -Body {
            $L = $script:Layout
            $hover = $L -and $L.Btn.Contains([single]$e.X, [single]$e.Y)
            if ($hover -ne $script:HoverBtn) { $script:HoverBtn = $hover; $f.Invalidate() }
            if ($script:DragOrigin -and $e.Button -eq [System.Windows.Forms.MouseButtons]::Left -and -not $script:Cfg.locked) {
                $cur = [System.Windows.Forms.Cursor]::Position
                $dx = $cur.X - $script:DragOrigin.X
                $dy = $cur.Y - $script:DragOrigin.Y
                if ([Math]::Abs($dx) + [Math]::Abs($dy) -gt 4) { $script:DragMoved = $true }
                if ($script:DragMoved) { $f.Location = [System.Drawing.Point]::new(($script:FormOrigin.X + $dx), ($script:FormOrigin.Y + $dy)) }
            }
        }
    })
    $f.Add_MouseUp({
        param($s, $e)
        Invoke-Guarded -Name 'MouseUp' -Body {
            if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
                $wasDrag = $script:DragMoved
                $script:DragOrigin = $null
                $script:FormOrigin = $null
                if (-not $wasDrag) { Start-Refresh }
                elseif ([bool]$script:Cfg.snapToEdge) { Snap-ToEdge }
                elseif ([bool]$script:Cfg.rememberPosition) { Save-Position }
            }
        }
    })
    $f.Add_MouseClick({
        param($s, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
            Invoke-Guarded -Name 'MenuShow' -Body {
                $script:UI.Menu.Show($f, [System.Drawing.Point]::new($e.X, $e.Y))
            }
        }
    })
    $f.Add_MouseDoubleClick({
        param($s, $e)
        if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }
        switch ([string]$script:Cfg.doubleClickAction) {
            'settings' { Show-SettingsDialog }
            'hide'     { $f.Hide() }
            default    { Start-Refresh }
        }
    })
    $f.Add_MouseEnter({ $script:UI.Tip.Active = $true })
    $f.Add_KeyDown({
        param($s, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::F5) { Start-Refresh }
        elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) { $f.Hide() }
        elseif ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::O) { Show-SettingsDialog }
    })
    $f.Add_FormClosing({
        param($s, $e)
        $e.Cancel = $true
        $f.Hide()
    })

    $script:UI.TrayMenu = $tray.ContextMenuStrip
    return $f
}

function Build-TrayMenu {
    $m = New-Object System.Windows.Forms.ContextMenuStrip
    $m.ShowImageMargin = $false
    $m.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)
    [void]$m.Items.Add((New-ActionItem '显示 / 隐藏通知条' { if ($script:UI.Form.Visible) { $script:UI.Form.Hide() } else { Show-Bar } }))
    [void]$m.Items.Add((New-ActionItem '立即刷新' { Start-Refresh }))
    [void]$m.Items.Add((New-ActionItem '复制 IP 地址' { if ($script:Result.Ip) { try { [System.Windows.Forms.Clipboard]::SetText($script:Result.Ip) } catch { } } }))
    [void]$m.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    [void]$m.Items.Add((New-ActionItem '设置…' { Show-SettingsDialog }))
    [void]$m.Items.Add((New-ActionItem '取消鼠标穿透' {
        $script:Cfg.mouseThrough = $false
        Save-Config
        Apply-Config
    }))
    [void]$m.Items.Add((New-ActionItem '重置位置到右上角' { Move-ToDefault }))
    [void]$m.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    [void]$m.Items.Add((New-ActionItem '退出' { Exit-App }))
    return $m
}

function Show-Bar {
    $f = $script:UI.Form
    if (-not $f.Visible) { $f.Show() }
    $f.TopMost = [bool]$script:Cfg.alwaysOnTop
    $f.Activate()
    $f.Invalidate()
}

function Snap-ToEdge {
    $f = $script:UI.Form
    $wa = [System.Windows.Forms.Screen]::FromPoint($f.Location).WorkingArea
    $margin = 6
    $x = $f.Left; $y = $f.Top
    if ([Math]::Abs($f.Left - $wa.Left) -lt 24) { $x = $wa.Left + $margin }
    if ([Math]::Abs($f.Right - $wa.Right) -lt 24) { $x = $wa.Right - $f.Width - $margin }
    if ([Math]::Abs($f.Top - $wa.Top) -lt 24) { $y = $wa.Top + $margin }
    if ([Math]::Abs($f.Bottom - $wa.Bottom) -lt 24) { $y = $wa.Bottom - $f.Height - $margin }
    $f.Location = New-Object System.Drawing.Point($x, $y)
    if ([bool]$script:Cfg.rememberPosition) { Save-Position }
}

function Save-Position {
    $f = $script:UI.Form
    $script:Cfg.posX = $f.Left
    $script:Cfg.posY = $f.Top
    Save-Config
}

function Move-ToDefault {
    $f = $script:UI.Form
    $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $f.Location = New-Object System.Drawing.Point(($wa.Right - $f.Width - 24), ($wa.Top + 24))
    Save-Position
}

function Restore-Position {
    $f = $script:UI.Form
    $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $x = [int]$script:Cfg.posX
    $y = [int]$script:Cfg.posY
    $ok = $false
    if ($x -ge 0 -and $y -ge 0) {
        foreach ($sc in [System.Windows.Forms.Screen]::AllScreens) {
            if ($sc.WorkingArea.Contains($x, $y)) { $ok = $true; break }
        }
    }
    if ($ok -and [bool]$script:Cfg.rememberPosition) {
        $f.Location = New-Object System.Drawing.Point($x, $y)
    } else {
        Move-ToDefault
    }
}

function Exit-App {
    try { Save-Position } catch { }
    try { $script:UI.Tray.Visible = $false; $script:UI.Tray.Dispose() } catch { }
    try { if ($script:Mutex) { $script:Mutex.ReleaseMutex() } } catch { }
    Write-Log '程序退出'
    [System.Windows.Forms.Application]::Exit()
    [Environment]::Exit(0)
}

# ---------------------------------------------------------------------------
# 15. 自检 / 预览 模式
# ---------------------------------------------------------------------------
function Invoke-SelfTest {
    Write-Host "== $($script:AppName) 自检 ==" -ForegroundColor Cyan
    Write-Host "配置: $($script:CfgPath)"
    $cfg = $script:Cfg.Clone()
    if ($Source) { $cfg.source = $Source }
    Write-Host ("数据源={0}  代理={1}  超时={2}s" -f $cfg.source, $cfg.proxyMode, $cfg.timeoutSeconds)
    $job = @{ Cfg = $cfg; Endpoints = $script:Endpoints; IsoText = $script:IsoText }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $res = Invoke-GeoLookup -Job $job
    $sw.Stop()
    if ($res.Ok) {
        Write-Host ("[成功] 用时 {0} ms (总 {1} ms)" -f $res.Elapsed, $sw.ElapsedMilliseconds) -ForegroundColor Green
        Write-Host ("  IP      : {0}" -f $res.Ip)
        Write-Host ("  国家/地区: {0}  ({1})" -f $res.Country, $res.Cc)
        Write-Host ("  省/市   : {0} / {1}" -f $res.Region, $res.City)
        Write-Host ("  运营商  : {0}" -f $res.Isp)
        Write-Host ("  时区    : {0}" -f $res.Tz)
        Write-Host ("  数据源  : {0}" -f $res.Source)
    } else {
        Write-Host "[失败] $($res.Error)" -ForegroundColor Red
    }
}

function Test-BitmapBlank {
    param($Bmp)
    $first = $Bmp.GetPixel(2, 2)
    for ($x = 0; $x -lt $Bmp.Width; $x += 7) {
        for ($y = 0; $y -lt $Bmp.Height; $y += 5) {
            $p = $Bmp.GetPixel($x, $y)
            if ($p.ToArgb() -ne $first.ToArgb()) { return $false }
        }
    }
    return $true
}

function Invoke-Preview {
    param([string]$Dir)
    if (-not (Test-Path $Dir)) { New-Item -ItemType Directory -Path $Dir -Force | Out-Null }
    $script:Result = @{ Ok = $true; Ip = '123.202.92.53'; Cc = 'HK'; Country = '中国香港'
                        Region = '葵青区'; City = '葵涌'; Isp = 'Hong Kong Broadband Network Ltd'
                        Org = ''; Tz = 'Asia/Hong_Kong'; Source = 'ip-api.com'; At = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                        Elapsed = 436; Cached = $false }
    $script:Cfg.showCity = $true
    $script:Cfg.showIsp = $true

    $f = New-MainWindow
    $f.CreateControl() | Out-Null
    $null = $f.Handle
    try {
        $g = $f.CreateGraphics()
        $script:DpiScale = [double]$g.DpiY / 96.0
        $g.Dispose()
    } catch { $script:DpiScale = 1.0 }
    Apply-Config
    [System.Windows.Forms.Application]::DoEvents()

    $bmp1 = New-Object System.Drawing.Bitmap($f.Width, $f.Height)
    $f.DrawToBitmap($bmp1, (New-Object System.Drawing.Rectangle(0, 0, $f.Width, $f.Height)))
    if (Test-BitmapBlank $bmp1) {
        $f.Show()
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 600
        [System.Windows.Forms.Application]::DoEvents()
        $bmp1.Dispose()
        $bmp1 = New-Object System.Drawing.Bitmap($f.Width, $f.Height)
        $f.DrawToBitmap($bmp1, (New-Object System.Drawing.Rectangle(0, 0, $f.Width, $f.Height)))
    }
    $p1 = Join-Path $Dir 'preview-bar.png'
    $bmp1.Save($p1, [System.Drawing.Imaging.ImageFormat]::Png)
    Write-Host "已生成: $p1  ($($bmp1.Width)x$($bmp1.Height))"
    $bmp1.Dispose()

    # 亮色主题 + 无徽标
    $script:Cfg.theme = 'light'
    Apply-Config
    [System.Windows.Forms.Application]::DoEvents()
    $bmp2 = New-Object System.Drawing.Bitmap($f.Width, $f.Height)
    $f.DrawToBitmap($bmp2, (New-Object System.Drawing.Rectangle(0, 0, $f.Width, $f.Height)))
    $p2 = Join-Path $Dir 'preview-bar-light.png'
    $bmp2.Save($p2, [System.Drawing.Imaging.ImageFormat]::Png)
    Write-Host "已生成: $p2"
    $bmp2.Dispose()

    # 设置窗口预览
    $script:Cfg.theme = 'dark'
    Apply-Config
    $script:PreviewMode = $true
    $script:PreviewDir = $Dir
    Show-SettingsDialog
    $script:PreviewMode = $false

    try { $f.Close() } catch { }
    Write-Host '预览完成'
}

function Get-AllControls {
    param($Root)
    $out = @()
    foreach ($c in $Root.Controls) {
        $out += $c
        if ($c.HasChildren) { $out += Get-AllControls $c }
    }
    return $out
}

# 端到端烟雾测试: 建窗口 -> 异步刷新 -> 右键菜单 -> 设置窗口 -> 触发保存处理器
function Invoke-SmokeTest {
    Write-Host "== $($script:AppName) 烟雾测试 ==" -ForegroundColor Cyan
    Write-Host "数据目录: $($script:DataDir)"
    $form = New-MainWindow
    $form.CreateControl() | Out-Null
    $null = $form.Handle
    try { $g = $form.CreateGraphics(); $script:DpiScale = [double]$g.DpiY / 96.0; $g.Dispose() } catch { }
    Apply-Config
    Restore-Position
    $form.Show()
    Write-Host ("[1] 主窗口 {0}x{1}  句柄有效={2}  托盘图标={3}" -f `
        $form.Width, $form.Height, ($form.Handle -ne [IntPtr]::Zero), $script:UI.Tray.Visible)

    Start-Refresh
    $deadline = (Get-Date).AddSeconds(30)
    while ($script:Busy -and (Get-Date) -lt $deadline) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 100
    }
    $r = $script:Result
    Write-Host ("[2] 刷新完成: Ok={0}  IP={1}  国家={2}  IP行文本长度={3}" -f $r.Ok, $r.Ip, $r.Country, "$($r.Ip)".Length)
    Write-Host ("[3] 结果缓存文件存在: {0}" -f (Test-Path $script:CachePath))
    $tl = $script:UI.Tip.GetToolTip($form)
    Write-Host ("[4] 悬浮提示行数: {0}" -f (@($tl -split "`n").Count))

    $script:TickTimer.Start()
    for ($i = 0; $i -lt 5; $i++) { [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 120 }

    $script:UI.Menu.Show($form, (New-Object System.Drawing.Point(8, 8)))
    [System.Windows.Forms.Application]::DoEvents()
    Start-Sleep -Milliseconds 250
    [System.Windows.Forms.Application]::DoEvents()
    $menuCount = $script:UI.Menu.Items.Count
    $subCount = (@(Get-MenuItemsRecursive $script:UI.Menu.Items).Count)
    $script:UI.Menu.Close()
    Write-Host ("[5] 右键菜单: 顶层 {0} 项, 含子项共 {1} 项" -f $menuCount, $subCount)

    $before = $script:Cfg.intervalSeconds
    $script:AutoCloseDlg = $true
    $script:PreviewDir = $null
    $script:DlgHook = {
        param($d)
        Write-Host ("[6] 设置窗口控件数: {0}" -f (@(Get-AllControls $d).Count))
        $script:nInterval.Value = 90
        $script:ckCity.Checked = $true
        $script:cbSource.SelectedIndex = 1
        $script:btnSave.PerformClick()
        [System.Windows.Forms.Application]::DoEvents()
        Write-Host ("[7] 保存处理器执行: 间隔 {0} -> {1}  显示城市={2}  数据源={3}" -f `
            $before, $script:Cfg.intervalSeconds, $script:Cfg.showCity, $script:Cfg.source)
    }
    Show-SettingsDialog
    $script:AutoCloseDlg = $false
    $script:DlgHook = $null

    $saved = $false
    $savedTxt = ''
    if (Test-Path $script:CfgPath) {
        $savedTxt = (Get-Content $script:CfgPath -Raw -Encoding UTF8)
        $saved = $savedTxt.Contains('"intervalSeconds":  90') -or $savedTxt.Contains('"intervalSeconds": 90')
    }
    Write-Host ("[8] 配置文件已写入且生效: {0}" -f $saved)

    for ($i = 0; $i -lt 6; $i++) { [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 100 }
    try { $script:UI.Tray.Visible = $false; $script:UI.Tray.Dispose() } catch { }
    Write-Host '烟雾测试完成' -ForegroundColor Green
}

# ---------------------------------------------------------------------------
# 16. 入口
# ---------------------------------------------------------------------------
if ($SelfTest) {
    Invoke-SelfTest
    exit 0
}

[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
Initialize-Worker
$script:Palette = Get-Palette -Theme ([string]$script:Cfg.theme)

if ($SmokeTest) {
    Invoke-SmokeTest
    exit 0
}

# 单实例
$script:Mutex = New-Object System.Threading.Mutex($false, $script:MutexName)
if (-not $script:Mutex.WaitOne(0)) {
    Write-Log '已有实例在运行, 本次启动忽略'
    exit 0
}

if ($PreviewPath) {
    Invoke-Preview -Dir $PreviewPath
    exit 0
}

# 读取上次结果, 先显示缓存
$cached = Load-Cache
if ($cached) { $script:Result = $cached }

$form = New-MainWindow
$script:DpiScale = 1.0
try {
    $g = $form.CreateGraphics()
    $script:DpiScale = [double]$g.DpiY / 96.0
    $g.Dispose()
} catch { }
if ($script:DpiScale -le 0) { $script:DpiScale = 1.0 }

$form.CreateControl() | Out-Null
$null = $form.Handle
Apply-Config
Restore-Position
$form.Opacity = [double]$script:Cfg.opacity
$form.Show()
$form.Activate()

# 首次开机自启状态同步
if ($script:Cfg.autoStart -ne (Test-AutoStart)) {
    if ([bool]$script:Cfg.autoStart) { Set-AutoStart -Enabled $true | Out-Null }
}

$script:TickTimer.Start()
Write-Log "启动 v$($script:AppVer)  尺寸=$($form.Width)x$($form.Height)  DPI缩放=$([Math]::Round($script:DpiScale,2))"

if ([bool]$script:Cfg.refreshOnStart) {
    Start-Refresh
    $script:NextAt = (Get-Date).AddSeconds([int]$script:Cfg.intervalSeconds)
} else {
    $script:NextAt = (Get-Date).AddSeconds([int]$script:Cfg.intervalSeconds)
}

[System.Windows.Forms.Application]::Run($form)
