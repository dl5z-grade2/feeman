param([switch]$NoLaunch)

# bootstrap_runtime.ps1
# 为 FeeMan 准备轻量运行环境：下载 Python 嵌入式运行时 + pip 安装依赖。
# 相比 PyInstaller 单文件打包，体积更小，且运行时自带 Tkinter（GUI 需要）。
$ErrorActionPreference = 'Stop'
$WORK_DIR = Split-Path -Parent $MyInvocation.MyCommand.Path
$RUNTIME = Join-Path $WORK_DIR 'runtime'
$PY_ZIP = Join-Path $RUNTIME 'python_embed.zip'
# FeeMan 运行依赖：openpyxl（Excel 导入导出）
# 注意：tkinter 不通过 pip 安装，它随 runtime 目录预置（见下方“配置 tkinter”）。
$PACKAGES = @('openpyxl')
$MIRRORS = @('https://mirrors.aliyun.com/pypi/simple', 'https://pypi.tuna.tsinghua.edu.cn/simple', 'https://mirrors.ustc.edu.cn/pypi/web/simple')

# 自动测速选最快镜像源
function Select-FastestMirror {
    param([string[]]$Mirrors)
    $results = @()
    foreach ($m in $Mirrors) {
        try {
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            $testUrl = $m.TrimEnd('/') + '/pip/'
            $resp = Invoke-WebRequest -Uri $testUrl -UseBasicParsing -TimeoutSec 10
            $sw.Stop()
            $speed = if ($resp.RawContentLength -gt 0) { $resp.RawContentLength / $sw.Elapsed.TotalSeconds } else { 0 }
            Write-Host ("测速 {0}  ->  {1:N0} KB/s" -f $m, ($speed / 1KB))
            $results += [PSCustomObject]@{ Url = $m; Speed = $speed }
        } catch {
            Write-Host ("测速 {0}  ->  不可用" -f $m)
        }
    }
    if ($results.Count -eq 0) { return $Mirrors }
    return ($results | Sort-Object Speed -Descending | ForEach-Object { $_.Url })
}
Write-Host '正在测速选择最快镜像源...'
$MIRRORS = @(Select-FastestMirror $MIRRORS)
Write-Host ("最快镜像源: {0}" -f $MIRRORS[0])

$pyExe = Join-Path $RUNTIME 'python.exe'
if (-not (Test-Path $pyExe)) {
    Write-Host 'Step 1: 下载 Python 嵌入式运行时 (国内镜像)'
    New-Item -ItemType Directory -Force -Path $RUNTIME | Out-Null
    $pyUrls = @('https://mirrors.huaweicloud.com/python/3.11.9/python-3.11.9-embed-amd64.zip', 'https://registry.npmmirror.com/-/binary/python/3.11.9/python-3.11.9-embed-amd64.zip', 'https://www.python.org/ftp/python/3.11.9/python-3.11.9-embed-amd64.zip')
    $downloaded = $false
    foreach ($url in $pyUrls) {
        try { Write-Host "尝试下载: $url"; Invoke-WebRequest -Uri $url -OutFile $PY_ZIP -UseBasicParsing -TimeoutSec 300; $downloaded = $true; break } catch { Write-Host '下载失败，尝试下一个源...' }
    }
    if (-not $downloaded) { Write-Host 'ERROR: Python 下载失败'; if (-not $NoLaunch) { Read-Host '按回车退出' }; exit 1 }
    Write-Host 'Step 2: 解压 Python'
    Expand-Archive -Path $PY_ZIP -DestinationPath $RUNTIME -Force
    Remove-Item $PY_ZIP -Force
    Write-Host 'Step 3: 启用 site 模块'
    $pth = Join-Path $RUNTIME 'python311._pth'
    if (Test-Path $pth) { (Get-Content $pth) -replace '#import site', 'import site' | Set-Content $pth }
    Write-Host 'Step 4: 安装 pip 并升级'
    $gp = Join-Path $RUNTIME 'get-pip.py'
    foreach ($url in @('https://mirrors.aliyun.com/pypi/get-pip.py', 'https://bootstrap.pypa.io/get-pip.py')) { try { Invoke-WebRequest -Uri $url -OutFile $gp -UseBasicParsing; break } catch { } }
    if (Test-Path $gp) { & $pyExe $gp --no-setuptools --no-wheel -i $MIRRORS[0]; Remove-Item $gp -Force }
    & $pyExe -m pip install --upgrade pip -i $MIRRORS[0]
}

Write-Host '检查并安装依赖包:'
Write-Host $PACKAGES
# 逐包探测，缺哪个装哪个（而不是只看某一个包）。
# 注意：$ErrorActionPreference='Stop' 时，原生命令写 stderr 会被当成终止性错误，
# 而“包未安装”时 python 正好会向 stderr 输出 traceback，所以必须 try/catch 包住。
$missing = @()
foreach ($pkg in $PACKAGES) {
    $mod = if ($pkg -eq 'openpyxl') { 'openpyxl' } else { $pkg }
    $probeOk = $false
    try {
        & $pyExe -c "import $mod" 2>$null
        $probeOk = ($LASTEXITCODE -eq 0)
    } catch {
        $probeOk = $false
    }
    if (-not $probeOk) { $missing += $pkg }
}
if ($missing.Count -gt 0) {
    Write-Host ("缺失依赖: {0}，开始安装..." -f ($missing -join ', '))
    $ok = $false
    foreach ($mirror in $MIRRORS) {
        Write-Host "尝试镜像: $mirror"
        try { & $pyExe -m pip install $missing -i $mirror --progress-bar on } catch { Write-Host "pip 输出: $($_.Exception.Message)" }
        if ($LASTEXITCODE -eq 0) { $ok = $true; break }
    }
    if (-not $ok) { Write-Host 'WARNING: 依赖安装失败，请手动执行:'; Write-Host "$pyExe -m pip install $missing"; if (-not $NoLaunch) { Read-Host '按回车退出' }; exit 1 }
} else { Write-Host '所有依赖包已安装。' }

# 配置 tkinter
# 嵌入式运行时本身不含 Tcl/Tk，本项目的 runtime 目录已随包附带提取好的
# tkinter 组件（_tkinter.pyd、tcl86t.dll、tk86t.dll、tcl/、tkinter/），
# 这里只做完整性校验、设置 Tcl/Tk 库路径，并验证能真正初始化 Tcl/Tk。
Write-Host '配置 tkinter...'

# 必需文件：缺任何一个 tkinter 都无法工作（嵌入式包不自带 _tkinter.pyd）
$tkRequired = @(
    @{ Path = '_tkinter.pyd';               Desc = 'tkinter 扩展模块' },
    @{ Path = 'tcl86t.dll';                 Desc = 'Tcl 运行库' },
    @{ Path = 'tk86t.dll';                  Desc = 'Tk 运行库' },
    @{ Path = 'tkinter\__init__.py';        Desc = 'tkinter 包' },
    @{ Path = 'tcl\tcl8.6\init.tcl';        Desc = 'Tcl 脚本库' },
    @{ Path = 'tcl\tk8.6\tk.tcl';           Desc = 'Tk 脚本库' }
)
$tkMissing = @()
foreach ($item in $tkRequired) {
    if (-not (Test-Path (Join-Path $RUNTIME $item.Path))) { $tkMissing += $item }
}
if ($tkMissing.Count -gt 0) {
    Write-Host 'ERROR: runtime 目录缺少以下 Tkinter 组件，程序无法启动图形界面：' -ForegroundColor Red
    foreach ($m in $tkMissing) { Write-Host ("  - {0}  ({1})" -f $m.Path, $m.Desc) -ForegroundColor Red }
    Write-Host '请重新解压完整的安装包（这些文件随程序一起发布，不需要联网下载）。' -ForegroundColor Yellow
    if (-not $NoLaunch) { Read-Host '按回车退出' }
    exit 1
}

# tk86t.dll / tcl86t.dll 由 _tkinter.pyd 隐式加载，必须让 Windows 能找到它们；
# 把 runtime 目录加入 PATH 并指定 Tcl/Tk 脚本库位置。
$env:PATH = $RUNTIME + ';' + $env:PATH
$env:TCL_LIBRARY = Join-Path $RUNTIME 'tcl\tcl8.6'
$env:TK_LIBRARY = Join-Path $RUNTIME 'tcl\tk8.6'
# 避免受系统级 Tcl/Tk 环境变量干扰
$env:TIX_LIBRARY = Join-Path $RUNTIME 'tcl\tix8.4.3'

# 用 Tcl() 验证（不创建窗口，避免安装时闪一下）
# 同样要 try/catch：失败时 python 会向 stderr 输出 traceback
$tkOk = $false
try {
    & $pyExe -c "import tkinter; t = tkinter.Tcl(); print('Tcl', t.eval('info patchlevel'))" 2>$null
    $tkOk = ($LASTEXITCODE -eq 0)
} catch {
    $tkOk = $false
}
if (-not $tkOk) {
    Write-Host 'ERROR: tkinter 初始化失败。runtime 中的 Python 与 Tcl/Tk 组件版本可能不匹配。' -ForegroundColor Red
    try { Write-Host ('       Python: ' + (& $pyExe -c "import sys; print(sys.version.split()[0])" 2>$null)) } catch { }
    if (-not $NoLaunch) { Read-Host '按回车退出' }
    exit 1
}
Write-Host '所有依赖安装并配置成功。'
Write-Host '正在启动程序: fee_app.py'
if ($NoLaunch) {
    Write-Host '运行环境已就绪。'
} else {
    Start-Process -FilePath $pyExe -ArgumentList ('"' + (Join-Path $WORK_DIR 'fee_app.py') + '"') -WorkingDirectory $WORK_DIR
    Write-Host '程序已启动，本窗口可关闭。'
    Start-Sleep -Seconds 3
}
