; ---------------------------------------------------------------------------
; 饭费管理系统 安装脚本（Inno Setup 6）
;
; 安装布局（嵌入式 Python 运行时 + C++ 启动器）：
;   {app}\FeeMan.exe              C++ 启动器（已内嵌图标）
;   {app}\app.ico                 程序图标（文件关联也用它）
;   {app}\fee_app.py / fee_core.py
;   {app}\bootstrap_runtime.ps1   首次运行下载 Python 3.11.9 embed + openpyxl
;   {app}\bootstrap_runtime.bat   手动重装运行环境的入口
;   {app}\runtime\                预置 Tkinter 组件（本安装包自带，无需联网）：
;       _tkinter.pyd               tkinter C 扩展（官方 embed 包不含，必须随包）
;       tcl86t.dll / tk86t.dll     Tcl / Tk 运行库
;       tcl\tcl8.6\                Tcl 脚本库
;       tcl\tk8.6\                 Tk 脚本库
;       tkinter\                   tkinter 纯 Python 包
;     注意：python.exe / python311.dll / Lib / Scripts 由首次运行下载，不打进安装包。
;
; 构建：
;   powershell -ExecutionPolicy Bypass -File build.ps1
;   （编译 launcher.cpp → 组装 dist\FeeMan → 调用本脚本）
;
; 产物：dist\FeeMan-Setup-{#AppVersion}.exe
; ---------------------------------------------------------------------------

#define AppName        "饭费管理系统"
#define AppNameEn      "FeeMan"
#define AppVersion     "1.0.0"
#define AppPublisher   "FeeMan"
#define AppExeName     "FeeMan.exe"
#define SourceDir      "dist\FeeMan"
#define FileAssocExt   ".cfee"
#define FileAssocName  "饭费工程文件"

; ---------------------------------------------------------------------------
; 编译期自检：缺文件就立刻报错，避免打出“装完跑不起来”的安装包
; ---------------------------------------------------------------------------
#if !FileExists(SourceDir + "\" + AppExeName)
  #error "缺少 " + SourceDir + "\" + AppExeName + "，请先运行 build.ps1 编译启动器"
#endif
#if !FileExists(SourceDir + "\runtime\_tkinter.pyd")
  #error "缺少 " + SourceDir + "\runtime\_tkinter.pyd（tkinter 扩展），请检查 build.ps1 的组装步骤"
#endif
#if !FileExists(SourceDir + "\runtime\tcl86t.dll")
  #error "缺少 " + SourceDir + "\runtime\tcl86t.dll（Tcl 运行库）"
#endif
#if !FileExists(SourceDir + "\runtime\tk86t.dll")
  #error "缺少 " + SourceDir + "\runtime\tk86t.dll（Tk 运行库）"
#endif
#if !FileExists(SourceDir + "\runtime\tcl\tcl8.6\init.tcl")
  #error "缺少 " + SourceDir + "\runtime\tcl\tcl8.6\init.tcl（Tcl 脚本库）"
#endif
#if !FileExists(SourceDir + "\runtime\tcl\tk8.6\tk.tcl")
  #error "缺少 " + SourceDir + "\runtime\tcl\tk8.6\tk.tcl（Tk 脚本库）"
#endif
#if !FileExists(SourceDir + "\runtime\tkinter\__init__.py")
  #error "缺少 " + SourceDir + "\runtime\tkinter\__init__.py（tkinter 包）"
#endif

[Setup]
AppId={{8F3A6C21-5B4D-4E7A-9C1F-2D8B6E4A7F30}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
; 按用户安装，无需管理员
DefaultDirName={localappdata}\Programs\{#AppNameEn}
; 必须装到可写目录：首次运行要把 Python 下载解压进 {app}\runtime。
; 旧版本曾默认装到 Program Files，升级时不能沿用那个目录，否则 bootstrap 无写权限。
UsePreviousAppDir=no
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
OutputDir=dist
OutputBaseFilename={#AppNameEn}-Setup-{#AppVersion}
Compression=lzma2/max
SolidCompression=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
WizardStyle=modern
SetupIconFile=assets\app.ico
UninstallDisplayIcon={app}\{#AppExeName}
UninstallDisplayName={#AppName}
MinVersion=10.0
; 安装目录会加入用户 PATH，需要通知系统环境变量已变化
ChangesEnvironment=yes

[Languages]
Name: "chinesesimp"; MessagesFile: "installer\Languages\ChineseSimplified.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "associate"; Description: "关联 {#FileAssocExt} 文件（双击直接打开工程）"; GroupDescription: "文件关联"

[Files]
; ---- 启动器、源码、文档 ----
Source: "{#SourceDir}\{#AppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\fee_app.py"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\fee_core.py"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\bootstrap_runtime.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\bootstrap_runtime.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\LICENSE"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\app.ico"; DestDir: "{app}"; Flags: ignoreversion

; ---- 预置 Tkinter 组件（随包分发，不联网）----
Source: "{#SourceDir}\runtime\_tkinter.pyd"; DestDir: "{app}\runtime"; Flags: ignoreversion
Source: "{#SourceDir}\runtime\tcl86t.dll"; DestDir: "{app}\runtime"; Flags: ignoreversion
Source: "{#SourceDir}\runtime\tk86t.dll"; DestDir: "{app}\runtime"; Flags: ignoreversion
Source: "{#SourceDir}\runtime\tcl\*"; DestDir: "{app}\runtime\tcl"; \
    Flags: ignoreversion recursesubdirs createallsubdirs; \
    Excludes: "__pycache__,*.pyc,*.pyo"
Source: "{#SourceDir}\runtime\tkinter\*"; DestDir: "{app}\runtime\tkinter"; \
    Flags: ignoreversion recursesubdirs createallsubdirs; \
    Excludes: "__pycache__,*.pyc,*.pyo"

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{group}\卸载 {#AppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[InstallDelete]
; 升级/重装时清掉旧版（PyInstaller 单目录方案）遗留的 _internal，
; 避免旧 DLL 与新的 runtime 方案混在一起
Type: filesandordirs; Name: "{app}\_internal"
Type: files; Name: "{app}\FeeMan_spec.bak"

[Registry]
; ---- .cfee 文件关联（仅当前用户，随 Tasks 可选）----
Root: HKCU; Subkey: "Software\Classes\{#FileAssocExt}"; ValueType: string; \
    ValueName: ""; ValueData: "{#FileAssocName}"; Flags: uninsdeletekey; Tasks: associate
Root: HKCU; Subkey: "Software\Classes\{#FileAssocExt}"; ValueType: string; \
    ValueName: "Content Type"; ValueData: "application/x-cfee"; Tasks: associate
Root: HKCU; Subkey: "Software\Classes\{#FileAssocExt}"; ValueType: string; \
    ValueName: "PerceivedType"; ValueData: "document"; Tasks: associate
Root: HKCU; Subkey: "Software\Classes\{#FileAssocExt}\DefaultIcon"; ValueType: string; \
    ValueName: ""; ValueData: "{app}\app.ico,0"; Tasks: associate
Root: HKCU; Subkey: "Software\Classes\{#FileAssocExt}\shell\open"; ValueType: string; \
    ValueName: ""; ValueData: "用{#AppName}打开"; Tasks: associate
Root: HKCU; Subkey: "Software\Classes\{#FileAssocExt}\shell\open\command"; ValueType: string; \
    ValueName: ""; ValueData: """{app}\{#AppExeName}"" ""%1"""; Tasks: associate

; 应用能力注册（让“打开方式”列表里出现本程序）
Root: HKCU; Subkey: "Software\Classes\Applications\{#AppExeName}"; ValueType: string; \
    ValueName: "FriendlyAppName"; ValueData: "{#AppName}"; Flags: uninsdeletekey; Tasks: associate
Root: HKCU; Subkey: "Software\Classes\Applications\{#AppExeName}\shell\open\command"; ValueType: string; \
    ValueName: ""; ValueData: """{app}\{#AppExeName}"" ""%1"""; Tasks: associate
Root: HKCU; Subkey: "Software\Classes\Applications\{#AppExeName}\SupportedTypes"; ValueType: string; \
    ValueName: "{#FileAssocExt}"; ValueData: ""; Tasks: associate

; 卸载时一并清掉扩展名关联
Root: HKCU; Subkey: "Software\Classes\{#FileAssocExt}"; Flags: uninsdeletekey

[Run]
Filename: "{app}\{#AppExeName}"; Description: "{cm:LaunchProgram,{#AppName}}"; \
    Flags: nowait postinstall skipifsilent

[UninstallDelete]
; 首次运行时下载的 Python、pip、site-packages 不在安装清单里，卸载时一并删除
Type: filesandordirs; Name: "{app}\runtime"
Type: filesandordirs; Name: "{app}\__pycache__"

[Code]
{ ---------------------------------------------------------------
  关联刷新：安装后通知资源管理器，让图标与“打开方式”立即生效
  --------------------------------------------------------------- }
procedure SHChangeNotify(wEventId: Integer; uFlags: Cardinal; dwItem1, dwItem2: Cardinal);
  external 'SHChangeNotify@shell32.dll stdcall';

procedure RefreshAssociations();
begin
  SHChangeNotify($08000000, $0000, 0, 0);  { SHCNE_ASSOCCHANGED, SHCNF_IDLIST }
end;

{ ---------------------------------------------------------------
  用户 PATH：安装时加入安装目录，卸载时移除
  --------------------------------------------------------------- }
const
  EnvironmentKey = 'Environment';

procedure AddToPath();
var
  OrigPath, AppDir: string;
begin
  AppDir := ExpandConstant('{app}');
  if not RegQueryStringValue(HKCU, EnvironmentKey, 'Path', OrigPath) then
    OrigPath := '';

  if Pos(';' + Uppercase(AppDir) + ';', ';' + Uppercase(OrigPath) + ';') = 0 then
  begin
    if OrigPath = '' then
      OrigPath := AppDir
    else if Copy(OrigPath, Length(OrigPath), 1) = ';' then
      OrigPath := OrigPath + AppDir
    else
      OrigPath := OrigPath + ';' + AppDir;
    RegWriteExpandStringValue(HKCU, EnvironmentKey, 'Path', OrigPath);
  end;
end;

procedure RemoveFromPath();
var
  OrigPath, AppDir, Wrapped: string;
  P: Integer;
begin
  if not RegQueryStringValue(HKCU, EnvironmentKey, 'Path', OrigPath) then
    Exit;

  AppDir := ExpandConstant('{app}');
  Wrapped := ';' + OrigPath + ';';
  P := Pos(';' + Uppercase(AppDir) + ';', Uppercase(Wrapped));
  if P = 0 then
    Exit;

  Delete(Wrapped, P, Length(AppDir) + 1);
  if (Length(Wrapped) > 0) and (Wrapped[1] = ';') then
    Delete(Wrapped, 1, 1);
  if (Length(Wrapped) > 0) and (Wrapped[Length(Wrapped)] = ';') then
    Delete(Wrapped, Length(Wrapped), 1);

  if Wrapped = '' then
    RegDeleteValue(HKCU, EnvironmentKey, 'Path')
  else
    RegWriteExpandStringValue(HKCU, EnvironmentKey, 'Path', Wrapped);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    RefreshAssociations();
    AddToPath();
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usPostUninstall then
  begin
    RemoveFromPath();
    RefreshAssociations();
  end;
end;

function InitializeSetup(): Boolean;
begin
  Result := True;
end;
