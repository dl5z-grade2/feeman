; ---------------------------------------------------------------------------
; 饭费管理系统 安装脚本（Inno Setup 6）
;
; 安装布局（嵌入式 Python 运行时 + C++ 启动器）：
;   {app}\FeeMan.exe              C++ 启动器
;   {app}\fee_app.py / fee_core.py
;   {app}\bootstrap_runtime.ps1   首次运行下载 Python 3.11.9 embed + openpyxl
;   {app}\runtime\                预置 Tcl/Tk、tkinter；python.exe 首次运行下载
;
; 构建：
;   powershell -ExecutionPolicy Bypass -File build.ps1
;   （会先编译 launcher.cpp，再调用本脚本）
;
; 产物：dist\FeeMan-Setup-1.0.0.exe
; ---------------------------------------------------------------------------

#define AppName        "饭费管理系统"
#define AppNameEn      "FeeMan"
#define AppVersion     "1.0.0"
#define AppPublisher   "FeeMan"
#define AppExeName     "FeeMan.exe"
#define SourceDir      "dist\FeeMan"
#define FileAssocExt   ".cfee"
#define FileAssocName  "饭费工程文件"

[Setup]
AppId={{8F3A6C21-5B4D-4E7A-9C1F-2D8B6E4A7F30}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
; 按用户安装，无需管理员
DefaultDirName={localappdata}\Programs\{#AppNameEn}
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
; 启动器、Python 源码、引导脚本、预置的 Tcl/Tk/tkinter、文档
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{group}\卸载 {#AppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Registry]
; ---- .cfee 文件关联（仅当前用户，随 Tasks 可选）----
Root: HKCU; Subkey: "Software\Classes\{#FileAssocExt}"; ValueType: string; \
    ValueName: ""; ValueData: "{#FileAssocName}"; Flags: uninsdeletekey; Tasks: associate
Root: HKCU; Subkey: "Software\Classes\{#FileAssocExt}"; ValueType: string; \
    ValueName: "Content Type"; ValueData: "application/x-cfee"; Tasks: associate
Root: HKCU; Subkey: "Software\Classes\{#FileAssocExt}"; ValueType: string; \
    ValueName: "PerceivedType"; ValueData: "document"; Tasks: associate
Root: HKCU; Subkey: "Software\Classes\{#FileAssocExt}\DefaultIcon"; ValueType: string; \
    ValueName: ""; ValueData: "{app}\{#AppExeName},0"; Tasks: associate
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
; 首次运行时下载的 Python、pip、site-packages 及 pycache 不在安装清单里，卸载时一并删除
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