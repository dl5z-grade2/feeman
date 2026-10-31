; ---------------------------------------------------------------------------
; 饭费管理系统 安装脚本（Inno Setup 6）
;
; 构建步骤：
;   1. uv sync
;   2. uv run pyinstaller fee_app.spec --noconfirm --clean
;   3. ISCC.exe installer.iss
;   （或直接运行 build.ps1，上述步骤会自动完成）
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
DefaultDirName={autopf}\{#AppNameEn}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
; 输出
OutputDir=dist
OutputBaseFilename={#AppNameEn}-Setup-{#AppVersion}
Compression=lzma2/max
SolidCompression=yes
; 架构与权限
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
; 外观
WizardStyle=modern
SetupIconFile=assets\app.ico
UninstallDisplayIcon={app}\{#AppExeName}
UninstallDisplayName={#AppName}
; 系统版本
MinVersion=10.0

[Languages]
; 简体中文语言文件随仓库提供（installer\Languages），不依赖 Inno Setup 安装目录里是否有中文
Name: "chinesesimp"; MessagesFile: "installer\Languages\ChineseSimplified.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "associate"; Description: "关联 {#FileAssocExt} 文件（双击直接打开工程）"; GroupDescription: "文件关联"

[Files]
; 主程序与全部依赖
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; 文档
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "LICENSE"; DestDir: "{app}"; Flags: ignoreversion

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
Type: filesandordirs; Name: "{app}\_internal"

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

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    RefreshAssociations();
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usPostUninstall then
    RefreshAssociations();
end;

{ ---------------------------------------------------------------
  初始化：若已存在更高版本的安装，提示覆盖
  --------------------------------------------------------------- }
function InitializeSetup(): Boolean;
begin
  Result := True;
end;
