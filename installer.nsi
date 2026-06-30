; AI-Excel-Agent Windows 安装包脚本 (NSIS 3.x)
; 编译：makensis installer.nsi
; 产物：AI-Excel-Agent-Setup-v1.0.exe
;
; 设计：
; - 安装到 %LOCALAPPDATA%\Programs\AI-Excel-Agent（免管理员权限，不污染系统目录）
; - 开始菜单 + 桌面快捷方式
; - 完整卸载器，卸载时保留用户生成的 outputs/data（可勾选清除）
; - 静默支持：/S 参数静默安装

!define APP_NAME "AI-Excel-Agent"
!define APP_VERSION "1.0"
!define APP_PUBLISHER "AI-Excel-Agent"
!define APP_EXE "AI-Excel-Agent.exe"
!define APP_URL "https://github.com/XuYing1128/AI-Excel-Agent"

Unicode true
ManifestDPIAware true

Name "${APP_NAME} ${APP_VERSION}"
OutFile "${APP_NAME}-Setup-v${APP_VERSION}.exe"
InstallDir "$LOCALAPPDATA\Programs\${APP_NAME}"
InstallDirRegKey HKCU "Software\${APP_NAME}" "InstallDir"
RequestExecutionLevel user

; 现代化 UI
!include "MUI2.nsh"
!include "LogicLib.nsh"

!define MUI_ABORTWARNING
!define MUI_FINISHPAGE_RUN "$INSTDIR\${APP_EXE}"
!define MUI_FINISHPAGE_RUN_TEXT "立即启动 ${APP_NAME}"
!define MUI_FINISHPAGE_SHOWREADME ""
!define MUI_FINISHPAGE_SHOWREADME_NOTCHECKED

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_WELCOME
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_UNPAGE_FINISH

!insertmacro MUI_LANGUAGE "SimpChinese"
!insertmacro MUI_LANGUAGE "English"

; ---- 版本信息 ----
VIAddVersionKey "ProductName" "${APP_NAME}"
VIAddVersionKey "ProductVersion" "${APP_VERSION}"
VIAddVersionKey "CompanyName" "${APP_PUBLISHER}"
VIAddVersionKey "FileDescription" "${APP_NAME} 安装程序"
VIAddVersionKey "LegalCopyright" "© 2026 ${APP_PUBLISHER}"
VIProductVersion "1.0.0.0"
VIFileVersion "1.0.0.0"

Section "Install"
  ; 安装前先关闭正在运行的旧版程序，避免 exe 被占用导致"无法写入文件"报错
  nsExec::ExecToLog 'taskkill /F /IM "${APP_EXE}"'
  Pop $0  ; 丢弃返回值（没在运行时 taskkill 也会返回错误码，不影响安装）
  Sleep 500  ; 给系统一点时间释放文件句柄

  SetOutPath "$INSTDIR"
  
  ; 先写卸载器（保证即便后续文件复制失败也能卸载）
  WriteUninstaller "$INSTDIR\Uninstall.exe"
  
  ; 打包整个 dist\AI-Excel-Agent 内容（exe + _internal）
  File /r "dist\${APP_NAME}\*.*"
  
  ; 注册安装信息（卸载入口）
  WriteRegStr HKCU "Software\${APP_NAME}" "InstallDir" "$INSTDIR"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "DisplayName" "${APP_NAME}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "UninstallString" "$\"$INSTDIR\Uninstall.exe$\""
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "DisplayVersion" "${APP_VERSION}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "Publisher" "${APP_PUBLISHER}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "DisplayIcon" "$INSTDIR\${APP_EXE}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "InstallLocation" "$INSTDIR"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "URLInfoAbout" "${APP_URL}"
  WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "NoModify" 1
  WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "NoRepair" 1
  
  ; 开始菜单快捷方式
  CreateDirectory "$SMPROGRAMS\${APP_NAME}"
  CreateShortcut "$SMPROGRAMS\${APP_NAME}\${APP_NAME}.lnk" "$INSTDIR\${APP_EXE}" "" "$INSTDIR\${APP_EXE}"
  CreateShortcut "$SMPROGRAMS\${APP_NAME}\卸载 ${APP_NAME}.lnk" "$INSTDIR\Uninstall.exe"
  
  ; 桌面快捷方式
  CreateShortcut "$DESKTOP\${APP_NAME}.lnk" "$INSTDIR\${APP_EXE}" "" "$INSTDIR\${APP_EXE}"
  
  ; 用户数据目录（首次安装预建，提示用户 outputs 在这里）
  CreateDirectory "$INSTDIR\outputs"
  CreateDirectory "$INSTDIR\data\private"
SectionEnd

Section "Uninstall"
  ; 卸载前先关闭正在运行的程序，否则 exe 被占用无法删除
  nsExec::ExecToLog 'taskkill /F /IM "${APP_EXE}"'
  Pop $0
  Sleep 500

  ; 删除程序文件（保留 outputs/data 用户数据）
  Delete "$INSTDIR\${APP_EXE}"
  Delete "$INSTDIR\Uninstall.exe"
  RMDir /r "$INSTDIR\_internal"
  
  ; 询问是否清除用户数据
  MessageBox MB_YESNO|MB_ICONQUESTION "是否同时删除已生成的表格和配置（outputs、data）？$\n$\n选择「否」将保留你的历史数据，便于重装后继续使用。" IDNO skip_data
    RMDir /r "$INSTDIR\outputs"
    RMDir /r "$INSTDIR\data"
    RMDir /r "$INSTDIR\config"
    RMDir /r "$INSTDIR\diagnostics"
  skip_data:
  
  ; 清理空目录
  RMDir "$INSTDIR"
  
  ; 删除快捷方式
  Delete "$SMPROGRAMS\${APP_NAME}\${APP_NAME}.lnk"
  Delete "$SMPROGRAMS\${APP_NAME}\卸载 ${APP_NAME}.lnk"
  RMDir "$SMPROGRAMS\${APP_NAME}"
  Delete "$DESKTOP\${APP_NAME}.lnk"
  
  ; 清理注册表
  DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}"
  DeleteRegKey HKCU "Software\${APP_NAME}"
SectionEnd
