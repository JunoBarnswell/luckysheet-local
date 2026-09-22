Unicode true
RequestExecutionLevel admin
ManifestSupportedOS win10

!ifndef APP_VERSION
  !define APP_VERSION "0.0.0"
!endif
!ifndef APP_STAGE_DIR
  !error "APP_STAGE_DIR is required"
!endif
!ifndef OUT_DIR
  !define OUT_DIR "${APP_STAGE_DIR}\..\installer"
!endif

Name "React Sheets"
Caption "React Sheets ${APP_VERSION}"
OutFile "${OUT_DIR}\ReactSheets-Setup-${APP_VERSION}.exe"
InstallDir "$PROGRAMFILES64\React Sheets"
InstallDirRegKey HKLM "Software\React Sheets" "InstallDir"
ShowInstDetails show
ShowUninstDetails show
VIProductVersion "${APP_VERSION}.0"
VIAddVersionKey "ProductName" "React Sheets"
VIAddVersionKey "ProductVersion" "${APP_VERSION}"
VIAddVersionKey "CompanyName" "React Sheets"
VIAddVersionKey "FileDescription" "React Sheets browser spreadsheet service"

!include "MUI2.nsh"
!include "LogicLib.nsh"

!define SERVICE_EXE "$INSTDIR\service\ReactSheetsService.exe"
!define SERVICE_NAME "ReactSheets"
!define DATA_ROOT "$COMMONAPPDATA\ReactSheets"

!macro ExecRequired command description
  nsExec::ExecToLog '${command}'
  Pop $0
  ${If} $0 != 0
    MessageBox MB_ICONSTOP "${description} failed with exit code $0. See the service log under ${DATA_ROOT}\logs."
    Abort
  ${EndIf}
!macroend

Function .onInit
  ; An upgrade must stop the existing service and make a verified backup before
  ; replacing the JAR/runtime. A first install has no service or data yet.
  ${If} ${FileExists} "${SERVICE_EXE}"
    nsExec::ExecToLog '"${SERVICE_EXE}" stop'
    Pop $0
    ${If} ${FileExists} "$INSTDIR\tools\backup-data.ps1"
      !insertmacro ExecRequired '"$SYSDIR\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "$INSTDIR\tools\backup-data.ps1" -DataRoot "${DATA_ROOT}\data" -OutputDirectory "${DATA_ROOT}\backups"' "Upgrade backup"
    ${EndIf}
  ${EndIf}
FunctionEnd

Function .onInstFailed
  ; Do not leave a half-installed service behind after a failed health check or
  ; service operation. ProgramData is deliberately preserved for recovery.
  ${If} ${FileExists} "${SERVICE_EXE}"
    nsExec::ExecToLog '"${SERVICE_EXE}" stop'
    Pop $0
    nsExec::ExecToLog '"${SERVICE_EXE}" uninstall'
    Pop $0
  ${EndIf}
FunctionEnd

Section "React Sheets service" SEC_SERVICE
  SetOutPath "$INSTDIR\app"
  File "${APP_STAGE_DIR}\app\react-sheets.jar"

  SetOutPath "$INSTDIR\runtime"
  File /r "${APP_STAGE_DIR}\runtime\*"

  SetOutPath "$INSTDIR\service"
  File "${APP_STAGE_DIR}\service\ReactSheetsService.exe"
  File "${APP_STAGE_DIR}\service\ReactSheetsService.xml"

  SetOutPath "$INSTDIR\tools"
  File "${APP_STAGE_DIR}\tools\*.ps1"

  ; Keep application data and operator configuration outside the installation
  ; directory so upgrades and uninstall do not remove workbooks or backups.
  CreateDirectory "${DATA_ROOT}"
  CreateDirectory "${DATA_ROOT}\data"
  CreateDirectory "${DATA_ROOT}\config"
  CreateDirectory "${DATA_ROOT}\backups"
  CreateDirectory "${DATA_ROOT}\logs"
  nsExec::ExecToLog '"$SYSDIR\icacls.exe" "${DATA_ROOT}" /grant "NT AUTHORITY\LOCAL SERVICE:(OI)(CI)M" /T /C'
  Pop $0
  ${If} $0 != 0
    MessageBox MB_ICONSTOP "Unable to grant the service account access to ${DATA_ROOT}. Exit code $0."
    Abort
  ${EndIf}

  ; Preserve operator edits on upgrade; seed the config only on first install.
  ${IfNot} ${FileExists} "${DATA_ROOT}\config\application.properties"
    SetOutPath "${DATA_ROOT}\config"
    File "${APP_STAGE_DIR}\config\application.properties"
  ${EndIf}

  WriteRegStr HKLM "Software\React Sheets" "InstallDir" "$INSTDIR"
  WriteUninstaller "$INSTDIR\uninstall.exe"

  ; WinSW install/start are required. A failed service operation aborts setup.
  !insertmacro ExecRequired '"${SERVICE_EXE}" install' "Service installation"
  !insertmacro ExecRequired '"${SERVICE_EXE}" start' "Service start"
  nsExec::ExecToLog '"$SYSDIR\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "$INSTDIR\tools\health-check.ps1"'
  Pop $0
  ${If} $0 != 0
    nsExec::ExecToLog '"${SERVICE_EXE}" stop'
    Pop $1
    nsExec::ExecToLog '"${SERVICE_EXE}" uninstall'
    Pop $1
    MessageBox MB_ICONSTOP "Service health check failed with exit code $0. See ${DATA_ROOT}\logs."
    Abort
  ${EndIf}

  CreateShortCut "$DESKTOP\React Sheets.lnk" "http://127.0.0.1:8082/" "" "$INSTDIR\service\ReactSheetsService.exe" 0
  CreateDirectory "$SMPROGRAMS\React Sheets"
  CreateShortCut "$SMPROGRAMS\React Sheets\React Sheets.lnk" "http://127.0.0.1:8082/" "" "$INSTDIR\service\ReactSheetsService.exe" 0
  CreateShortCut "$SMPROGRAMS\React Sheets\Uninstall React Sheets.lnk" "$INSTDIR\uninstall.exe"
SectionEnd

Section "Uninstall"
  ${If} ${FileExists} "${SERVICE_EXE}"
    !insertmacro ExecRequired '"${SERVICE_EXE}" stop' "Service stop"
    !insertmacro ExecRequired '"${SERVICE_EXE}" uninstall' "Service removal"
  ${EndIf}
  Delete "$DESKTOP\React Sheets.lnk"
  Delete "$SMPROGRAMS\React Sheets\React Sheets.lnk"
  Delete "$SMPROGRAMS\React Sheets\Uninstall React Sheets.lnk"
  RMDir "$SMPROGRAMS\React Sheets"
  DeleteRegKey HKLM "Software\React Sheets"
  RMDir /r "$INSTDIR"
  ; ${DATA_ROOT} intentionally remains for reinstall or offline recovery.
SectionEnd
