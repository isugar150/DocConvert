; Script generated by the HM NIS Edit Script Wizard.
; FileEncoding: euc-kr

; HM NIS Edit Wizard helper defines
!define PRODUCT_NAME "DocConvert"
!define PRODUCT_VERSION "1.0.0"
!define PRODUCT_PUBLISHER "Jm's Corp"
!define PRODUCT_DIR_REGKEY "Software\Microsoft\Windows\CurrentVersion\App Paths\SuperSocket.SocketService.exe"
!define PRODUCT_UNINST_KEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}"
!define PRODUCT_UNINST_ROOT_KEY "HKLM"

; MUI 1.67 compatible ------
!include "MUI.nsh"

; MUI Settings
!define MUI_ABORTWARNING
!define MUI_ICON "${NSISDIR}\Contrib\Graphics\Icons\modern-install.ico"
!define MUI_UNICON "${NSISDIR}\Contrib\Graphics\Icons\modern-uninstall.ico"

; Welcome page
!insertmacro MUI_PAGE_WELCOME
; License page
!insertmacro MUI_PAGE_LICENSE ".\LICENSE.txt"
; Directory page
!insertmacro MUI_PAGE_DIRECTORY
; Instfiles page
!insertmacro MUI_PAGE_INSTFILES
; Finish page
; !define MUI_FINISHPAGE_RUN "$INSTDIR\DocConvert_Server.exe"
!insertmacro MUI_PAGE_FINISH

; Uninstaller pages
!insertmacro MUI_UNPAGE_INSTFILES

; Language files
!insertmacro MUI_LANGUAGE "Korean"

; MUI end ------

Name "${PRODUCT_NAME}"
OutFile "DocConvert Setup.exe"
InstallDir "$PROGRAMFILES\DocConvert"
InstallDirRegKey HKLM "${PRODUCT_DIR_REGKEY}" ""
ShowInstDetails show
ShowUnInstDetails show

Section "MainSection" SEC01
  ExecWait "taskkill /f /im DocConvert_Server.exe"
  ExecWait "taskkill /f /im DocConvert_Util.exe"
  SetOutPath "$INSTDIR"
  SetOverwrite on
  File "..\bin\Release\JAVA_API\DocConvert_API.jar"
  File "..\bin\Release\JAVA_API\DocConvert_API_JAVA_TEST.jar"
  File "..\bin\Release\JAVA_API\setting.properties"
  File "..\bin\Release\nl\PdfiumViewer.resources.dll"
  File "..\bin\Release\Install\FileZilla_Server-0_9_60_2.exe"
  File "..\bin\Release\Utility\Office Uninstall v1.8.2.exe"
  CreateDirectory "$INSTDIR\JAVA_API"
  CopyFiles "$INSTDIR\DocConvert_API.jar" "$INSTDIR\JAVA_API\DocConvert_API.jar"
  CopyFiles "$INSTDIR\DocConvert_API_JAVA_TEST.jar" "$INSTDIR\JAVA_API\DocConvert_API_JAVA_TEST.jar"
  CopyFiles "$INSTDIR\setting.properties" "$INSTDIR\JAVA_API\setting.properties"
  Delete "$INSTDIR\DocConvert_API.jar"
  Delete "$INSTDIR\DocConvert_API_JAVA_TEST.jar"
  Delete "$INSTDIR\setting.properties"
  CreateDirectory "$INSTDIR\nl"
  CopyFiles "$INSTDIR\PdfiumViewer.resources.dll" "$INSTDIR\nl\PdfiumViewer.resources.dll"
  Delete "$INSTDIR\PdfiumViewer.resources.dll"
  CreateDirectory "$INSTDIR\Utility"
  CopyFiles "$INSTDIR\Office Uninstall v1.8.2.exe" "$INSTDIR\Utility\Office Uninstall v1.8.2.exe"
  Delete "$INSTDIR\Office Uninstall v1.8.2.exe"
  CreateDirectory "$INSTDIR\Install"
  CopyFiles "$INSTDIR\FileZilla_Server-0_9_60_2.exe" "$INSTDIR\Install\FileZilla_Server-0_9_60_2.exe"
  Delete "$INSTDIR\FileZilla_Server-0_9_60_2.exe"
  File "..\bin\Release\vtortola.WebSockets.Rfc6455.dll"
  File "..\bin\Release\vtortola.WebSockets.dll"
  File "..\bin\Release\vtortola.WebSockets.Deflate.dll"
  File "..\bin\Release\VBIDEApi.dll"
  File "..\bin\Release\System.Threading.Tasks.Dataflow.dll"
  File "..\bin\Release\System.Runtime.InteropServices.dll"
  File "..\bin\Release\System.Runtime.Extensions.dll"
  File "..\bin\Release\System.Runtime.dll"
  File "..\bin\Release\System.Net.Sockets.dll"
  File "..\bin\Release\System.Net.NameResolution.dll"
  File "..\bin\Release\System.Drawing.Common.dll"
  File "..\bin\Release\System.Collections.Specialized.dll"
  File "..\bin\Release\SuperSocket.SocketService.exe"
  File "..\bin\Release\SuperSocket.SocketEngine.dll"
  File "..\bin\Release\SuperSocket.SocketBase.dll"
  File "..\bin\Release\SuperSocket.Facility.dll"
  File "..\bin\Release\SuperSocket.Common.dll"
  File "..\bin\Release\stdole.dll"
  File "..\bin\Release\rasterize.js"
  File "..\bin\Release\phantomjs.exe"
  File "..\bin\Release\PdfiumViewer.dll"
  File "..\bin\Release\pdfium.dll"
  File "..\bin\Release\OfficeApi.dll"
  File "..\bin\Release\NLog.dll"
  File "..\bin\Release\NLog.config"
  File "..\bin\Release\Newtonsoft.Json.dll"
  File "..\bin\Release\NetOffice.dll"
  File "..\bin\Release\log4net.dll"
  File "..\bin\Release\Interop.HWPCONTROLLib.dll"
  File "..\bin\Release\ICSharpCode.SharpZipLib.dll"
  File "..\bin\Release\FluentFTP.dll"
  File "..\bin\Release\FilePathCheckerModuleExample.dll"
  File "..\bin\Release\DocConvert_Util.exe"
  CreateShortCut "$DESKTOP\DocConvert Util.lnk" "$INSTDIR\DocConvert_Util.exe"
  CreateDirectory "$SMPROGRAMS\DocConvert"
  CreateShortCut "$SMPROGRAMS\DocConvert\DocConvert Util.lnk" "$INSTDIR\DocConvert_Util.exe"
  File "..\bin\Release\DocConvert_Server.exe"
  CreateShortCut "$SMPROGRAMS\DocConvert\DocConvert Server.lnk" "$INSTDIR\DocConvert_Server.exe"
  CreateShortCut "$DESKTOP\DocConvert Server.lnk" "$INSTDIR\DocConvert_Server.exe"
  File "..\bin\Release\DocConvert_Core.dll"
  File "..\bin\Release\AxInterop.HWPCONTROLLib.dll"
SectionEnd

Section -AdditionalIcons
  CreateShortCut "$SMPROGRAMS\DocConvert\Uninstall.lnk" "$INSTDIR\uninst.exe"
SectionEnd

Section -Post
  WriteUninstaller "$INSTDIR\uninst.exe"
  WriteRegStr HKLM "${PRODUCT_DIR_REGKEY}" "" "$INSTDIR\SuperSocket.SocketService.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayName" "$(^Name)"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "UninstallString" "$INSTDIR\uninst.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayIcon" "$INSTDIR\SuperSocket.SocketService.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayVersion" "${PRODUCT_VERSION}"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "Publisher" "${PRODUCT_PUBLISHER}"
SectionEnd


Function un.onUninstSuccess
  HideWindow
  MessageBox MB_ICONINFORMATION|MB_OK "$(^Name)를 완전히 삭제하였습니다.."
FunctionEnd

Function un.onInit
  MessageBox MB_ICONQUESTION|MB_YESNO|MB_DEFBUTTON2 "$(^Name)를 컴퓨터에서 삭제하시겠습니까?" IDYES +2
  Abort
FunctionEnd

Section Uninstall
  ExecWait "taskkill /f /im DocConvert_Server.exe"
  ExecWait "taskkill /f /im DocConvert_Util.exe"
  Delete "$INSTDIR\${PRODUCT_NAME}.url"
  Delete "$INSTDIR\uninst.exe"
  Delete "$INSTDIR\AxInterop.HWPCONTROLLib.dll"
  Delete "$INSTDIR\DocConvert_Core.dll"
  Delete "$INSTDIR\DocConvert_Server.exe"
  Delete "$INSTDIR\DocConvert_Util.exe"
  Delete "$INSTDIR\FilePathCheckerModuleExample.dll"
  Delete "$INSTDIR\FluentFTP.dll"
  Delete "$INSTDIR\ICSharpCode.SharpZipLib.dll"
  Delete "$INSTDIR\Interop.HWPCONTROLLib.dll"
  Delete "$INSTDIR\log4net.dll"
  Delete "$INSTDIR\NetOffice.dll"
  Delete "$INSTDIR\Newtonsoft.Json.dll"
  Delete "$INSTDIR\NLog.config"
  Delete "$INSTDIR\NLog.dll"
  Delete "$INSTDIR\OfficeApi.dll"
  Delete "$INSTDIR\pdfium.dll"
  Delete "$INSTDIR\PdfiumViewer.dll"
  Delete "$INSTDIR\phantomjs.exe"
  Delete "$INSTDIR\rasterize.js"
  Delete "$INSTDIR\stdole.dll"
  Delete "$INSTDIR\SuperSocket.Common.dll"
  Delete "$INSTDIR\SuperSocket.Facility.dll"
  Delete "$INSTDIR\SuperSocket.SocketBase.dll"
  Delete "$INSTDIR\SuperSocket.SocketEngine.dll"
  Delete "$INSTDIR\SuperSocket.SocketService.exe"
  Delete "$INSTDIR\System.Collections.Specialized.dll"
  Delete "$INSTDIR\System.Drawing.Common.dll"
  Delete "$INSTDIR\System.Net.NameResolution.dll"
  Delete "$INSTDIR\System.Net.Sockets.dll"
  Delete "$INSTDIR\System.Runtime.dll"
  Delete "$INSTDIR\System.Runtime.Extensions.dll"
  Delete "$INSTDIR\System.Runtime.InteropServices.dll"
  Delete "$INSTDIR\System.Threading.Tasks.Dataflow.dll"
  Delete "$INSTDIR\VBIDEApi.dll"
  Delete "$INSTDIR\vtortola.WebSockets.Deflate.dll"
  Delete "$INSTDIR\vtortola.WebSockets.dll"
  Delete "$INSTDIR\vtortola.WebSockets.Rfc6455.dll"
  Delete "$INSTDIR\JAVA_API\DocConvert_API.jar"
  Delete "$INSTDIR\JAVA_API\DocConvert_API_JAVA_TEST.jar"
  Delete "$INSTDIR\JAVA_API\setting.properties"
  RMDir "$INSTDIR\JAVA_API"
  Delete "$INSTDIR\nl\PdfiumViewer.resources.dll"
  RMDir "$INSTDIR\nl"
  Delete "$INSTDIR\Utility\Office Uninstall v1.8.2.exe"
  RMDir "$INSTDIR\Utility"
  Delete "$INSTDIR\Install\FileZilla_Server-0_9_60_2.exe"
  RMDir "$INSTDIR\Install"

  Delete "$SMPROGRAMS\DocConvert\Uninstall.lnk"
  Delete "$DESKTOP\DocConvert Server.lnk"
  Delete "$SMPROGRAMS\DocConvert\DocConvert Server.lnk"
  Delete "$SMPROGRAMS\DocConvert\DocConvert Util.lnk"
  Delete "$DESKTOP\DocConvert Util.lnk"

  RMDir "$SMPROGRAMS\DocConvert"
  RMDir "$INSTDIR"

  DeleteRegKey ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}"
  DeleteRegKey HKLM "${PRODUCT_DIR_REGKEY}"
  SetAutoClose true
SectionEnd