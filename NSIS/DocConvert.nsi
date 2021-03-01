; Script generated by the HM NIS Edit Script Wizard.

; HM NIS Edit Wizard helper defines
!define PRODUCT_NAME "DocConvert"
!define PRODUCT_VERSION "2.0.7730.26564"
!define PRODUCT_PUBLISHER "Jm's Corp"
!define PRODUCT_WEB_SITE "https://www.namejm.com/"
!define PRODUCT_DIR_REGKEY "Software\Microsoft\Windows\CurrentVersion\App Paths\DocConvert.exe"
!define PRODUCT_UNINST_KEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}"
!define PRODUCT_UNINST_ROOT_KEY "HKLM"

; MUI 1.67 compatible ------
!include "MUI.nsh"

; MUI Settings
!define MUI_ABORTWARNING
!define MUI_ICON "${NSISDIR}\Contrib\Graphics\Icons\modern-install.ico"
!define MUI_UNICON "${NSISDIR}\Contrib\Graphics\Icons\modern-uninstall.ico"

; Language Selection Dialog Settings
!define MUI_LANGDLL_REGISTRY_ROOT "${PRODUCT_UNINST_ROOT_KEY}"
!define MUI_LANGDLL_REGISTRY_KEY "${PRODUCT_UNINST_KEY}"
!define MUI_LANGDLL_REGISTRY_VALUENAME "NSIS:Language"

; Welcome page
!insertmacro MUI_PAGE_WELCOME
; License page
!define MUI_LICENSEPAGE_CHECKBOX
!insertmacro MUI_PAGE_LICENSE ".\LICENSE.txt"
; Directory page
!insertmacro MUI_PAGE_DIRECTORY
; Instfiles page
!insertmacro MUI_PAGE_INSTFILES
; Finish page
!insertmacro MUI_PAGE_FINISH

; Uninstaller pages
!insertmacro MUI_UNPAGE_INSTFILES

; Language files
!insertmacro MUI_LANGUAGE "English"
!insertmacro MUI_LANGUAGE "Korean"

; MUI end ------

Name "${PRODUCT_NAME} ${PRODUCT_VERSION}"
OutFile "DocConvert Setup.exe"
; InstallDir "$PROGRAMFILES\DocConvert"
InstallDir "C:\DocConvert"
InstallDirRegKey HKLM "${PRODUCT_DIR_REGKEY}" ""
ShowInstDetails show
ShowUnInstDetails show

Function .onInit
  !insertmacro MUI_LANGDLL_DISPLAY
FunctionEnd

Section "MainSection" SEC01
  SetOutPath "$INSTDIR\API"
  File "..\bin\Release\API\DocConvert_API.jar"
  SetOutPath "$INSTDIR"
  File "..\bin\Release\AxInterop.HWPCONTROLLib.dll"
  File "..\bin\Release\DocConvert.exe"
  CreateDirectory "$SMPROGRAMS\DocConvert"
  CreateShortCut "$SMPROGRAMS\DocConvert\DocConvert.lnk" "$INSTDIR\DocConvert.exe"
  CreateShortCut "$DESKTOP\DocConvert.lnk" "$INSTDIR\DocConvert.exe"
  File "..\bin\Release\DocConvert_Core.dll"
  File "..\bin\Release\FilePathCheckerModuleExample.dll"
  File "..\bin\Release\ICSharpCode.SharpZipLib.dll"
  File "..\bin\Release\Interop.HWPCONTROLLib.dll"
  File "..\bin\Release\NetOffice.dll"
  File "..\bin\Release\Newtonsoft.Json.dll"
  SetOutPath "$INSTDIR\nl"
  File "..\bin\Release\nl\PdfiumViewer.resources.dll"
  SetOutPath "$INSTDIR"
  File "..\bin\Release\NLog.config"
  File "..\bin\Release\NLog.dll"
  File "..\bin\Release\OfficeApi.dll"
  File "..\bin\Release\pdfium.dll"
  File "..\bin\Release\PdfiumViewer.dll"
  File "..\bin\Release\stdole.dll"
  File "..\bin\Release\System.Collections.Specialized.dll"
  File "..\bin\Release\System.Drawing.Common.dll"
  File "..\bin\Release\System.Net.NameResolution.dll"
  File "..\bin\Release\System.Net.Sockets.dll"
  File "..\bin\Release\System.Runtime.dll"
  File "..\bin\Release\System.Runtime.Extensions.dll"
  File "..\bin\Release\System.Runtime.InteropServices.dll"
  File "..\bin\Release\VBIDEApi.dll"
SectionEnd

Section -AdditionalIcons
  CreateShortCut "$SMPROGRAMS\DocConvert\Uninstall.lnk" "$INSTDIR\uninst.exe"
SectionEnd

Section -Post
  WriteUninstaller "$INSTDIR\uninst.exe"
  WriteRegStr HKLM "${PRODUCT_DIR_REGKEY}" "" "$INSTDIR\DocConvert.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayName" "$(^Name)"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "UninstallString" "$INSTDIR\uninst.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayIcon" "$INSTDIR\DocConvert.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayVersion" "${PRODUCT_VERSION}"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "URLInfoAbout" "${PRODUCT_WEB_SITE}"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "Publisher" "${PRODUCT_PUBLISHER}"
SectionEnd


Function un.onUninstSuccess
  HideWindow
  MessageBox MB_ICONINFORMATION|MB_OK "$(^Name)는(은) 완전히 제거되었습니다."
FunctionEnd

Function un.onInit
!insertmacro MUI_UNGETLANGUAGE
  MessageBox MB_ICONQUESTION|MB_YESNO|MB_DEFBUTTON2 "$(^Name)을(를) 제거하시겠습니까?" IDYES +2
  Abort
FunctionEnd

Section Uninstall
  Delete "$INSTDIR\uninst.exe"
  Delete "$INSTDIR\VBIDEApi.dll"
  Delete "$INSTDIR\System.Runtime.InteropServices.dll"
  Delete "$INSTDIR\System.Runtime.Extensions.dll"
  Delete "$INSTDIR\System.Runtime.dll"
  Delete "$INSTDIR\System.Net.Sockets.dll"
  Delete "$INSTDIR\System.Net.NameResolution.dll"
  Delete "$INSTDIR\System.Drawing.Common.dll"
  Delete "$INSTDIR\System.Collections.Specialized.dll"
  Delete "$INSTDIR\stdole.dll"
  Delete "$INSTDIR\PdfiumViewer.dll"
  Delete "$INSTDIR\pdfium.dll"
  Delete "$INSTDIR\OfficeApi.dll"
  Delete "$INSTDIR\NLog.dll"
  Delete "$INSTDIR\NLog.config"
  Delete "$INSTDIR\nl\PdfiumViewer.resources.dll"
  Delete "$INSTDIR\Newtonsoft.Json.dll"
  Delete "$INSTDIR\NetOffice.dll"
  Delete "$INSTDIR\Interop.HWPCONTROLLib.dll"
  Delete "$INSTDIR\ICSharpCode.SharpZipLib.dll"
  Delete "$INSTDIR\FilePathCheckerModuleExample.dll"
  Delete "$INSTDIR\DocConvert_Core.dll"
  Delete "$INSTDIR\DocConvert.exe"
  Delete "$INSTDIR\AxInterop.HWPCONTROLLib.dll"
  Delete "$INSTDIR\API\DocConvert_API.jar"

  Delete "$SMPROGRAMS\DocConvert\Uninstall.lnk"
  Delete "$DESKTOP\DocConvert.lnk"
  Delete "$SMPROGRAMS\DocConvert\DocConvert.lnk"

  RMDir "$SMPROGRAMS\DocConvert"
  RMDir "$INSTDIR\nl"
  RMDir "$INSTDIR\API"
  RMDir "$INSTDIR"

  DeleteRegKey ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}"
  DeleteRegKey HKLM "${PRODUCT_DIR_REGKEY}"
  SetAutoClose true
SectionEnd