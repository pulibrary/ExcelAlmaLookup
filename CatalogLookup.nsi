; NSIS Excel Add-In Installer Script

RequestExecutionLevel user 
!include MUI.nsh
!include LogicLib.nsh

; General
!define filename "CatalogLookup.xlam"
!define yazdllx86 "yaz5x86.dll"
!define yazdllx64 "yaz5x64.dll"
!define yazdll "yaz5.dll"
!define libxml2x86 "libxml2x86.dll"
!define libxsltx86 "libxsltx86.dll"
!define libxml2x64 "libxml2x64.dll"
!define libxsltx64 "libxsltx64.dll"
!define libxml2 "libxml2.dll"
!define libxslt "libxslt.dll"
!define displayname "Excel Local Catalog Lookup"

Name "${displayname}"
OutFile "CatalogLookupInstaller.exe"
InstallDir "$APPDATA\${displayname}"
InstallDirRegKey HKCU "Software\${displayname}" "InstallDir" ;

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_WELCOME
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_UNPAGE_FINISH

!insertmacro MUI_LANGUAGE "English"

; Interface Settings
!define MUI_ABORTWARNING

;Prerequisites section

Section "-Prerequisites"
SectionEnd

SetOverwrite On

; Installer Section
Section "-Install"
  	SetOutPath $INSTDIR

	CreateDirectory $INSTDIR\x86
	CreateDirectory $INSTDIR\x64
	
	; ADD FILES HERE
	File "${filename}"
	File "${yazdllx86}"
	Rename "${yazdllx86}" "x86\${yazdll}"
	File "${yazdllx64}"
	Rename "${yazdllx64}" "x64\${yazdll}"
	File "${libxml2x86}"
	Rename "${libxml2x86}" "x86\${libxml2}"
	File "${libxsltx86}"
	Rename "${libxsltx86}" "x86\${libxslt}"
	File "${libxml2x64}"
	Rename "${libxml2x64}" "x64\${libxml2}"
	File "${libxsltx64}"
	Rename "${libxsltx64}" "x64\${libxslt}"

	; Check Installed Excel Version
	ReadRegStr $1 HKCR "Excel.Application\CurVer" ""

	${If} $1 == 'Excel.Application.12' ; Excel 2007
		StrCpy $2 "12.0"
	${ElseIf} $1 == 'Excel.Application.14' ; Excel 2010
		StrCpy $2 "14.0"
	${ElseIf} $1 == 'Excel.Application.15' ; Excel 2013
		StrCpy $2 "15.0"
	${ElseIf} $1 == 'Excel.Application.16' ; Excel 2016
		StrCpy $2 "16.0"
	${Else}
		Abort "An appropriate version of Excel is not installed. $\n${displayname} setup will be canceled."
	${EndIf}

	; Find available "OPEN" key
	StrCpy $3 ""
	loop:
		ReadRegStr $4 HKCU "Software\Microsoft\Office\$2\Excel\Options" "OPEN$3"
	${If} $4 == ""
		; Available OPEN key found
	${Else}
		IntOp $3 $3 + 1
		Goto loop
	${EndIf}

	; Write install data to registry
	WriteRegStr HKCU "Software\${displayname}" "InstallDir" $INSTDIR
	; Install Directory
	WriteRegStr HKCU "Software\${displayname}" "ExcelCurVer" $2
	; Current Excel Version

	; Write key to install AddIn in Excel Addin Manager
	WriteRegStr HKCU "Software\Microsoft\Office\$2\Excel\Options" "OPEN$3" '"$INSTDIR\${filename}"'

	; Write keys to uninstall
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${displayname}" "DisplayName" "${displayname}"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${displayname}" "UninstallString" '"$INSTDIR\uninstall.exe"'
	WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${displayname}" "NoModify" 1
	WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${displayname}" "NoRepair" 1

	; Create uninstaller
	WriteUninstaller "$INSTDIR\Uninstall.exe"
SectionEnd


; Uninstaller Section
Section "Uninstall"
StrCpy $INSTDIR "$APPDATA\${displayname}"
; ADD FILES HERE...
Delete "$INSTDIR\${filename}"
Delete "$INSTDIR\x86\*"
RMDir "$INSTDIR\x86"
Delete "$INSTDIR\x64\*"
RMDir "$INSTDIR\x64"
Delete "$INSTDIR\uninstall.exe"

RMDir "$INSTDIR"

; Find AddIn Manager Key and Delete
; AddIn Manager key name and location may have changed since installation depending on actions taken by user in AddIn Manager.
; Need to search for the target AddIn key and delete if found.
ReadRegStr $2 HKCU "Software\${displayname}" "ExcelCurVer"
StrCpy $3 ""

loop:
ReadRegStr $4 HKCU "Software\Microsoft\Office\$2\Excel\Options" "OPEN$3"
${If} $4 == '"$INSTDIR\${filename}"'
	; Found Key
	DeleteRegValue HKCU "Software\Microsoft\Office\$2\Excel\Options" "OPEN$3"
${ElseIf} $4 == ""
	; Blank Key Found. Addin is no longer installed in AddIn Manager.
	; Need to delete Addin Manager Reference.
	DeleteRegValue HKCU "Software\Microsoft\Office\$2\Excel\Add-in Manager" "$INSTDIR\${filename}"
${Else}
	IntOp $3 $3 + 1
	Goto loop
${EndIf}

DeleteRegKey HKCU "Software\${displayname}"
DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${displayname}"
SectionEnd