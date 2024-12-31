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
!insertmacro MUI_PAGE_DIRECTORY
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

Var /GLOBAL xlVerReg
Var /GLOBAL xlVerNo
Var /GLOBAL i
Var /GLOBAL keyname
Var /GLOBAL keyname2
Var /GLOBAL keyprefix
Var /GLOBAL openpath
Var /GLOBAL openname
Var /GLOBAL removeold
Var /GLOBAL namelen
Var /GLOBAL lastblankkey

; Installer Section
Section "-Install"
  	SetOutPath $INSTDIR

	ClearErrors
	FileOpen $R0 $INSTDIR\tmp.dat w
	FileClose $R0
	Delete $INSTDIR\tmp.dat
	${If} ${Errors}
  		Abort "User does not have permission to write to the output directory."
	${EndIf}

	ClearErrors
	FileOpen $R0 "$INSTDIR\${filename}" a
	FileClose $R0 
	${If} ${Errors} 
		Abort "Excel is open.  Please close Excel before trying to install."
	${EndIf}

	; Check Installed Excel Version
	ReadRegStr $xlVerReg HKCR "Excel.Application\CurVer" ""

	${If} $xlVerReg == 'Excel.Application.12' ; Excel 2007
		StrCpy $xlVerNo "12.0"
	${ElseIf} $xlVerReg == 'Excel.Application.14' ; Excel 2010
		StrCpy $xlVerNo "14.0"
	${ElseIf} $xlVerReg == 'Excel.Application.15' ; Excel 2013
		StrCpy $xlVerNo "15.0"
	${ElseIf} $xlVerReg == 'Excel.Application.16' ; Excel 2016
		StrCpy $xlVerNo "16.0"
	${Else}
		Abort "An appropriate version of Excel is not installed. $\n${displayname} setup will be canceled."
	${EndIf}


	StrCpy $removeold "false"
	StrLen $namelen "${filename}"
	IntOp $namelen $namelen  * -1
	IntOp $namelen $namelen - 1
	StrCpy $i 0
	loop:
		EnumRegValue $keyname HKCU "Software\Microsoft\Office\$xlVerNo\Excel\Options" $i
		StrCmp $keyname "" done
		StrCpy $keyprefix $keyname 4
		${If} $keyprefix == "OPEN"
			ReadRegStr $openpath HKCU "Software\Microsoft\Office\$xlVerNo\Excel\Options" $keyname
			StrCpy $openname $openpath "" $namelen
			StrCpy $openname $openname -1
			StrCpy $openpath $openpath $namelen
			StrCpy $openpath $openpath "" 1
			${If} "$openname" == "${filename}"
				${If} $removeold == "false"
					MessageBox MB_YESNO "This plugin is already installed.  Replace existing version?" IDYES 0 IDNO abort
					StrCpy $removeold "true"
				${EndIf}
				Delete "$openpath\x86\*"
				RMDir "$openpath\x86"
				Delete "$openpath\x64\*"
				RMDir "$openpath\x64"
				Delete "$openpath\*"
				RMDir "$openpath"
				DeleteRegValue HKCU "Software\Microsoft\Office\$xlVerNo\Excel\Options" $keyname
				DeleteRegValue HKCU "Software\Microsoft\Office\$xlVerNo\Excel\Add-in Manager" "$openpath\${filename}"
						
				EnumRegValue $keyname2 HKCU "Software\Microsoft\Office\$xlVerNo\Excel\Options" $i
				${If} $keyname == $keyname2
				${Else}
					IntOp $i $i - 1
				${EndIf}
			${EndIf}
		${EndIf}
		IntOp $i $i + 1
		Goto loop
	done:
	Goto writekeys
	abort:
	Abort

	writekeys:
	; Find available "OPEN" key
	Var /GLOBAL keyvalue
	StrCpy $i ""
	loop2:
		ReadRegStr $keyvalue HKCU "Software\Microsoft\Office\$xlVerNo\Excel\Options" "OPEN$i"
	${If} $keyvalue == ""
		; Available OPEN key found
		WriteRegStr HKCU "Software\Microsoft\Office\$xlVerNo\Excel\Options" "OPEN$i" '"$INSTDIR\${filename}"'
	${Else}
		IntOp $i $i + 1
		Goto loop2
	${EndIf}

	;Remove any other gaps
	StrCpy $i ""
	StrCpy $lastblankkey ""
	loop3:
		ReadRegStr $keyvalue HKCU "Software\Microsoft\Office\$xlVerNo\Excel\Options" "OPEN$i"
		${If} $keyvalue == ""
			${If} $lastblankkey == ""
				StrCpy $lastblankkey "OPEN$i"
			${EndIf}
		${Else}
			${If} $lastblankkey == "" 
			${Else}
				WriteRegStr HKCU "Software\Microsoft\Office\$xlVerNo\Excel\Options" $lastblankkey $keyvalue
				DeleteRegValue HKCU "Software\Microsoft\Office\$xlVerNo\Excel\Options" "OPEN$i"
				StrCpy $lastblankkey "OPEN$i"
			${EndIf}		
		${EndIf}
		IntOp $i $i + 1
		${If} $i < 1000
			Goto loop3
		${EndIf}


	; Write install data to registry
	WriteRegStr HKCU "Software\${displayname}" "InstallDir" $INSTDIR
	; Install Directory
	WriteRegStr HKCU "Software\${displayname}" "ExcelCurVer" $xlVerNo
	; Current Excel Version

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

; Find AddIn Manager Key and Delete
; AddIn Manager key name and location may have changed since installation depending on actions taken by user in AddIn Manager.
; Need to search for the target AddIn key and delete if found.
ReadRegStr $xlVerNo HKCU "Software\${displayname}" "ExcelCurVer"

StrLen $namelen "${filename}"
IntOp $namelen $namelen  * -1
IntOp $namelen $namelen - 1

StrCpy $i 0
loop:
	EnumRegValue $keyname HKCU "Software\Microsoft\Office\$xlVerNo\Excel\Options" $i
	StrCmp $keyname "" done
	StrCpy $keyprefix $keyname 4
	${If} $keyprefix == "OPEN"
		ReadRegStr $openpath HKCU "Software\Microsoft\Office\$xlVerNo\Excel\Options" $keyname
		StrCpy $openname $openpath "" $namelen
		StrCpy $openname $openname -1
		StrCpy $openpath $openpath $namelen
		StrCpy $openpath $openpath "" 1
		${If} "$openname" == "${filename}"
			Delete "$openpath\x86\*"
			RMDir "$openpath\x86"
			Delete "$openpath\x64\*"
			RMDir "$openpath\x64"
			Delete "$openpath\*"
			RMDir "$openpath"
			DeleteRegValue HKCU "Software\Microsoft\Office\$xlVerNo\Excel\Options" $keyname
			DeleteRegValue HKCU "Software\Microsoft\Office\$xlVerNo\Excel\Add-in Manager" "$openpath\${filename}"	

			EnumRegValue $keyname2 HKCU "Software\Microsoft\Office\$xlVerNo\Excel\Options" $i
			${If} $keyname == $keyname2
			${Else}
				IntOp $i $i - 1
			${EndIf}
		${EndIf}
	${EndIf}
	IntOp $i $i + 1
	Goto loop
done:

	;Remove any other gaps
	StrCpy $i ""
	StrCpy $lastblankkey ""
	loop3:
		ReadRegStr $keyvalue HKCU "Software\Microsoft\Office\$xlVerNo\Excel\Options" "OPEN$i"
		${If} $keyvalue == ""
			${If} $lastblankkey == ""
				StrCpy $lastblankkey "OPEN$i"
			${EndIf}
		${Else}
			${If} $lastblankkey == "" 
			${Else}
				WriteRegStr HKCU "Software\Microsoft\Office\$xlVerNo\Excel\Options" $lastblankkey $keyvalue
				DeleteRegValue HKCU "Software\Microsoft\Office\$xlVerNo\Excel\Options" "OPEN$i"
				StrCpy $lastblankkey "OPEN$i"
			${EndIf}		
		${EndIf}
		IntOp $i $i + 1
		${If} $i < 1000
			Goto loop3
		${EndIf}


DeleteRegKey HKCU "Software\${displayname}"
DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${displayname}"
SectionEnd