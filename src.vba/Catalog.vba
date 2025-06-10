Attribute VB_Name = "Catalog"
Global oRegEx As Object
Global oXMLHTTP As Object
Global oXMLDOM As Object
Global oConverter As Object
Global oZConn As LongPtr

Global aExplainFields As Variant
Global bTerminateLoop As Boolean
Global bKeepTryingURL As Boolean
Global bIsoholdEnabled  As Boolean
Global bIsAlma As Boolean
Global sCatalogURL As String
Global sAuth As String
Global aAlmaSearchKeys As Variant

Global sFileName As String
Global sSheetName As String

Public Const HKEY_CURRENT_USER = &H80000001
Public Const sVersion = "v1.4.0"
Public Const sRepoURL = "https://github.com/pulibrary/ExcelAlmaLookup"
Public Const sBlacklightURL = "https://catalog.princeton.edu/catalog.json?q="
Public Const sLCCatURL = "http://lx2.loc.gov:210/LCDB"
Public Const sIPLCReshareURL = "https://borrowdirect.reshare.indexdata.com/api/v1/search?type=AllFields&field%5B%5D=fullRecord&lookfor="

Public Const iMaximumRecords = 25

Public Const sWCZhost = "zcat.oclc.org"
Public Const sWCZport = 210
Public Const sWCZDB = "OLUCWorldCat"

Public Const sRegistryDir = "Excel Catalog Lookup"
Public Const sYAZdll = "yaz5"

#If Win64 Then
Public Const sDllVersion = "x64"
Private Declare PtrSafe Sub CopyMemory Lib "ntdll" Alias "RtlCopyMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare PtrSafe Function SetDefaultDllDirectories Lib "kernel32" (ByVal dwFlags As Long) As LongPtr
Private Declare PtrSafe Function AddDllDirectory Lib "kernel32" (ByVal lpLibDirectory As String) As LongPtr
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr

Private Declare PtrSafe Function ZOOM_connection_create Lib "yaz5.dll" (ByVal Options As Integer) As LongPtr
Private Declare PtrSafe Sub ZOOM_connection_connect Lib "yaz5.dll" (ByVal c As LongPtr, ByVal Host As String, ByVal portnum As Integer)
Private Declare PtrSafe Function ZOOM_connection_option_get Lib "yaz5.dll" (ByVal c As LongPtr, ByVal key As String) As LongPtr
Private Declare PtrSafe Sub ZOOM_connection_option_set Lib "yaz5.dll" (ByVal c As LongPtr, ByVal key As String, ByVal val As String)
Private Declare PtrSafe Sub ZOOM_connection_destroy Lib "yaz5.dll" (ByVal c As LongPtr)
Private Declare PtrSafe Function ZOOM_connection_errcode Lib "yaz5.dll" (ByVal c As LongPtr) As LongPtr
Private Declare PtrSafe Function ZOOM_connection_search_pqf Lib "yaz5.dll" (ByVal c As LongPtr, ByVal q As String) As LongPtr

Private Declare PtrSafe Function ZOOM_resultset_size Lib "yaz5.dll" (ByVal r As LongPtr) As Integer
Private Declare PtrSafe Function ZOOM_resultset_record Lib "yaz5.dll" (ByVal r As LongPtr, ByVal pos As Integer) As LongPtr
Private Declare PtrSafe Sub ZOOM_resultset_option_set Lib "yaz5.dll" (ByVal r As LongPtr, ByVal key As String, ByVal val As String)
Private Declare PtrSafe Sub ZOOM_resultset_destroy Lib "yaz5.dll" (ByVal r As LongPtr)

Private Declare PtrSafe Function ZOOM_record_get Lib "yaz5.dll" (ByVal r As LongPtr, ByVal typ As String, ByRef size As Long) As LongPtr
Private Declare PtrSafe Sub ZOOM_record_destroy Lib "yaz5.dll" (ByVal r As LongPtr)

#Else
Public Const sDllVersion = "x86"
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare PtrSafe Function SetDefaultDllDirectories Lib "kernel32" (ByVal dwFlags As Long) As LongPtr
Private Declare PtrSafe Function AddDllDirectory Lib "kernel32" (ByVal lpLibDirectory As String) As LongPtr
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr

Private Declare PtrSafe Function ZOOM_connection_create Lib "yaz5.dll" Alias "_ZOOM_connection_create@4" (ByVal Options As Integer) As LongPtr
Private Declare PtrSafe Sub ZOOM_connection_connect Lib "yaz5.dll" Alias "_ZOOM_connection_connect@12" (ByVal c As LongPtr, ByVal Host As String, ByVal portnum As Integer)
Private Declare PtrSafe Function ZOOM_connection_option_get Lib "yaz5.dll" Alias "_ZOOM_connection_option_get@8" (ByVal c As LongPtr, ByVal key As String) As LongPtr
Private Declare PtrSafe Sub ZOOM_connection_option_set Lib "yaz5.dll" Alias "_ZOOM_connection_option_set@12" (ByVal c As LongPtr, ByVal key As String, ByVal val As String)
Private Declare PtrSafe Sub ZOOM_connection_destroy Lib "yaz5.dll" Alias "_ZOOM_connection_destroy@4" (ByVal c As LongPtr)
Private Declare PtrSafe Function ZOOM_connection_errcode Lib "yaz5.dll" Alias "_ZOOM_connection_errcode@4" (ByVal c As LongPtr) As LongPtr
Private Declare PtrSafe Function ZOOM_connection_search_pqf Lib "yaz5.dll" Alias "_ZOOM_connection_search_pqf@8" (ByVal c As LongPtr, ByVal q As String) As LongPtr

Private Declare PtrSafe Function ZOOM_resultset_size Lib "yaz5.dll" Alias "_ZOOM_resultset_size@4" (ByVal r As LongPtr) As Integer
Private Declare PtrSafe Function ZOOM_resultset_record Lib "yaz5.dll" Alias "_ZOOM_resultset_record@8" (ByVal r As LongPtr, ByVal pos As Integer) As LongPtr
Private Declare PtrSafe Sub ZOOM_resultset_option_set Lib "yaz5.dll" Alias "_ZOOM_resultset_option_set@12" (ByVal r As LongPtr, ByVal key As String, ByVal val As String)
Private Declare PtrSafe Sub ZOOM_resultset_destroy Lib "yaz5.dll" Alias "_ZOOM_resultset_destroy@4" (ByVal r As LongPtr)

Private Declare PtrSafe Function ZOOM_record_get Lib "yaz5.dll" Alias "_ZOOM_record_get@12" (ByVal r As LongPtr, ByVal typ As String, ByRef size As Long) As LongPtr
Private Declare PtrSafe Sub ZOOM_record_destroy Lib "yaz5.dll" Alias "_ZOOM_record_destroy@4" (ByVal r As LongPtr)
#End If

'Initialize global objects
Private Sub Initialize()
    sVer = GetSetting("Excel Catalog Lookup", "General", "Version", "NONE")
    If sVer = "NONE" Then
        MigrateSettings
    End If
    sPluginDir = Application.AddIns("CatalogLookup").Path
    
    On Error GoTo ErrHandler
    Set oRegEx = CreateObject("vbscript.regexp")
    With oRegEx
        .MultiLine = False
        .Global = True
        .IgnoreCase = True
    End With
    Set oXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    Set oXMLDOM = CreateObject("MSXML2.DomDocument")
    oXMLDOM.SetProperty "SelectionLanguage", "XPath"
    Set oConverter = CreateObject("ADODB.Stream")
    
    SetDefaultDllDirectories (&H1000&)
    sDllPath = StrConv(sPluginDir & "\" & sDllVersion & "\", vbUnicode)
    AddDllDirectory (sDllPath)
    LoadLibrary (sYAZdll)
    
    Exit Sub
ErrHandler:
    MsgBox ("There was an error initializing the plugin.  Please try again.")
End Sub

Sub MigrateSettings()
    SaveSetting sRegistryDir, "General", "Version", sVersion
    Set oReg = CreateObject("WScript.Shell")
    On Error Resume Next
    sAuths = oReg.RegRead("HKEY_CURRENT_USER\Software\Excel Local Catalog Lookup\CatalogAuth")
    sFieldSets = oReg.RegRead("HKEY_CURRENT_USER\Software\Excel Local Catalog Lookup\FieldSets")
    sURLs = oReg.RegRead("HKEY_CURRENT_USER\Software\Excel Local Catalog Lookup\CatalogURL")
    If (InStr(1, sAuths, "|") > 0 And InStr(1, sAuths, ChrW(166)) = 0) Or _
        (InStr(1, sFieldSets, "|") > 0 And InStr(1, sFieldSets, ChrW(166)) = 0) Then
        Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
        oReg.GetStringValue HKEY_CURRENT_USER, "Software\Excel Local Catalog Lookup", "CatalogAuth", sAuths
        oReg.GetStringValue HKEY_CURRENT_USER, "Software\Excel Local Catalog Lookup", "FieldSets", sFieldSets
        oReg.GetStringValue HKEY_CURRENT_USER, "Software\Excel Local Catalog Lookup", "CatalogURL", sURLs
    End If
    If sURLs <> "" Then
        aURLs = Split(sURLs, "|")
        aAuths = Split(sAuths, "|")
        SaveSetting sRegistryDir, "Sources", "MAX", UBound(aURLs)
        SaveSetting sRegistryDir, "Sources", "SELECTED", aURLs(0)
        For i = 0 To UBound(aURLs)
            SaveSetting sRegistryDir, "Sources", "SOURCE" & Format(i, "000"), aURLs(i)
            For j = 0 To UBound(aAuths)
                If InStr(1, aAuths(j), aURLs(i) & ChrW(166)) = 1 Then
                    sAuthValue = Mid(aAuths(j), Len(aURLs(i)) + 2)
                    SaveSetting sRegistryDir, "Sources", "AUTH" & Format(i, "000"), sAuthValue
                End If
            Next j
        Next i
    End If
    If sFieldSets <> "" Then
        aFieldSets = Split(sFieldSets, "|")
        SaveSetting sRegistryDir, "FieldSets", "MAXALL", UBound(aFieldSets)
        For i = 0 To UBound(aFieldSets)
            aFieldList = Split(aFieldSets(i), ChrW(166))
            SaveSetting sRegistryDir, "FieldSets", "NAME" & Format(i, "000"), aFieldList(0)
            SaveSetting sRegistryDir, "FieldSets", "MAX" & Format(i, "000"), UBound(aFieldList) - 1
            For j = 1 To UBound(aFieldList)
                SaveSetting sRegistryDir, "FieldSets", "FIELD" & Format(i, "000") & "-" & Format(j - 1, "000"), aFieldList(j)
            Next j
        Next i
    End If
End Sub

Function GetLatestVersionNumber()
    If oXMLHTTP Is Nothing Then
        Initialize
    End If
    
    sAPIUrl = Replace(sRepoURL, "github.com", "api.github.com/repos") & "/releases/latest"
    sTagLabel = "tag_name"
    With oXMLHTTP
        .Open "GET", sAPIUrl, True
        .Send
        Do While .readyState <> 4
            DoEvents
        Loop
        GetLatestVersionNumber = sVersion
        If .Status = 200 Then
            iTagStart = InStr(1, .responseText, sTagLabel)
            iTagStart = iTagStart + Len(sTagLabel & """:""")
            If iTagStart > 0 Then
                iTagEnd = InStr(iTagStart, .responseText, """")
                If iTagEnd > 0 Then
                    GetLatestVersionNumber = Mid(.responseText, iTagStart, iTagEnd - iTagStart)
                End If
            End If
        End If
    End With
End Function

'Main function called when toolbar button is pressed.  Sets up dialog box.
Sub LookupInterface(control As IRibbonControl)
    If Right(ActiveWorkbook.FullName, 4) = ".xls" Then
        iResult = MsgBox("File must be in XLSX format.  Convert Now?", vbYesNo, "Question")
        If iResult = vbYes Then
            sXLSname = ActiveWorkbook.FullName
            sXLSXname = Replace(sXLSname, ".xls", ".xlsx")
            ActiveWorkbook.SaveAs Filename:=sXLSXname, FileFormat:=xlOpenXMLWorkbook
            Kill sXLSname
            Workbooks.Open sXLSXname
        Else
            Exit Sub
        End If
    End If
    
    If GetSetting(sRegistryDir, "General", "Version", "") = "" Then
        Initialize
    End If
    
    PopulateCombos
    RedrawButtons
    
    sFileName = ActiveWorkbook.Name
    sSheetName = ActiveSheet.Name
        
    LookupDialog.ResultColumnSpinner.Value = FindLastColumn() + 1
    sSourceRange = Selection.Address
    LookupDialog.LookupRange.Text = sSourceRange
    sSourceColumn = Split(Cells(1, Range(Selection.Address).Column).Address(True, False), "$")(0)
    LookupDialog.SearchValueBox.Value = "[[" & sSourceColumn & "]]"

    sLatestVersion = GetLatestVersionNumber
    If sLatestVersion = sVersion Then
        LookupDialog.VersionLabel.Caption = "You are using the latest version. (" & sVersion & ")"
    ElseIf StrComp(sLatestVersion, sVersion) < 0 Then
        LookupDialog.VersionLabel.Caption = "You are using a pre-release version. (" & sVersion & ")"
    Else
        LookupDialog.VersionLabel.Caption = "A newer version is available! (" & sLatestVersion & ")"
    End If
    LookupDialog.Show
End Sub

Sub AddURLtoRegistry(sURL)
    bFoundEmptySlot = False
    bDuplicate = False
    iMax = GetSetting(Catalog.sRegistryDir, "Sources", "MAX", -1)
    For i = 0 To iMax
        sRegURL = GetSetting(Catalog.sRegistryDir, "Sources", "SOURCE" & Format(i, "000"), "")
        If sRegURL = "" Then
            bFoundEmptySlot = True
            SaveSetting Catalog.sRegistryDir, "Sources", "SOURCE" & Format(i, "000"), sURL
            Exit For
        End If
        If sRegURL = sURL Then
            bDuplicate = True
            Exit Sub
        End If
    Next i
    If Not bFoundEmptySlot Then
        iMax = iMax + 1
        SaveSetting Catalog.sRegistryDir, "Sources", "MAX", iMax
        SaveSetting Catalog.sRegistryDir, "Sources", "SOURCE" & Format(iMax, "000"), sURL
    End If
End Sub

Sub RemoveURLfromRegistry(sURL)
    iMax = GetSetting(Catalog.sRegistryDir, "Sources", "MAX", -1)
    For i = 0 To iMax
        sRegURL = GetSetting(Catalog.sRegistryDir, "Sources", "SOURCE" & Format(i, "000"), "")
        If sURL = sRegURL Then
            DeleteSetting Catalog.sRegistryDir, "Sources", "SOURCE" & Format(i, "000")
            If GetSetting(Catalog.sRegistryDir, "Sources", "AUTH" & Format(i, "000"), "") <> "" Then
                DeleteSetting Catalog.sRegistryDir, "Sources", "AUTH" & Format(i, "000")
            End If
        End If
    Next i
End Sub

Sub SaveCatalogAuthToRegistry()
    If GetSetting(sRegistryDir, "General", "Version", "NONE") = "NONE" Then
        Initialize
    End If
    
    sCatalogURL = LookupDialog.CatalogURLBox.Text
    AddURLtoRegistry (sCatalogURL)
    iMax = GetSetting(sRegistryDir, "Sources", "MAX", -1)
    For i = 0 To iMax
        sRegURL = GetSetting(sRegistryDir, "Sources", "SOURCE" & Format(i, "000"), "")
        If sCatalogURL = sRegURL Then
            SaveSetting sRegistryDir, "Sources", "AUTH" & Format(i, "000"), sAuth
        End If
    Next i
End Sub

Sub ClearRegistryAuth(sURL)
    iMax = GetSetting(sRegistryDir, "Sources", "MAX", -1)
    For i = 0 To iMax
        sRegURL = GetSetting(sRegistryDir, "Sources", "SOURCE" & Format(i, "000"), "")
        If sURL = sRegURL Then
            sRegAuth = GetSetting(sRegistryDir, "Sources", "AUTH" & Format(i, "000"), "")
            If sRegAuth <> "" Then
                DeleteSetting Catalog.sRegistryDir, "Sources", "AUTH" & Format(i, "000")
            End If
        End If
    Next i
End Sub
 

Function SaveFieldSet(sSetName)
    SaveFieldSet = True
    If LookupDialog.ResultTypeList.ListCount = 0 Then
        MsgBox ("Please add at least one result type to the set")
        SaveFieldSet = False
    Else
        bFound = False
        iFoundIndex = -1
        iGapIndex = -1
        iMax = GetSetting(sRegistryDir, "FieldSets", "MAXALL", -1)
        For i = 0 To iMax
            sRegName = GetSetting(sRegistryDir, "FieldSets", "NAME" & Format(i, "000"), "")
            If Not bFound And sRegName = "" Then
                iGapIndex = i
            End If
            If Not bFound And sRegName = sSetName Then
                bFound = True
                iFoundIndex = i
            End If
        Next i
        If bFound Then
            iSetMax = GetSetting(sRegistryDir, "FieldSets", "MAX" & Format(iFoundIndex, "000"), -1)
            For i = 0 To iSetMax
                DeleteSetting sRegistryDir, "FieldSets", "FIELD" & Format(iFoundIndex, "000") & "-" & Format(i, "000")
            Next i
        Else
            If iGapIndex > -1 Then
                iFoundIndex = iGapIndex
            Else
                iFoundIndex = iMax + 1
                SaveSetting sRegistryDir, "FieldSets", "MAXALL", iFoundIndex
            End If
            SaveSetting sRegistryDir, "FieldSets", "NAME" & Format(iFoundIndex, "000"), sSetName
        End If
        iNewSetMax = LookupDialog.ResultTypeList.ListCount - 1
        SaveSetting sRegistryDir, "FieldSets", "MAX" & Format(iFoundIndex, "000"), iNewSetMax
        For i = 0 To iNewSetMax
            sNewField = LookupDialog.ResultTypeList.List(i)
            SaveSetting sRegistryDir, "FieldSets", "FIELD" & Format(iFoundIndex, "000") & "-" & Format(i, "000"), sNewField
        Next i
    End If
End Function

Sub DeleteFieldSet(sSetName)
    iMax = CInt(GetSetting(sRegistryDir, "FieldSets", "MAXALL", -1))
    For i = 0 To iMax
        sRegName = GetSetting(sRegistryDir, "FieldSets", "NAME" & Format(i, "000"), "")
        If sRegName = sSetName Then
            iSetMax = GetSetting(sRegistryDir, "FieldSets", "MAX" & Format(i, "000"), -1)
            For j = 0 To iSetMax
                If GetSetting(sRegistryDir, "FieldSets", "MAX" & Format(i, "000"), "NONE") <> "NONE" Then
                    DeleteSetting sRegistryDir, "FieldSets", "FIELD" & Format(i, "000") & "-" & Format(j, "000")
                End If
            Next j
            If GetSetting(sRegistryDir, "FieldSets", "NAME" & Format(i, "000"), "NONE") <> "NONE" Then
                DeleteSetting sRegistryDir, "FieldSets", "NAME" & Format(i, "000")
            End If
            If GetSetting(sRegistryDir, "FieldSets", "MAX" & Format(i, "000"), "NONE") <> "NONE" Then
                DeleteSetting sRegistryDir, "FieldSets", "MAX" & Format(i, "000")
            End If
            If i = iMax Then
                SaveSetting sRegistryDir, "FieldSets", "MAXALL", i - 1
            End If
        End If
    Next i
End Sub

Function GetSourceRegIndex(sSource) As Integer
    GetSourceRegIndex = -1
    iURLsMax = GetSetting(sRegistryDir, "Sources", "MAX", 0)
    For i = 0 To iURLsMax
        sURL = GetSetting(sRegistryDir, "Sources", "SOURCE" & Format(i, "000"))
        If sURL = sSource Then
            GetSourceRegIndex = i
        End If
    Next i
End Function

Sub PopulateCombos()
    Dim sCatalogURL As String
    
    LookupDialog.CatalogURLBox.Clear
    iURLsMax = GetSetting(sRegistryDir, "Sources", "MAX", -1)
    
    For i = 0 To iURLsMax
        sURL = GetSetting(sRegistryDir, "Sources", "SOURCE" & Format(i, "000"))
        If sURL <> "" Then
            LookupDialog.CatalogURLBox.AddItem sURL
        End If
    Next i
    
    sSelected = GetSetting(sRegistryDir, "Sources", "SELECTED", "")
    iSelected = GetSourceRegIndex(sSelected)
    LookupDialog.CatalogURLBox.Value = sSelected
    sAuth = GetSetting(sRegistryDir, "Sources", "AUTH" & Format(iSelected, "000"), "")

    LookupDialog.FieldSetList.Clear
    iFieldSetsMax = GetSetting(sRegistryDir, "FieldSets", "MAXALL", 0)
    For i = 0 To iFieldSetsMax
        sName = GetSetting(sRegistryDir, "FieldSets", "NAME" & Format(i, "000"), "")
        If sName <> "" Then
            LookupDialog.FieldSetList.AddItem sName
        End If
    Next i

    LookupDialog.BooleanCombo.AddItem "AND"
    LookupDialog.BooleanCombo.AddItem "OR"
    
    Dim aOtherSources(4, 2) As Variant
    aOtherSources(0, 0) = "source:recap"
    aOtherSources(0, 1) = "ReCAP"
    aOtherSources(1, 0) = "source:borrowdirect"
    aOtherSources(1, 1) = "BorrowDirect (IPLC ReShare)"
    aOtherSources(2, 0) = "source:lccat"
    aOtherSources(2, 1) = "Library of Congress"
    aOtherSources(3, 0) = "source:worldcat"
    aOtherSources(3, 1) = "WorldCat"

    OtherSourcesDialog.OtherSourcesListBox.List = aOtherSources

    PopulateSourceDependentOptions

End Sub

Sub PopulateSourceDependentOptions()
    LookupDialog.ResultTypeCombo.Clear
    LookupDialog.ResultTypeCombo.AddItem "True/False"
    If Catalog.bIsAlma Then
        LookupDialog.ResultTypeCombo.AddItem "MMS ID"
    ElseIf LookupDialog.CatalogURLBox = "source:worldcat" Then
        LookupDialog.ResultTypeCombo.AddItem "OCLC No."
    Else
        LookupDialog.ResultTypeCombo.AddItem "Catalog ID"
    End If
    LookupDialog.ResultTypeCombo.AddItem "ISBN"
    If Not LookupDialog.CatalogURLBox = "source:worldcat" Then
        LookupDialog.ResultTypeCombo.AddItem "OCLC No."
    End If
    LookupDialog.ResultTypeCombo.AddItem "Title"
    If Not LookupDialog.CatalogURLBox = "source:recap" Then
        LookupDialog.ResultTypeCombo.AddItem "Language code"
        LookupDialog.ResultTypeCombo.AddItem "Leader"
    End If
    If Catalog.bIsAlma Then
        LookupDialog.ResultTypeCombo.AddItem "*Call No."
        LookupDialog.ResultTypeCombo.AddItem "*Location/DB Name"
        LookupDialog.ResultTypeCombo.AddItem "*Coverage"
        LookupDialog.ResultTypeCombo.AddItem "**Barcode"
        LookupDialog.ResultTypeCombo.AddItem "**Item Location"
        LookupDialog.ResultTypeCombo.AddItem "**Item Enum/Chron"
        LookupDialog.ResultTypeCombo.AddItem "**Shelf Locator"
    End If
    
    If LookupDialog.CatalogURLBox = "source:recap" Then
        LookupDialog.ResultTypeCombo.Style = 2 'fmStyleDropDownList
        LookupDialog.ResultTypeCombo.AddItem "LCCN"
        LookupDialog.ResultTypeCombo.AddItem "ReCAP Holdings"
        LookupDialog.ResultTypeCombo.AddItem "ReCAP CGD"
    Else
        LookupDialog.ResultTypeCombo.Style = 0 'fmStyleDropDownCombo
    End If
    
    If LookupDialog.CatalogURLBox = "source:borrowdirect" Then
        LookupDialog.ResultTypeCombo.AddItem "BorrowDirect Holdings"
    End If
    
    If LookupDialog.CatalogURLBox = "source:worldcat" Then
        LookupDialog.ResultTypeCombo.AddItem "WorldCat Holdings"
        LookupDialog.ResultTypeCombo.AddItem "Holdings Count"
    End If
    
    LookupDialog.SearchFieldCombo.Clear
        
    If Not bIsAlma Then
        LookupDialog.SearchFieldCombo.Style = 2 'fmStyleDropDownList
        LookupDialog.SearchFieldCombo.AddItem "Keywords"
        If LookupDialog.CatalogURLBox = "source:worldcat" Or LookupDialog.CatalogURLBox = "source:recap" Then
            LookupDialog.SearchFieldCombo.Enabled = True
            LookupDialog.SearchFieldCombo.AddItem "Title"
            LookupDialog.SearchFieldCombo.AddItem "ISBN"
            If LookupDialog.CatalogURLBox = "source:recap" Then
                LookupDialog.SearchFieldCombo.AddItem "LCCN"
            Else
                LookupDialog.SearchFieldCombo.AddItem "ISSN"
                LookupDialog.SearchFieldCombo.AddItem "Z-Title-Date"
                LookupDialog.SearchFieldCombo.AddItem "Z-Author-Title-Date"
            End If
            LookupDialog.SearchFieldCombo.AddItem "OCLC No."
                        
        Else
            LookupDialog.SearchFieldCombo.Enabled = False
        End If
        LookupDialog.BooleanCombo.Enabled = False
        LookupDialog.BooleanCombo.Value = ""
        LookupDialog.OperatorCombo.Enabled = False
        LookupDialog.OperatorCombo.Value = "="
        LookupDialog.SearchValueBox.Enabled = False
        LookupDialog.SearchListBox.Clear
        LookupDialog.SearchListBox.Enabled = False
        sSourceColumn = Split(Cells(1, Range(Selection.Address).Column).Address(True, False), "$")(0)
        LookupDialog.SearchValueBox.Value = "[[" & sSourceColumn & "]]"
    Else
        LookupDialog.SearchFieldCombo.Style = 0 'fmStyleDropDownCombo
        LookupDialog.SearchFieldCombo.Clear
        LookupDialog.SearchFieldCombo.Enabled = True
        ReDim aAlmaSearchKeys(7, 2) As Variant
        aAlmaSearchKeys(0, 0) = "Keywords"
        aAlmaSearchKeys(0, 1) = "alma.all_for_ui"
        aAlmaSearchKeys(1, 0) = "Call No."
        aAlmaSearchKeys(1, 1) = "alma.PermanentCallNumber"
        aAlmaSearchKeys(2, 0) = "Title"
        aAlmaSearchKeys(2, 1) = "alma.title"
        aAlmaSearchKeys(3, 0) = "ISBN"
        aAlmaSearchKeys(3, 1) = "alma.isbn"
        aAlmaSearchKeys(4, 0) = "ISSN"
        aAlmaSearchKeys(4, 1) = "alma.issn"
        aAlmaSearchKeys(5, 0) = "MMS ID"
        aAlmaSearchKeys(5, 1) = "rec.id"
        aAlmaSearchKeys(6, 0) = "Barcode"
        aAlmaSearchKeys(6, 1) = "alma.barcode"
        For i = 0 To UBound(aAlmaSearchKeys) - 1
            LookupDialog.SearchFieldCombo.AddItem aAlmaSearchKeys(i, 0)
        Next i
        LookupDialog.SearchFieldCombo.AddItem "Other fields..."
        LookupDialog.SearchFieldCombo.ListIndex = 0
        
        LookupDialog.BooleanCombo.Enabled = True
        LookupDialog.OperatorCombo.Enabled = True
        LookupDialog.SearchValueBox.Enabled = True
        LookupDialog.SearchListBox.Enabled = True
        
        PopulateOperatorCombo
    End If
    If LookupDialog.SearchFieldCombo.ListCount > 0 Then
        LookupDialog.SearchFieldCombo.ListIndex = 0
    End If
    LookupDialog.ResultTypeCombo.ListIndex = 0

End Sub

Sub PopulateOperatorCombo()
    If Not bIsAlma Then
        Exit Sub
    End If
    LookupDialog.OperatorCombo.Clear
    sKey = GetAlmaSearchKey(LookupDialog.SearchFieldCombo.Value)
    bFound = False
    If IsEmpty(aExplainFields) Then
        aExplainFields = GetAllFields()
    End If
    If IsNull(aExplainFields) Then
        Exit Sub
    End If
    If UBound(aExplainFields) = 0 Then
        LookupDialog.OperatorCombo.Enabled = False
        Exit Sub
    End If
    sDefaultValue = "="
    For i = 0 To UBound(aExplainFields)
        If sKey = aExplainFields(i, 1) Then
            aOperators = aExplainFields(i, 2)
            For j = 0 To UBound(aOperators)
                sOperator = aOperators(j)
                If sOperator = "" Then
                    sOperator = "empty"
                End If
                LookupDialog.OperatorCombo.AddItem sOperator
            Next j
            bFound = True
        End If
        If bFound Then
            Exit For
        End If
    Next i
    If bFound Then
        LookupDialog.OperatorCombo.Value = sDefaultValue
    Else
        MsgBox ("Search field not set to a valid index name")
        LookupDialog.SearchFieldCombo.Value = "Keywords"
        PopulateOperatorCombo
    End If
End Sub

Function GetAlmaSearchKey(sPhrase As String) As String
    sKey = sPhrase
    For i = 0 To UBound(aAlmaSearchKeys)
        If sPhrase = aAlmaSearchKeys(i, 0) Then
            sKey = aAlmaSearchKeys(i, 1)
        End If
    Next i
    GetAlmaSearchKey = sKey
End Function

Sub RedrawButtons()
    With LookupDialog
        .ResultTypeList.Enabled = True
        .ResultTypeCombo.Style = fmStyleDropDownCombo
        .AddResultButton.Enabled = True
        If .ResultTypeList.ListCount > 0 And .ResultTypeList.ListIndex > -1 Then
            .RemoveResultButton.Enabled = True
            If .ResultTypeList.ListIndex > 0 Then
                .MoveUpButton.Enabled = True
            Else
                .MoveUpButton.Enabled = False
            End If
            If .ResultTypeList.ListIndex < .ResultTypeList.ListCount - 1 Then
                .MoveDownButton.Enabled = True
            Else
                .MoveDownButton.Enabled = False
            End If
        Else
            .RemoveResultButton.Enabled = False
            .MoveUpButton.Enabled = False
            .MoveDownButton.Enabled = False
        End If
        .NewSetButton.Enabled = True
        If .FieldSetList.ListCount > 0 And .FieldSetList.ListIndex > -1 Then
            .SaveSetButton.Enabled = True
            .LoadSetButton.Enabled = True
            .DeleteSetButton.Enabled = True
        Else
            .SaveSetButton.Enabled = False
            .LoadSetButton.Enabled = False
            .DeleteSetButton.Enabled = False
        End If
    End With
End Sub

'Determine the rightmost column containing data
Function FindLastColumn() As Integer
    Dim LastColumn As Integer
    If WorksheetFunction.CountA(Cells) > 0 Then
        'Search for any entry, by searching backwards by Columns.
        LastColumn = Cells.Find(What:="*", After:=[A1], _
            SearchOrder:=xlByColumns, _
            SearchDirection:=xlPrevious).Column
        FindLastColumn = LastColumn
    End If
End Function

'Converts a number to an Excel column index (A,B,C,....AA,AB, etc.)
Function ColumnLetterConvert(sInput As String) As String
    iVal = val(sInput)
    If iVal > 0 Then
        If iVal > 26 Then
            ColumnLetterConvert = Chr(Int((iVal - 1) / 26) + 64) & Chr(((iVal - 1) Mod 26) + 65)
        Else
            ColumnLetterConvert = Chr(iVal + 64)
        End If
    Else
        ColumnLetterConvert = "A"
    End If
End Function

Function EncodeURI(ByVal sStr As String) As String

    Dim i As Long
    Dim a As Long
    Dim res As String
    Dim code As String

res = ""
For i = 1 To Len(sStr)
    a = AscW(Mid(sStr, i, 1)) And &HFFFF&
    Select Case a
    Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
        code = Mid(sStr, i, 1)
    Case 32
        code = "%20" '"+"
    Case 0 To 127
        code = EncodeByte(CInt(a))
    Case 128 To 2047
        code = EncodeByte(((a \ 64) Or 192))
        code = code & EncodeByte(((a And 63) Or 128))
    Case Else
        code = EncodeByte(((a \ 4096) Or 224))
        code = code & EncodeByte((((a \ 64) And 63) Or 128))
        code = code & EncodeByte(((a And 63) Or 128))
    End Select
    res = res & code
Next i
EncodeURI = res
End Function

''=============================================================
'' URL encode single byte
''=============================================================
Private Function EncodeByte(val As Integer) As String
    Dim res As String
    res = "%" & Right("0" & Hex(val), 2)
    EncodeByte = res
End Function

Function ConstructURL(sBaseURL As String, sQuery1 As String, sSearchType As String, bAdvancedSearch As Boolean, oQueryRow As Range) As String
    sURL = sBaseURL & "?operation=searchRetrieve&version=1.2&maximumRecords=" & iMaximumRecords & "&query="
    
    iSearchTermsCount = 1
    If bAdvancedSearch Then
        iSearchTermsCount = LookupDialog.SearchListBox.ListCount
    End If
    
    Dim aSearchTerms() As Variant
    ReDim aSearchTerms(iSearchTermsCount, 4)
    
    Dim sIndex As String
    
    If Not bAdvancedSearch Then
        sIndex = GetAlmaSearchKey(sSearchType)
        aSearchTerms(0, 0) = ""
        aSearchTerms(0, 1) = sSearchType
        aSearchTerms(0, 2) = LookupDialog.OperatorCombo.Value
        If sSearchType = "alma.PermanentCallNumber" Then
            aSearchTerms(0, 2) = "all"
        End If
        aSearchTerms(0, 3) = sQuery1
    Else
        With LookupDialog.SearchListBox
            For i = 0 To iSearchTermsCount - 1
              aSearchTerms(i, 0) = .List(i, 0)
              aSearchTerms(i, 1) = .List(i, 1)
              aSearchTerms(i, 2) = .List(i, 2)
              aSearchTerms(i, 3) = .List(i, 3)
            Next i
        End With
    End If
    
    For i = 0 To UBound(aSearchTerms) - 1
        sBoolean = aSearchTerms(i, 0)
        sIndex = aSearchTerms(i, 1)
        sIndex = GetAlmaSearchKey(sIndex)
        sOperator = aSearchTerms(i, 2)
        Dim sQuery As String
        sQuery = CStr(aSearchTerms(i, 3))
        sQuery = GetColumnContents(oQueryRow, sQuery)
        sQuery = Replace(sQuery, "http://", "")
        sQuery = Replace(sQuery, "-", " ")
        sQuery = Replace(sQuery, "_", " ")
        If Trim(sQuery) = "" Then
            GoTo NextTerm
        End If
        If sBoolean <> "" Then
            sURL = sURL & "+" & sBoolean & "+"
        End If
        Select Case sIndex
            Case "alma.isbn"
                sQuery = NormalizeISBN(sQuery)
                If sQuery = "" Then
                    sQuery = "FALSE"
                End If
            Case "alma.issn"
                sQuery = NormalizeISSN(sQuery)
                If sQuery = "" Then
                    sQuery = "FALSE"
                End If
            Case "alma.barcode"
                sQuery = Replace(sQuery, " ", "")
        End Select
        Do While (InStr(1, sQuery, "||") > 0) Or (InStr(1, sQuery, "%7C%7C") > 0)
            sQuery = Replace(sQuery, "||", "|")
            sQuery = Replace(sQuery, "%7C%7C", "%7C")
        Loop
        If Right(sQuery, 1) = "|" Then
            sQuery = Left(sQuery, Len(sQuery) - 1)
        End If
        If Right(sQuery, 3) = "%7C" Then
            sQuery = Left(sQuery, Len(sQuery) - 3)
        End If
        If InStr(1, sQuery, "|") Or InStr(1, sQuery, "%7C") Then
            sQuery = Replace(sQuery, "|", "%22+OR+" & sIndex & "+" & sOperator & "+%22")
            sQuery = Replace(sQuery, "%7C", "%22+OR+" & sIndex & "+" & sOperator & "+%22")
            sURL = sURL & "(+" & sIndex & "+" & sOperator & "+%22" + EncodeURI(sQuery) + "%22+)"
        Else
            If sOperator = "empty" Then
                sURL = sURL & sIndex & "+==+%22%22"
            Else
                sURL = sURL & sIndex & "+" & sOperator & "+%22" + EncodeURI(sQuery) + "%22"
            End If
        End If
NextTerm:
    Next i
    If Not LookupDialog.IncludeSuppressed Then
        sURL = sURL & "+AND+alma.mms_tagSuppressed=false"
    End If
    ConstructURL = sURL
End Function


Function GenerateAuth(sUsername As String, sPassword As String)
    Set oB64Obj = CreateObject("MSXML2.DOMDocument")
    Set oB64Node = oB64Obj.createElement("b64")
    oB64Node.DataType = "bin.base64"
    sUserPass = sUsername & ":" & sPassword
    Dim yUserPassBytes() As Byte
    yUserPassBytes = StrConv(sUserPass, vbFromUnicode)
    oB64Node.nodeTypedValue = yUserPassBytes
    GenerateAuth = oB64Node.Text
End Function

Function DecodeAuth() As String
    DecodeAuth = ""
    If sAuth <> "" Then
        Set oB64Obj = CreateObject("MSXML2.DOMDocument")
        Set oB64Node = oB64Obj.createElement("b64")
        oB64Node.DataType = "bin.base64"
        oB64Node.Text = sAuth
        DecodeAuth = StrConv(oB64Node.nodeTypedValue, vbUnicode)
    End If
End Function

Function GetAllFields()
    If oXMLHTTP Is Nothing Then
        Initialize
    End If
    Dim sCatalogURL As String
    sCatalogURL = CStr(LookupDialog.CatalogURLBox.Text)
    bInvalidURL = False
    If Left(sCatalogURL, 4) <> "http" Then
        invalidURL = True
    End If
    bKeepTryingURL = True
    bNeedsAuthentication = False
    bIsoholdEnabled = False
    If Not bInvalidURL Then
        While bKeepTryingURL
            bInvalidURL = False
            sExplainURL = sCatalogURL & "?version=1.2&operation=explain"
            With oXMLHTTP
                .Open "GET", sExplainURL, True
                .setRequestHeader "Cache-Control", "no-cache,max-age=0"
                .setRequestHeader "pragma", "no-cache"
                If sAuth <> "" Then
                    .setRequestHeader "Authorization", "Basic " + sAuth
                End If
                .Send
                
                Do While .readyState <> 4
                    DoEvents
                Loop

                sResponse = .responseText
                bKeepTryingURL = False
                If .Status <> 200 Or InStr(sResponse, "explainResponse") = 0 Then
                    bInvalidURL = True
                End If
                If .Status = 401 Then
                    bNeedsAuthentication = True
                    bKeepTryingURL = True
                    bInvalidURL = True
                    UserPassForm.UserNameBox.Value = ""
                    UserPassForm.PasswordBox.Value = ""
                    UserPassForm.Show
                    If bKeepTryingURL Then
                        sAuth = GenerateAuth(UserPassForm.UserNameBox.Value, UserPassForm.PasswordBox.Value)
                    End If
                End If
                If .Status = 200 And sAuth <> "" And UserPassForm.RememberCheckbox.Value Then
                    SaveCatalogAuthToRegistry
                End If
            End With
        Wend
    End If
    If bInvalidURL Then
        If bNeedsAuthentication Then
            MsgBox ("Cannot log in to catalog.")
        Else
            MsgBox ("Cannot access catalog.  Please confirm the Alma URL is correct.")
        End If
        GetAllFields = Null
        Exit Function
    End If
    'Check if ISO Holdings are enabled
    sExplainURL = sExplainURL & "&recordSchema=isohold"
    With oXMLHTTP
        .Open "GET", sExplainURL, True
        If sAuth <> "" Then
            .setRequestHeader "Authorization", "Basic " + sAuth
        End If
        .Send
                
        Do While .readyState <> 4
            DoEvents
        Loop
        If .Status = 200 Then
            bIsoholdEnabled = True
        End If
        
    End With
    
    
    oXMLDOM.SetProperty "SelectionNamespaces", "xmlns:xr='http://www.loc.gov/zing/srw/' " & _
        "xmlns:xpl='http://explain.z3950.org/dtd/2.0/' " & _
        "xmlns:ns='http://explain.z3950.org/dtd/2.1/'"
    oXMLDOM.LoadXML (sResponse)
    sFields = ""
    Set aFields = oXMLDOM.SelectNodes("xr:explainResponse/xr:record/xr:recordData/" & _
        "xpl:explain/xpl:indexInfo/xpl:index")
    Dim aFieldMap() As Variant
    ReDim aFieldMap(aFields.length - 1, 2)
    For i = 0 To aFields.length - 1
        sLabel = aFields(i).SelectSingleNode("ns:title").Text
        sIndexCode = aFields(i).SelectSingleNode("xpl:map/xpl:name").Text
        sIndexSet = aFields(i).SelectSingleNode("xpl:map/xpl:name/@set").Text
        Dim oOperatorGroup As Variant
        Set oOperatorGroup = aFields(i).SelectNodes("xpl:configInfo/xpl:supports")
        Dim aOperators() As String
        ReDim aOperators(oOperatorGroup.length - 1)
        For j = 0 To oOperatorGroup.length - 1
            aOperators(j) = oOperatorGroup.Item(j).Text
        Next j

        aFieldMap(i, 0) = sLabel
        aFieldMap(i, 1) = sIndexSet & "." & sIndexCode
        aFieldMap(i, 2) = aOperators
    Next i
    
    For i = 0 To UBound(aFieldMap)
        For j = i + 1 To UBound(aFieldMap)
            SearchI = Replace(UCase(aFieldMap(i, 0)), "(", "")
            SearchJ = Replace(UCase(aFieldMap(j, 0)), "(", "")
            If UCase(SearchI > SearchJ) Then
                t1 = aFieldMap(j, 0)
                t2 = aFieldMap(j, 1)
                t3 = aFieldMap(j, 2)
                aFieldMap(j, 0) = aFieldMap(i, 0)
                aFieldMap(j, 1) = aFieldMap(i, 1)
                aFieldMap(j, 2) = aFieldMap(i, 2)
                aFieldMap(i, 0) = t1
                aFieldMap(i, 1) = t2
                aFieldMap(i, 2) = t3
            End If
        Next j
    Next i
    GetAllFields = aFieldMap
End Function

Function Z3950Connect(sSource As String) As Boolean
    bKeepTryingURL = True
    sZHost = ""
    sZPort = ""
    sZDB = ""
    sZUserName = ""
    sZPassword = ""
    If sSource = "source:worldcat" Then
        sZHost = sWCZhost
        sZPort = 210
        sZDB = sWCZDB
    End If
    
    sUserPass = DecodeAuth()
    iDelimPos = InStr(1, sUserPass, ":")
    If iDelimPos > 0 Then
        sZUserName = Left(sUserPass, iDelimPos - 1)
        sZPassword = Mid(sUserPass, iDelimPos + 1)
    End If
    
    bValidConnection = False
    While bKeepTryingURL And Not bValidConnection
        oZConn = ZOOM_connection_create(0)
        ZOOM_connection_option_set oZConn, "databaseName", sZDB
        ZOOM_connection_option_set oZConn, "preferredRecordSyntax", "USmarc"
        ZOOM_connection_option_set oZConn, "elementSetName", "FA"
        ZOOM_connection_option_set oZConn, "largeSetLowerBound", "10000"
        ZOOM_connection_option_set oZConn, "user", sZUserName
        ZOOM_connection_option_set oZConn, "password", sZPassword
        ZOOM_connection_connect oZConn, sZHost, sZPort
        errcode = ZOOM_connection_errcode(oZConn)
        If errcode > 0 And bKeepTryingURL Then
            UserPassForm.Show
            sZUserName = UserPassForm.UserNameBox.Value
            sZPassword = UserPassForm.PasswordBox.Value
            ZOOM_connection_destroy (oZConn)
            oZConn = 0
        Else
            bValidConnection = True
            If sZUserName <> "" And sZPassword <> "" And UserPassForm.RememberCheckbox Then
                sAuth = GenerateAuth(UserPassForm.UserNameBox.Value, UserPassForm.PasswordBox.Value)
                SaveCatalogAuthToRegistry
            End If
        End If
    Wend
    
    If Not bValidConnection Then
        Z3950Connect = False
        MsgBox ("Cannot connect to catalog.")
        ZOOM_connection_destroy (oZConn)
        oZConn = 0
        Exit Function
    End If
    Z3950Connect = True
End Function

Function Z3950Search(sQuery As String, sSearchType As String, sSource As String)
    If oZConn = 0 Then
        bSuccess = Z3950Connect(sSource)
        If Not bSuccess Then
            Z3950Search = ""
            Exit Function
        End If
    End If
    
    oConverter.Open
    oConverter.Charset = "UTF-8"
    oConverter.Type = 2
    oConverter.WriteText sQuery
    oConverter.Position = 0
    oConverter.Charset = "ISO-8859-1"
    sQuery = oConverter.ReadText
    sQuery = Replace(sQuery, ChrW(239) & ChrW(187) & ChrW(191), "") 'BOM
    oConverter.Close
    
    sSearchIndex = "1016"
    If sSearchType = "Title" Then
        sSearchIndex = "4"
    ElseIf sSearchType = "ISBN" Then
        sSearchIndex = "7"
        sQuery = NormalizeISBN(sQuery)
    ElseIf sSearchType = "ISSN" Then
        sSearchIndex = "8"
        sQuery = NormalizeISSN(sQuery)
    ElseIf sSearchType = "OCLC No." Then
        sSearchIndex = "12"
        sQuery = NormalizeOCLC(sQuery)
    End If
    
    sQuery = Replace(sQuery, """", "\""")
    
    sCQLQuery = ""
    If sSearchType = "Z-Author-Title-Date" Then
        oRegEx.Pattern = "^AUTHOR *= *(.*) AND TITLE *= *(.*) AND YEAR *= *(.*)"
        If oRegEx.Test(sQuery) Then
            Set oFields = oRegEx.Execute(sQuery)
            sAuthor = oFields(0).Submatches(0)
            sAuthor = Replace(sAuthor, "*", "?")
            sTitle = oFields(0).Submatches(1)
            sTitle = Replace(sTitle, "*", "?")
            sYear = oFields(0).Submatches(2)
            sCQLQuery = "@and @attr 1=31 " & sYear & " @and " & _
                "@attr 1=1 @attr 3=1 @attr 4=1 """ & sAuthor & """ " & _
                "@attr 1=4 @attr 3=1 @attr 4=1 """ & sTitle & """"
        Else
            sCQLQuery = "@attr 4=1 @attr 1=1016 " & sQuery
        End If
    ElseIf sSearchType = "Z-Title-Date" Then
        oRegEx.Pattern = "^TITLE *= *(.*) AND YEAR *= *(.*)"
        If oRegEx.Test(sQuery) Then
            Set oFields = oRegEx.Execute(sQuery)
            sTitle = oFields(0).Submatches(0)
            sTitle = Replace(sTitle, "*", "?")
            sYear = oFields(0).Submatches(1)
            sCQLQuery = "@and @attr 1=31 " & sYear & " " & _
                "@attr 1=4 @attr 3=1 @attr 4=1 """ & sTitle & """"
        Else
            sCQLQuery = "@attr 4=1 @attr 1=1016 " & sQuery
        End If
    Else
        aSearchKeys = Split(sQuery, "|")
        For i = 0 To UBound(aSearchKeys)
            If sCQLQuery <> "" Then
                sCQLQuery = "@or " & sCQLQuery
            End If
            sCQLQuery = sCQLQuery & "@attr 4=1 @attr 1=" & sSearchIndex & " """ & aSearchKeys(i) & """"
        Next i
    End If
    
    zrs = ZOOM_connection_search_pqf(oZConn, sCQLQuery)
    ZOOM_resultset_option_set zrs, "count", iMaximumRecords
    zcount = ZOOM_resultset_size(zrs)
    If zcount > 0 Then
        sAllRecords = "<searchRetrieveResponse xmlns=""http://www.loc.gov/zing/srw/""><records>"
        If zcount > iMaximumRecords Then
            zcount = iMaximumRecords
        End If
        For i = 0 To zcount - 1
            Dim zptr As LongPtr
            Dim zsize As Long
            zrec = ZOOM_resultset_record(zrs, i)
            zptr = ZOOM_record_get(zrec, "xml;charset=marc8,utf8", zsize)
            If zsize = 0 Then
                ZOOM_resultset_option_set zrs, "elementSetName", "F"
                zrec = ZOOM_resultset_record(zrs, i)
                zptr = ZOOM_record_get(zrec, "xml;charset=marc8,utf8", zsize)
                ZOOM_resultset_option_set zrs, "elementSetName", "FA"
            End If
            Dim recBytes() As Byte
            ReDim recBytes(zsize)
            CopyMemory recBytes(0), ByVal zptr, zsize
            If zsize > 0 Then
                ReDim Preserve recBytes(zsize - 1) 'remove null terminator
            End If
            sResultXML = StrConv(recBytes, vbUnicode)
            oConverter.Open
            oConverter.Type = 1
            oConverter.Write recBytes
            oConverter.Position = 0
            oConverter.Type = 2
            oConverter.Charset = "UTF-8"
            sResultXML = oConverter.ReadText
            oConverter.Close
        
            sResultXML = Replace(sResultXML, "<record", "<record><recordData><record")
            sResultXML = Replace(sResultXML, "</record>", "</record></recordData></record>")
            sResultXML = Replace(sResultXML, Chr(10), "")
            sAllRecords = sAllRecords & sResultXML
        Next i
        sAllRecords = sAllRecords & "</records></searchRetrieveResponse>"
    End If
    Z3950Search = sAllRecords
    ZOOM_resultset_destroy (zrs)
End Function

Function GetColumnContents(ByVal oRow As Range, sValue As String) As String
    oRegEx.Pattern = "^\[\[[A-Z]+\]\]$"
    If oRegEx.Test(sValue) Then
        sValue = Replace(sValue, "[[", "")
        sValue = Replace(sValue, "]]", "")
        sValue = CStr(oRow.Cells(1, Cells(1, sValue).Column).Value)
    End If
    sValue = Replace(sValue, ChrW(160), " ")
    sValue = Replace(sValue, ChrW(166), "|")
    sValue = Trim(sValue)
    GetColumnContents = sValue
End Function

Function Lookup(ByVal sQueryRow As Range, sCatalogURL As String) As String
    If oXMLHTTP Is Nothing Then
        Initialize
    End If
    
    Dim bAdvancedSearch As Boolean
    bAdvancedSearch = False
    If bIsAlma And LookupDialog.SearchListBox.ListCount > 0 Then
        bAdvancedSearch = True
    End If
    
    Dim sQuery1 As String
    sQuery1 = ""
    
    If Not bAdvancedSearch Then
        Dim sSearchString As String
        sSearchString = GetColumnContents(sQueryRow, LookupDialog.SearchValueBox.Value)
        If sSearchString = "FALSE" Or sSearchString = "" Then
            Lookup = ""
            Exit Function
        End If
        sQuery1 = sSearchString
    End If
    
    Dim sSearchType As String
    sSearchType = CStr(LookupDialog.SearchFieldCombo.Value)
    Dim sFormat As String
    sURL = ""
    
    If LookupDialog.ValidateCheckBox.Value And Not bAdvancedSearch Then
        If sSearchType = "ISBN" Then
            Dim sISBN As String
            sISBN = NormalizeISBN(sQuery1)
            iVbarPos = InStr(1, sISBN, "|")
            If iVbarPos > 0 Then
                sISBN = Left(sISBN, iVbarPos - 1)
            End If
            If sISBN = "INVALID" Or sISBN <> GenerateCheckDigit(sISBN) Then
                Lookup = "INVALID"
                Exit Function
            End If
        ElseIf sSearchType = "ISSN" Then
            Dim sISSN As String
            sISSN = NormalizeISSN(sQuery1)
            If sISSN = "INVALID" Or sISSN <> GenerateCheckDigit(sISSN) Then
                Lookup = "INVALID"
                Exit Function
            End If
        End If
    End If
      
    If sCatalogURL = "source:recap" Then
        Select Case sSearchType
            Case "ISBN"
                sQuery1 = "isbn_s:" & Replace(NormalizeISBN(sQuery1), "|", "+OR+")
            Case "Title"
                sQuery1 = "%22" & EncodeURI(sQuery1) & "%22&search_field=title"
            Case "OCLC No."
                sQuery1 = "oclc_s:" & Replace(NormalizeOCLC(sQuery1), "|", "+OR+")
            Case "LCCN"
                sQuery1 = "lccn_s:" & Replace(sQuery1, "|", "+OR+")
            Case Else
                sQuery1 = "%22" & EncodeURI(sQuery1) & "%22"
        End Select
        sQuery1 = sQuery1 & "&per_page=" & iMaximumRecords
        If sQuery1 = "" Then
            sQuery1 = False
        Else 'Throttle ReCAP queries to one per second
            Application.Wait (Now() + TimeValue("0:00:01"))
        End If
        sURL = sBlacklightURL & sQuery1
    ElseIf sCatalogURL = "source:borrowdirect" Then
        sURL = sIPLCReshareURL & "%22" & EncodeURI(sQuery1) & "%22&limit=" & iMaximumRecords
    ElseIf sCatalogURL = "source:lccat" Then
        sURL = sLCCatURL & "?version=1.1&operation=searchRetrieve" & _
            "&maximumRecords=" & iMaximumRecords & "&recordSchema=marcxml&query=%22" & sQuery1 & "%22"
    ElseIf sCatalogURL = "source:worldcat" Then
        sURL = "z3950"
    Else
        sURL = ConstructURL(sCatalogURL, sQuery1, sSearchType, bAdvancedSearch, sQueryRow)
    End If
    sHoldingsURL = Replace(sURL, "&query", "&recordSchema=isohold&query")
    sResponse = ""

    If sURL = "z3950" Then
        sResponse = Z3950Search(sQuery1, sSearchType, sCatalogURL)
    Else
        With oXMLHTTP
            .Open "GET", sURL, True
            If sAuth <> "" Then
                .setRequestHeader "Authorization", "Basic " + sAuth
            End If
            .Send
            Do While .readyState <> 4
                DoEvents
            Loop
            sResponse = .responseText
            sHoldings = ""
            If sCatalogURL = "source:borrowdirect" Then
                sAllRecords = "<searchRetrieveResponse xmlns=""http://www.loc.gov/zing/srw/""><records>"
                iXMLstart = InStr(1, sResponse, "<record")
                While iXMLstart > 0
                    If iXMLstart < 1 Then
                        sResponse = ""
                    Else
                        sResponse = Mid(sResponse, iXMLstart)
                        iXMLend = InStr(1, sResponse, "<\/collection>")
                        sThisRecord = Left(sResponse, iXMLend - 1)
                        sResponse = Mid(sResponse, iXMLend)
                        sThisRecord = Replace(sThisRecord, "<record>", _
                        "<record><recordData><record xmlns=""http://www.loc.gov/MARC21/slim"">")
                        sThisRecord = Replace(sThisRecord, "<\/record>", "<\/record><\/recordData><\/record>")
                        sThisRecord = Replace(sThisRecord, "\n", "")
                        sThisRecord = Replace(sThisRecord, "\""", """")
                        sThisRecord = Replace(sThisRecord, "\/", "/")
                        sAllRecords = sAllRecords & sThisRecord
                    End If
                    iXMLstart = InStr(1, sResponse, "<record")
                Wend
                sAllRecords = sAllRecords & "</records></searchRetrieveResponse>"
                sResponse = sAllRecords
            End If
            If bIsoholdEnabled Then
                .Open "GET", sHoldingsURL, True
                If sAuth <> "" Then
                   .setRequestHeader "Authorization", "Basic " + sAuth
                End If
                .Send
                Do While .readyState <> 4
                    DoEvents
                Loop
                sHoldings = .responseText
                If InStr(1, sHoldings, "searchRetrieveResponse") = 0 Then
                    bIsoholdEnabled = False
                End If
            End If
        End With
    End If
    Lookup = sResponse & sHoldings
End Function

Function ExtractField(sResultTypeAll As String, sResultXML As String, bHoldings As Boolean, Optional sBarcode As Variant) As String
    aResultFields = Split(sResultTypeAll, "|", -1, 0)
    iResultTypes = UBound(aResultFields)
    sBasePath = ""
    If bHoldings Then
        oXMLDOM.SetProperty "SelectionNamespaces", "xmlns:sr='http://www.loc.gov/zing/srw/' " & _
            "xmlns:hold='http://www.loc.gov/standards/iso20775/'"
        sBasePath = "sr:searchRetrieveResponse/sr:records/sr:record/sr:recordData/hold:holdings"
    Else
        oXMLDOM.SetProperty "SelectionNamespaces", "xmlns:sr='http://www.loc.gov/zing/srw/' " & _
            "xmlns:marc='http://www.loc.gov/MARC21/slim'"
        sBasePath = "sr:searchRetrieveResponse/sr:records/sr:record/sr:recordData/marc:record"

    End If
    If LookupDialog.CatalogURLBox.Value = "source:recap" Then
        oRegEx.Global = True
        oRegEx.Pattern = "[\[,]{""id"":""([^""]*)"""
        
        sResultJSON = sResultXML
        sResultString = ""
        iCurrentPos = 1
        Set oRecords = oRegEx.Execute(sResultJSON)
        iRecords = oRecords.Count
        If iRecords = 0 Then
            ExtractField = "FALSE"
            Exit Function
        End If
        For Each m In oRecords
            sID = m.Submatches(0)
            sResultJSON = Mid(sResultJSON, InStr(1, sResultJSON, m.Submatches(0)))
            iRecLength = InStr(Len(m.Submatches(0)), sResultJSON, "},{""id""")
            If iRecLength > 0 Then
                sCurrentRecord = Left(sResultJSON, InStr(Len(m.Submatches(0)), sResultJSON, "},{""id"""))
            Else
                sCurrentRecord = sResultJSON
            End If
            For h = 0 To UBound(aResultFields)
                If ExtractField <> "" And Right(ExtractField, 1) <> "|" Then
                    ExtractField = ExtractField & "|"
                End If
                sResultType = aResultFields(h)
                Select Case sResultType
                    Case "exists"
                        ExtractField = "TRUE "
                    Case "001"
                        ExtractField = ExtractField & sID
                    Case "010"
                        oRegEx.Pattern = "\[([^\]]*)\],""label"":""Lccn S"""
                        Set oLCCNs = oRegEx.Execute(sCurrentRecord)
                        If oLCCNs.Count > 0 Then
                            sLCCNs = oLCCNs(0).Submatches(0)
                            sLCCNs = Replace(sLCCNs, """,""", ChrW(166))
                            sLCCNs = Replace(sLCCNs, """", "")
                            ExtractField = ExtractField & sLCCNs
                        Else
                            ExtractField = ExtractField & " "
                        End If
                    Case "020"
                        oRegEx.Pattern = "\[([^\]]*)\],""label"":""Isbn S"""
                        Set oISBNs = oRegEx.Execute(sCurrentRecord)
                        If oISBNs.Count > 0 Then
                            sISBNs = oISBNs(0).Submatches(0)
                            sISBNs = Replace(sISBNs, """,""", ChrW(166))
                            sISBNs = Replace(sISBNs, """", "")
                            ExtractField = ExtractField & sISBNs
                        Else
                            ExtractField = ExtractField & " "
                        End If
                    Case "035$a#(OCoLC)"
                        oRegEx.Pattern = "\[([^\]]*)\],""label"":""Oclc S"""
                        Set oOCLCs = oRegEx.Execute(sCurrentRecord)
                        If oOCLCs.Count > 0 Then
                            sOCLCs = oOCLCs(0).Submatches(0)
                            sOCLCs = Replace(sOCLCs, """,""", ChrW(166))
                            sOCLCs = Replace(sOCLCs, """", "")
                            ExtractField = ExtractField & sOCLCs
                        Else
                            ExtractField = ExtractField & " "
                        End If
                    Case "245"
                        oRegEx.Pattern = """attributes"":{""title"":""([^""]*)"""
                        Set oTitle = oRegEx.Execute(sCurrentRecord)
                        If oTitle.Count > 0 Then
                            ExtractField = ExtractField & oTitle(0).Submatches(0)
                        End If
                    Case "recap"
                        oRegEx.Pattern = """location_code"":""([^""]*)"""
                        Set oLoc = oRegEx.Execute(sCurrentRecord)
                        oRegEx.Pattern = """location"":""([^""]*)"""
                        Set oLocName = oRegEx.Execute(sCurrentRecord)
                        If oLoc.Count > 0 Then
                            sLoc = oLoc(0).Submatches(0)
                            Select Case sLoc
                                Case "scsbhl"
                                    sLoc = "Harvard"
                                Case "scsbnypl"
                                    sLoc = "NYPL"
                                Case "scsbcul"
                                    sLoc = "Columbia"
                                Case Else
                                    For i = 0 To oLocName.Count - 1
                                        If InStr(1, oLocName(i).Submatches(0), "Remote Storage") > 0 Then
                                            sLoc = "Princeton"
                                            Exit For
                                        End If
                                    Next i
                                    If sLoc <> "Princeton" Then
                                        sLoc = ""
                                    End If
                                    
                            End Select
                        End If
                        If InStr(1, ExtractField, sLoc) = 0 Then
                            ExtractField = ExtractField & sLoc
                        Else
                            ExtractField = Left(ExtractField, Len(ExtractField) - 1)
                        End If
                    Case "recap_cgd"
                        oRegEx.Pattern = "(?:""location_code"":""([^""]*)""[^}]*)?(?:""description"":""([^""]*)""[^}]*)?(?:""use_statement"":""([^""]*)""[^}]*)?""cgd"":""([^""]*)""[^}]*""collection_code"":""([^""]*)"""
                        sCGD = ""
                        Set oCGD = oRegEx.Execute(sCurrentRecord)
                        sRecapLoc = ""
                        For i = 0 To oCGD.Count - 1
                            If oCGD(i).Submatches(0) <> "" Then
                                sRecapLoc = oCGD(i).Submatches(0)
                                sRecapLoc = Replace(sRecapLoc, "scsb", "")
                            End If
                            If sCGD <> "" Then
                                sCGD = sCGD & ChrW(166)
                            End If
                            sCGD = sCGD & sRecapLoc & "-" & oCGD(i).Submatches(4) & "-" & oCGD(i).Submatches(3)
                            If oCGD(i).Submatches(2) <> "" Then
                                sCGD = sCGD & "-" & oCGD(i).Submatches(2)
                            End If
                            If oCGD(i).Submatches(1) <> "" Then
                                sCGD = sCGD & "-" & oCGD(i).Submatches(1)
                            End If
                        Next i
                        ExtractField = ExtractField & sCGD
                    Case Else
                        ExtractField = "ERROR:InvalidRecap"
                        Exit Function
                End Select
            Next h
        Next m
        Exit Function
    End If
    
    oXMLDOM.LoadXML (sResultXML)
    Set aRecords = oXMLDOM.SelectNodes(sBasePath)
        
    iRecords = aRecords.length
    If iRecords = 0 Then
        ExtractField = "FALSE"
        Exit Function
    End If
    
    ExtractField = ""
       
    'Iterate through results, compile result string
    For i = 0 To iRecords - 1
        If oXMLDOM.parseError.ErrorCode = 0 Then
           sRecord = ""
           For h = 0 To UBound(aResultFields)
              If ExtractField <> "" And Right(ExtractField, 1) <> "|" Then
                 ExtractField = ExtractField & ChrW(166)
              End If
              sResultType = aResultFields(h)
              sResultFilter = ""
              iFilterPos = InStr(1, sResultType, "#")
              If iFilterPos > 0 Then
                sResultFilter = Mid(sResultType, iFilterPos + 1)
                sResultType = Left(sResultType, iFilterPos - 1)
              End If
              
              iSubStartPos = -1
              iSubLength = -1
              oRegEx.Pattern = "\(([0-9]+),([0-9]+)\)$"
              Set oMatch = oRegEx.Execute(sResultType)
              If oMatch.Count = 1 Then
                iSubStartPos = oMatch(0).Submatches(0)
                iSubLength = oMatch(0).Submatches(1)
                If iSubLength = 0 Then
                    iSubLength = 9999
                End If
                sResultType = Left(sResultType, Len(sResultType) - Len(oMatch.Item(0)))
              End If
              
              sBibPrefix = "marc:datafield"
               If sResultType = "000" Then
                  sBibPrefix = "marc:leader"
               ElseIf Left(sResultType, 2) Like "00" Then
                  sBibPrefix = "marc:controlfield"
               End If
               sHoldingsPrefix1 = "hold:holding/hold:holdingSimple/hold:copyInformation"
               sHoldingsPrefix2 = "hold:holding/hold:holdingStructured/hold:set/hold:component"
               Dim oFieldList As IXMLDOMNodeList
               Select Case sResultType
                  Case "exists"
                     ExtractField = "TRUE "
                  Case "Barcode"
                     Set oFieldList = aRecords(i).SelectNodes(sHoldingsPrefix1 & _
                            "/hold:pieceIdentifier/hold:value")
                     If oFieldList.length = 0 Then
                        Set oFieldList = aRecords(i).SelectNodes(sHoldingsPrefix2 & _
                            "/hold:pieceIdentifier/hold:value")
                     End If
                     For j = 0 To oFieldList.length - 1
                        If sRecord <> "" Then
                            sRecord = sRecord & ChrW(166)
                        End If
                        sRecord = sRecord & oFieldList.Item(j).XML
                        oRegEx.Pattern = "<[^>]*>"
                        sRecord = oRegEx.Replace(sRecord, "")
                    Next j
                    ExtractField = ExtractField & sRecord
                  Case "Item Location"
                        Set oHoldings = aRecords(i).SelectNodes("hold:holding")
                        For j = 0 To oHoldings.length - 1
                            sRecord = ""
                            Set oLibraryCode = oHoldings(j).SelectNodes("hold:physicalLocation")
                            sLibraryCode = ""
                            If oLibraryCode.length = 1 Then
                               sLibraryCode = oLibraryCode.Item(0).Text
                            End If
                            If Not IsMissing(sBarcode) Then
                                Set oFieldList = oHoldings(j).SelectNodes(Replace(sHoldingsPrefix1, "hold:holding/", "") & _
                                    "[hold:pieceIdentifier/hold:value='" & sBarcode & "']/hold:sublocation")
                                If oFieldList.length = 0 Then
                                   Set oFieldList = oHoldings(j).SelectNodes(Replace(sHoldingsPrefix2, "hold:holding/", "") & _
                                    "[hold:pieceIdentifier/hold:value='" & sBarcode & "']/hold:sublocation")
                                End If
                            Else
                                Set oFieldList = oHoldings(j).SelectNodes(Replace(sHoldingsPrefix1, "hold:holding/", "") & "/hold:sublocation")
                                If oFieldList.length = 0 Then
                                   Set oFieldList = oHoldings(j).SelectNodes(Replace(sHoldingsPrefix2, "hold:holding/", "") & "/hold:sublocation")
                                End If
                            End If
                            For k = 0 To oFieldList.length - 1
                                If sRecord <> "" Then
                                    sRecord = sRecord & ChrW(166)
                                End If
                                sRecord = sRecord & sLibraryCode & " " & oFieldList.Item(k).XML
                                oRegEx.Pattern = "<[^>]*>"
                                sRecord = oRegEx.Replace(sRecord, "")
                            Next k
                            If ExtractField <> "" And sRecord <> "" Then
                                ExtractField = ExtractField & "|"
                            End If
                            ExtractField = ExtractField & sRecord
                        Next j
                  Case "Item Enum/Chron"
                        If Not IsMissing(sBarcode) Then
                            Set oFieldList = aRecords(i).SelectNodes(sHoldingsPrefix1 & _
                                "[hold:pieceIdentifier/hold:value='" & sBarcode & "']/hold:enumerationAndChronology/hold:text")
                            If oFieldList.length = 0 Then
                                Set oFieldList = aRecords(i).SelectNodes(sHoldingsPrefix2 & _
                                    "[hold:pieceIdentifier/hold:value='" & sBarcode & "']/hold:enumerationAndChronology/hold:text")
                            End If
                        Else
                            Set oFieldList = aRecords(i).SelectNodes(sHoldingsPrefix1 & "/hold:enumerationAndChronology/hold:text")
                            If oFieldList.length = 0 Then
                                Set oFieldList = aRecords(i).SelectNodes(sHoldingsPrefix2 & "/hold:enumerationAndChronology/hold:text")
                            End If
                        End If
                        For j = 0 To oFieldList.length - 1
                            If sRecord <> "" Then
                                sRecord = sRecord & ChrW(166)
                            End If
                            sRecord = sRecord & oFieldList.Item(j).XML
                            oRegEx.Pattern = "<[^>]*>"
                            sRecord = oRegEx.Replace(sRecord, "")
                        Next j
                        ExtractField = ExtractField & sRecord
                  Case "Shelf Locator"
                        If Not IsMissing(sBarcode) Then
                            Set oFieldList = aRecords(i).SelectNodes(sHoldingsPrefix1 & _
                                "[hold:pieceIdentifier/hold:value='" & sBarcode & "']/hold:shelfLocator")
                            If oFieldList.length = 0 Then
                                Set oFieldList = aRecords(i).SelectNodes(sHoldingsPrefix2 & _
                                    "[hold:pieceIdentifier/hold:value='" & sBarcode & "']/hold:shelfLocator")
                            End If
                        Else
                            Set oFieldList = aRecords(i).SelectNodes(sHoldingsPrefix1 & "/hold:shelfLocator")
                            If oFieldList.length = 0 Then
                                Set oFieldList = aRecords(i).SelectNodes(sHoldingsPrefix2 & "/hold:shelfLocator")
                            End If
                        End If
                        For j = 0 To oFieldList.length - 1
                            If sRecord <> "" Then
                                sRecord = sRecord & ChrW(166)
                            End If
                            sRecord = sRecord & oFieldList.Item(j).XML
                            oRegEx.Pattern = "<[^>]*>"
                            sRecord = oRegEx.Replace(sRecord, "")
                        Next j
                        ExtractField = ExtractField & sRecord
                  Case "000" To "999z", "AVA" To "AVAz", "AVD" To "AVDz", "AVE" To "AVEz"
                     If sResultType = "000" Then
                       Set oFieldList = aRecords(i).SelectNodes(sBibPrefix)
                       sRecord = oFieldList.Item(0).XML
                       oRegEx.Pattern = "<[^>]*>"
                       sRecord = oRegEx.Replace(sRecord, "")
                     ElseIf sResultType Like "###" Then
                       Set oFieldList = aRecords(i).SelectNodes(sBibPrefix & "[@tag='" & sResultType & "']")
                       For j = 0 To oFieldList.length - 1
                         If sRecord <> "" Then
                            sRecord = sRecord & ChrW(166)
                         End If
                         sRecord = sRecord & oFieldList.Item(j).XML
                       Next j
                    
                      If LookupDialog.IncludeExtrasCheckBox.Value = True Then
                           oRegEx.Pattern = "<subfield code=.(.).>"
                           sRecord = oRegEx.Replace(sRecord, "$$$1 ")
                           sRecord = Replace(sRecord, "ind1="" """, "ind1=""_""")
                           sRecord = Replace(sRecord, "ind2="" """, "ind2=""_""")
                           oRegEx.Pattern = "<datafield[^>]*\s*ind1=.(.).\s*ind2=.(.).[^>]*>"
                           sRecord = oRegEx.Replace(sRecord, "$1$2")
                       Else
                          oRegEx.Pattern = "<subfield code=.6.>[^<]*</subfield>"
                          sRecord = oRegEx.Replace(sRecord, "")
                       End If
                       oRegEx.Pattern = "<[^>]*>"
                       sRecord = oRegEx.Replace(sRecord, " ")
                       oRegEx.Pattern = "^\s+"
                       sRecord = oRegEx.Replace(sRecord, "")
                       If Not Left(sResultType, 2) Like "00" Then
                          oRegEx.Pattern = "\s+$"
                          sRecord = oRegEx.Replace(sRecord, "")
                          oRegEx.Pattern = "\s\s+"
                          sRecord = oRegEx.Replace(sRecord, " ")
                       End If
                     ElseIf sResultType Like "###-880" Then
                        sMainField = Left(sResultType, 3)
                        Set oFieldList = aRecords(i).SelectNodes(sBibPrefix & "[@tag='880'][marc:subfield[@code='6' and starts-with(text(),'" & sMainField & "')]]")
                        For j = 0 To oFieldList.length - 1
                          If sRecord <> "" Then
                            sRecord = sRecord & ChrW(166)
                          End If
                          sRecord = sRecord & oFieldList.Item(j).XML
                        Next j
                        
                        If LookupDialog.IncludeExtrasCheckBox.Value = True Then
                           oRegEx.Pattern = "<subfield code=.(.).>"
                           sRecord = oRegEx.Replace(sRecord, "$$$1 ")
                           sRecord = Replace(sRecord, "ind1="" """, "ind1=""_""")
                           sRecord = Replace(sRecord, "ind2="" """, "ind2=""_""")
                           oRegEx.Pattern = "<datafield[^>]*\s*ind1=.(.).\s*ind2=.(.).[^>]*>"
                           sRecord = oRegEx.Replace(sRecord, "$1$2")
                        Else
                           oRegEx.Pattern = "<subfield code=.6.>[^<]*</subfield>"
                           sRecord = oRegEx.Replace(sRecord, "")
                        End If
                        
                        oRegEx.Pattern = "<[^>]*>"
                        sRecord = oRegEx.Replace(sRecord, " ")
                        oRegEx.Pattern = "^\s+"
                        sRecord = oRegEx.Replace(sRecord, "")
                        If Not Left(sResultType, 2) Like "00" Then
                           oRegEx.Pattern = "\s+$"
                           sRecord = oRegEx.Replace(sRecord, "")
                           oRegEx.Pattern = "\s\s+"
                           sRecord = oRegEx.Replace(sRecord, " ")
                        End If
                     ElseIf sResultType Like "###$?*" Or sResultType Like "AVA$?*" _
                        Or sResultType Like "AVD$?*" Or sResultType Like "AVE$?*" Then
                        
                        sMainField = Left(sResultType, 3)
                        sSubfield = Mid(sResultType, 5, 99)
                        sSubfieldQuery = "[@code='"
                        For j = 1 To Len(sSubfield)
                           If j > 1 Then
                             sSubfieldQuery = sSubfieldQuery & "' or @code='"
                           End If
                           sSubfieldQuery = sSubfieldQuery & Mid(sSubfield, j, 1)
                        Next j
                        sSubfieldQuery = sSubfieldQuery & "']"
                    
                        Set oFieldList = aRecords(i).SelectNodes(sBibPrefix & "[@tag='" & sMainField & "']")
                        For j = 0 To oFieldList.length - 1
                           If sRecord <> "" And Right(sRecord, 1) <> ChrW(166) Then
                             sRecord = sRecord & ChrW(166)
                           End If
                           Set oSubfieldList = oFieldList.Item(j).SelectNodes("marc:subfield" & sSubfieldQuery)
                           For k = 0 To oSubfieldList.length - 1
                             sRecord = sRecord & oSubfieldList.Item(k).XML
                           Next k
                        Next j
                        oRegEx.Pattern = "<[^>]*>"
                        sRecord = oRegEx.Replace(sRecord, " ")
                        oRegEx.Pattern = "  *"
                        sRecord = oRegEx.Replace(sRecord, " ")
                    ElseIf sResultType Like "###-880$?*" Then
                       sField = Left(sResultType, 3)
                       sSubfield = Mid(sResultType, 9, 99)
                       sSubfieldQuery = "[@code='"
                       For j = 1 To Len(sSubfield)
                         If j > 1 Then
                            sSubfieldQuery = sSubfieldQuery & "' or @code='"
                         End If
                         sSubfieldQuery = sSubfieldQuery & Mid(sSubfield, j, 1)
                       Next j
                       sSubfieldQuery = sSubfieldQuery & "']"
                    
                       Set oFieldList = aRecords(i).SelectNodes(sBibPrefix & "[@tag='880'][marc:subfield[@code='6' and starts-with(text(),'" & sField & "')]]")
                       For j = 0 To oFieldList.length - 1
                          If sRecord <> "" And Right(sRecord, 1) <> ChrW(166) Then
                            sRecord = sRecord & ChrW(166)
                          End If
                          Set oSubfieldList = oFieldList.Item(j).SelectNodes("marc:subfield" & sSubfieldQuery)
                          For k = 0 To oSubfieldList.length - 1
                            sRecord = sRecord & oSubfieldList.Item(k).XML
                          Next k
                       Next j
                    
                       oRegEx.Pattern = "<[^>]*>"
                       sRecord = oRegEx.Replace(sRecord, " ")
                       oRegEx.Pattern = "  *"
                       sRecord = oRegEx.Replace(sRecord, " ")
                      Else
                        sRecord = "Error in field/subfield name"
                      End If
                      oRegEx.Pattern = " \u00A6 "
                      sRecord = Trim(oRegEx.Replace(sRecord, ChrW(166)))
                      
                      If iSubStartPos > -1 And iSubLength > 0 Then
                         sRecordFiltered = ""
                         aResults = Split(sRecord, ChrW(166), -1, 0)
                         For j = 0 To UBound(aResults)
                            If sRecordFiltered <> "" Then
                                sRecordFiltered = sRecordFiltered & ChrW(166)
                            End If
                            sRecordFiltered = sRecordFiltered & Mid(aResults(j), iSubStartPos + 1, iSubLength)
                         Next j
                         sRecord = sRecordFiltered
                      End If
                      
                      If sResultFilter <> "" Then
                         sRecordFiltered = ""
                         aResults = Split(sRecord, ChrW(166), -1, 0)
                         For j = 0 To UBound(aResults)
                            If InStr(1, aResults(j), sResultFilter) > 0 Then
                                If sRecordFiltered <> "" Then
                                    sRecordFiltered = sRecordFiltered & ChrW(166)
                                End If
                                sRecordFiltered = sRecordFiltered & aResults(j)
                            End If
                         Next j
                         sRecord = sRecordFiltered
                      End If
                      ExtractField = ExtractField & sRecord
                  Case Else
                     ExtractField = ExtractField & "ERROR"
              End Select
              sRecord = ""
           Next h
           ExtractField = ExtractField & "|"
        Else
           ExtractField = ExtractField & "ERROR" & "|"
        End If
    Next i
    If Len(ExtractField) > 0 Then
        ExtractField = Left(ExtractField, Len(ExtractField) - 1)
        If Right(ExtractField, 1) = ChrW(166) Then
            ExtractField = Left(ExtractField, Len(ExtractField) - 1)
        End If
        ExtractField = Replace(ExtractField, Chr(10), "")
        ExtractField = Replace(ExtractField, Chr(13), "")
    Else
        If sResultType = "exists" Then
            ExtractField = "FALSE"
        Else
            ExtractField = "TRUE"
        End If
    End If
    oRegEx.Pattern = "&[^; ]+;"
    If oRegEx.Test(ExtractField) Then
        ExtractField = HtmlDecode(ExtractField)
        If oRegEx.Test(ExtractField) Then
            ExtractField = HtmlDecode(ExtractField)
        End If
    End If
    If LookupDialog.CatalogURLBox.Value = "source:borrowdirect" Then
        ExtractField = DecodeIPLCUnicode(ExtractField)
    End If
    If sResultType = "999$sp" Then
        ExtractField = CollapseIPLCHoldings(ExtractField)
    End If
    If LookupDialog.CatalogURLBox.Value = "source:worldcat" And sResultTypeAll = "948$ch#" Then
        sCodesDeDupe = ""
        iCodeCount = 0
        oRegEx.Pattern = "[0-9]* OTHER HOLDINGS"
        Set oMatch = oRegEx.Execute(ExtractField)
        If oMatch.Count > 0 Then
            ExtractField = CStr(CInt(Replace(oMatch(0), " OTHER HOLDINGS", "")) + 1)
        Else
            aCodesA = Split(ExtractField, "|")
            For i = 0 To UBound(aCodesA)
                aCodesB = Split(aCodesA(i), ChrW(166))
                For j = 0 To UBound(aCodesB)
                    If InStr(1, sCodesDeDupe, aCodesB(j)) = 0 Then
                        sCodesDeDupe = sCodesDeDupe & aCodesB(j) & "|"
                        iCodeCount = iCodeCount + 1
                    End If
                Next j
            Next i
            ExtractField = CStr(iCodeCount)
        End If
    End If
End Function

Function CollapseIPLCHoldings(sHoldings)
    sResult = ""
    aHoldingsA = Split(sHoldings, "|")
    For i = 0 To UBound(aHoldingsA)
        aHoldingsB = Split(aHoldingsA(i), ChrW(166))
        For j = 0 To UBound(aHoldingsB)
            sHCode = aHoldingsB(j)
            sHCode = Replace(sHCode, "ISIL:", "")
            iSpace = InStr(1, sHCode, " ")
            If iSpace > 0 Then
                sHCodeA = Left(sHCode, iSpace - 1)
                sHCodeB = Mid(sHCode, iSpace + 1)
                If InStr(1, sResult, sHCodeA) = 0 Then
                    If sResult <> "" Then
                        sResult = sResult & "|"
                    End If
                    sResult = sResult & sHCode
                ElseIf InStr(1, sResult, sHCode) = 0 Then
                    sResult = Replace(sResult, sHCodeA, sHCode)
                End If
            Else
                If InStr(1, sResult, sHCode) = 0 Then
                    If sResult <> "" Then
                        sResult = sResult & "|"
                    End If
                    sResult = sResult & sHCode
                End If
            End If
        Next j
    Next i
    CollapseIPLCHoldings = sResult
End Function

Function DecodeIPLCUnicode(sSource As String) As String
    oRegEx.Pattern = "\\u[0-9a-f]{4}"
    oRegEx.Global = True
    Set oMatch = oRegEx.Execute(sSource)
    For i = 0 To oMatch.Count - 1
        sDecoded = Mid(CStr(oMatch.Item(i)), 3)
        sDecoded = ChrW(CDec("&H" & sDecoded))
        sSource = Replace(sSource, oMatch.Item(i), sDecoded)
    Next i
    DecodeIPLCUnicode = sSource
End Function

Function NormalizeISBN(sQuery As String) As String
    With oRegEx
        .MultiLine = False
        .Global = True
        .IgnoreCase = True
    End With
    sQuery = Replace(sQuery, "-", "")
    oRegEx.Pattern = "[0-9]{13}"
    Set oMatch = oRegEx.Execute(sQuery)
    If oMatch.Count = 0 Then
        oRegEx.Pattern = "[0-9]{9}([0-9]|X)"
        Set oMatch = oRegEx.Execute(sQuery)
        If oMatch.Count = 0 Then
            NormalizeISBN = ""
        Else
            NormalizeISBN = oMatch.Item(0)
        End If
    Else
        NormalizeISBN = oMatch.Item(0)
    End If
    If LookupDialog.ValidateCheckBox.Value Then
        NormalizeISBN = GenerateCheckDigit(NormalizeISBN)
        NormalizeISBN = GetOtherISBN(NormalizeISBN)
    End If
End Function

Function GetOtherISBN(sISBN As String) As String
    sOtherISBN = ""
    If Len(sISBN) = 10 Then
        sOtherISBN = "978" & Left(sISBN, 9)
        iChecksum = 0
        For i = 1 To 12
            iMul = 1
            If i Mod 2 = 0 Then
                iMul = 3
            End If
            iChecksum = iChecksum + (iMul * CInt(Mid(sOtherISBN, i, 1)))
        Next i
        iChecksum = iChecksum Mod 10
        If iChecksum > 0 Then
            iChecksum = 10 - iChecksum
        End If
        sOtherISBN = sOtherISBN & iChecksum
    ElseIf Len(sISBN) = 13 Then
        sOtherISBN = Mid(sISBN, 4, 9)
        For i = 1 To 9
            iMul = 11 - i
            iChecksum = iChecksum + (iMul * CInt(Mid(sOtherISBN, i, 1)))
        Next i
        iChecksum = iChecksum Mod 11
        If iChecksum > 0 Then
            iChecksum = 11 - iChecksum
        End If
        If iChecksum = 10 Then
            sOtherISBN = sOtherISBN & "X"
        Else
            sOtherISBN = sOtherISBN & iChecksum
        End If
    End If
    If sOtherISBN <> "" Then
        GetOtherISBN = sISBN & "|" & sOtherISBN
    Else
        GetOtherISBN = sISBN
    End If
End Function

Function GenerateCheckDigit(sISXN As String) As String
    iChecksum = 0
    If Len(sISXN) <> 8 And Len(sISXN) <> 10 And Len(sISXN) <> 13 Then
        GenerateCheckDigit = "INVALID"
        Exit Function
    End If
    GenerateCheckDigit = Left(sISXN, Len(sISXN) - 1)
    If Len(sISXN) = 8 Then
        For i = 1 To 7
            iMul = 9 - i
            iChecksum = iChecksum + (iMul * CInt(Mid(sISXN, i, 1)))
        Next i
        iChecksum = iChecksum Mod 11
        If iChecksum > 0 Then
            iChecksum = 11 - iChecksum
        End If
        If iChecksum = 10 Then
            GenerateCheckDigit = GenerateCheckDigit & "X"
        Else
            GenerateCheckDigit = GenerateCheckDigit & iChecksum
        End If
    ElseIf Len(sISXN) = 10 Then
        For i = 1 To 9
            iMul = 11 - i
            iChecksum = iChecksum + (iMul * CInt(Mid(sISXN, i, 1)))
        Next i
        iChecksum = iChecksum Mod 11
        If iChecksum > 0 Then
            iChecksum = 11 - iChecksum
        End If
        If iChecksum = 10 Then
            GenerateCheckDigit = GenerateCheckDigit & "X"
        Else
            GenerateCheckDigit = GenerateCheckDigit & iChecksum
        End If
    ElseIf Len(sISXN) = 13 Then
        For i = 1 To 12
            iMul = 1
            If i Mod 2 = 0 Then
                iMul = 3
            End If
            iChecksum = iChecksum + (iMul * CInt(Mid(GenerateCheckDigit, i, 1)))
        Next i
        iChecksum = iChecksum Mod 10
        If iChecksum > 0 Then
            iChecksum = 10 - iChecksum
        End If
        GenerateCheckDigit = GenerateCheckDigit & iChecksum
    Else
        GenerateCheckDigit = "INVALID"
    End If
    If GenerateCheckDigit <> sISXN Then
        GenerateCheckDigit = "INVALID"
    End If
End Function

Function NormalizeOCLC(sQuery As String) As String
    With oRegEx
       .MultiLine = False
        .Global = True
        .IgnoreCase = True
    End With
    oRegEx.Pattern = "^[^0-9]*"
    sQuery = oRegEx.Replace(sQuery, "")
    oRegEx.Pattern = "[^0-9].*$"
    sQuery = oRegEx.Replace(sQuery, "")
    oRegEx.Pattern = "^0*"
    sQuery = oRegEx.Replace(sQuery, "")
    If sQuery = "" Then
        sQuery = "FALSE"
    End If
    NormalizeOCLC = sQuery
End Function

Function NormalizeISSN(sQuery As String) As String
    With oRegEx
        .MultiLine = False
        .Global = True
        .IgnoreCase = True
    End With
    sQuery = Replace(sQuery, "-", "")
    oRegEx.Pattern = "[0-9]{7}([0-9]|X)"
    Set oMatch = oRegEx.Execute(sQuery)
    If oMatch.Count = 0 Then
        NormalizeISSN = ""
    Else
        NormalizeISSN = Left(sQuery, 8)
    End If
    If LookupDialog.ValidateCheckBox.Value Then
        NormalizeISSN = GenerateCheckDigit(NormalizeISSN)
    End If
End Function

Public Function HtmlDecode(StringToDecode As Variant) As String
    oRegEx.Global = True
    oRegEx.Pattern = "&[^; ]+;"
    Set oMatch = oRegEx.Execute(StringToDecode)
    For i = 0 To oMatch.Count - 1
        sEntity = CStr(oMatch.Item(i))
        StringToDecode = Replace(StringToDecode, sEntity, LCase(sEntity))
    Next i
    Set oMSHTML = CreateObject("htmlfile")
    Set E = oMSHTML.createElement("T")
    E.innerHTML = StringToDecode
    HtmlDecode = E.innerText
End Function