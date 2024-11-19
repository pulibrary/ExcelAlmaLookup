Attribute VB_Name = "Catalog"
Global oRegEx As Object
Global oXMLHTTP As Object
Global oXMLDOM As Object
Global oConverter As Object
Global oRegistry As Object
Global oZConn As LongPtr

Global aExplainFields As Variant
Global bTerminateLoop As Boolean
Global bKeepTryingURL As Boolean
Global bIsoholdEnabled  As Boolean
Global bIsAlma As Boolean
Global sCatalogURL As String
Global sAuth As String

Global sFileName As String
Global sSheetName As String

Public Const HKEY_CURRENT_USER = &H80000001
Public Const sRegString = "Software\Excel Local Catalog Lookup"
Public Const sVersion = "v1.3.4"
Public Const sRepoURL = "https://github.com/pulibrary/ExcelAlmaLookup"
Public Const sBlacklightURL = "https://catalog.princeton.edu/catalog.json?q="
Public Const sLCCatURL = "http://lx2.loc.gov:210/LCDB"
Public Const sIPLCReshareURL = "https://borrowdirect.reshare.indexdata.com/api/v1/search?type=AllFields&field%5B%5D=fullRecord&lookfor="

Public Const iMaximumRecords = 25

Public Const sWCZhost = "zcat.oclc.org"
Public Const sWCZport = 210
Public Const sWCZDB = "OLUCWorldCat"

Public Const sPluginDir = "Excel Local Catalog Lookup"
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
    On Error GoTo ErrHandler
    Set oRegEx = CreateObject("vbscript.regexp")
    With oRegEx
        .MultiLine = False
        .Global = True
        .IgnoreCase = True
    End With
    Set oRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    Set oXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    Set oXMLDOM = CreateObject("MSXML2.DomDocument")
    oXMLDOM.SetProperty "SelectionLanguage", "XPath"
    Set oConverter = CreateObject("ADODB.Stream")
    
    SetDefaultDllDirectories (&H1000&)
    sDllPath = StrConv(Environ("AppData") & "\" & sPluginDir & "\" & sDllVersion & "\", vbUnicode)
    AddDllDirectory (sDllPath)
    LoadLibrary (sYAZdll)
    Exit Sub
ErrHandler:
    MsgBox ("There was an error initializing the plugin.  Please try again.")
    End
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
    
    PopulateCombos
    RedrawButtons
    
    sFileName = ActiveWorkbook.Name
    sSheetName = ActiveSheet.Name
        
    LookupDialog.ResultColumnSpinner.Value = FindLastColumn() + 1
    sSourceRange = Selection.Address
    LookupDialog.LookupRange.Value = sSourceRange
    
    
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

Function GetRegistryURLs()
    On Error Resume Next
    If oRegistry Is Nothing Then
        Initialize
    End If
    oRegistry.GetStringValue HKEY_CURRENT_USER, "Software\Excel Local Catalog Lookup", "CatalogURL", sValue
    If IsNull(sValue) Then
        sValue = ""
    End If
    GetRegistryURLs = sValue
    If Err.Number <> 0 Then
        oRegistry.SetStringValue HKEY_CURRENT_USER, "Software\Excel Local Catalog Lookup", "CatalogURL", ""
        GetRegistryURLs = ""
    End If
End Function

Sub SetRegistryURLsFromCombo()
    If oRegistry Is Nothing Then
        Initialize
    End If
    iSelected = LookupDialog.CatalogURLBox.ListIndex
    sCatalogURLs = LookupDialog.CatalogURLBox.Value
    addToCombo = True
    For i = 0 To LookupDialog.CatalogURLBox.ListCount - 1
        If i <> iSelected Then
            sCatalogURLs = sCatalogURLs & "|" & LookupDialog.CatalogURLBox.List(i)
        End If
        If LookupDialog.CatalogURLBox.List(i) = LookupDialog.CatalogURLBox.Value Then
            addToCombo = False
        End If
    Next i
    If addToCombo Then
        LookupDialog.CatalogURLBox.AddItem LookupDialog.CatalogURLBox.Value
    End If
    oRegistry.SetStringValue HKEY_CURRENT_USER, "Software\Excel Local Catalog Lookup", "CatalogURL", sCatalogURLs
    
    sCatalogAuths = GetRegistryAuths
    aCatalogAuths = Split(sCatalogAuths, "|", -1, 0)
    sCatalogAuths = ""
    For i = 0 To UBound(aCatalogAuths)
        aURLAuth = Split(aCatalogAuths(i), ChrW(166), -1, 0)
        If InStr(1, sCatalogURLs, aURLAuth(0)) Then
            If sCatalogAuths <> "" Then
                sCatalogAuths = sCatalogAuths & "|"
            End If
            sCatalogAuths = sCatalogAuths & aCatalogAuths(i)
        End If
    Next i
    oRegistry.SetStringValue HKEY_CURRENT_USER, "Software\Excel Local Catalog Lookup", "CatalogAuth", sCatalogAuths
End Sub

Function GetRegistryAuths()
    On Error Resume Next
    If oRegistry Is Nothing Then
        Initialize
    End If
    oRegistry.GetStringValue HKEY_CURRENT_USER, "Software\Excel Local Catalog Lookup", "CatalogAuth", sValue
    If IsNull(sValue) Then
        sValue = ""
    End If
    GetRegistryAuths = sValue
    If Err.Number <> 0 Then
        oRegistry.SetStringValue HKEY_CURRENT_USER, "Software\Excel Local Catalog Lookup", "CatalogAuth", ""
        GetRegistryAuths = ""
    End If
End Function

Sub SaveCatalogAuthToRegistry()
    If oRegistry Is Nothing Then
        Initialize
    End If
    sCatalogAuths = GetRegsistryAuths
    sCatalogURL = LookupDialog.CatalogURLBox.Text
    aCatalogAuths = Split(sCatalogAuths, "|", -1, 0)
    sCatalogAuths = ""
    bFound = False
    For i = 0 To UBound(aCatalogAuths)
        If i > 0 Then
            sCatalogAuths = sCatalogAuths & "|"
        End If
        aURLAuth = Split(aCatalogAuths(i), ChrW(166), -1, 0)
        If aURLAuth(0) = LookupDialog.CatalogURLBox.Text Then
            bFound = True
            sCatalogAuths = sCatalogAuths & LookupDialog.CatalogURLBox.Text & ChrW(166) & sAuth
        Else
            sCatalogAuths = sCatalogAuths & aCatalogAuths(i)
        End If
    Next i
    If Not bFound Then
        If sCatalogAuths <> "" Then
            sCatalogAuths = sCatalogAuths & "|"
        End If
        sCatalogAuths = sCatalogAuths & sCatalogURL & ChrW(166) & sAuth
    End If
        
    
    oRegistry.SetStringValue HKEY_CURRENT_USER, "Software\Excel Local Catalog Lookup", "CatalogAuth", sCatalogAuths
End Sub

Sub ClearRegistryAuth()
    sCatalogURL = LookupDialog.CatalogURLBox.Text
    sCatalogAuths = GetRegistryAuths
    aCatalogAuths = Split(sCatalogAuths, "|", -1, 0)
    sCatalogAuths = ""
    For i = 0 To UBound(aCatalogAuths)
        aURLAuth = Split(aCatalogAuths(i), ChrW(166), -1, 0)
        If sCatalogURL <> aURLAuth(0) Then
            If sCatalogAuths <> "" Then
                sCatalogAuths = sCatalogAuths & "|"
            End If
            sCatalogAuths = sCatalogAuths & aCatalogAuths(i)
        End If
    Next i
    oRegistry.SetStringValue HKEY_CURRENT_USER, "Software\Excel Local Catalog Lookup", "CatalogAuth", sCatalogAuths
End Sub

Function GetFieldSets()
    On Error Resume Next
    If oRegistry Is Nothing Then
        Initialize
    End If
    oRegistry.GetStringValue HKEY_CURRENT_USER, "Software\Excel Local Catalog Lookup", "FieldSets", sValue
    If IsNull(sValue) Then
        sValue = ""
    End If
    GetFieldSets = sValue
     If Err.Number <> 0 Then
        oRegistry.SetStringValue HKEY_CURRENT_USER, "Software\Excel Local Catalog Lookup", "FieldSets", ""
        GetFieldSets = ""
    End If
End Function

Function SetFieldSets(sSetString As String)
    If oRegistry Is Nothing Then
        Initialize
    End If
    oRegistry.SetStringValue HKEY_CURRENT_USER, "Software\Excel Local Catalog Lookup", "FieldSets", sSetString
End Function

Sub PopulateCombos()
    Dim sCatalogURL As String
    On Error Resume Next
    sCatalogURLs = GetRegistryURLs()
    sCatalogAuths = GetRegistryAuths()
    If Err.Number = 0 Then
        LookupDialog.CatalogURLBox.Clear
        aCatalogURLs = Split(sCatalogURLs, "|", -1, 0)
        For i = 0 To UBound(aCatalogURLs)
            LookupDialog.CatalogURLBox.AddItem aCatalogURLs(i)
        Next i
        LookupDialog.CatalogURLBox.ListIndex = 0
        
        aCatalogAuths = Split(sCatalogAuths, "|", -1, 0)
        For i = 0 To UBound(aCatalogAuths)
            aURLAuth = Split(aCatalogAuths(i), ChrW(166), -1, 0)
            If aURLAuth(0) = LookupDialog.CatalogURLBox.Text Then
                sAuth = aURLAuth(1)
                Exit For
            End If
        Next i
    End If

    sFieldSets = GetFieldSets()
    If Err.Number = 0 Then
        LookupDialog.FieldSetList.Clear
        aFieldSets = Split(sFieldSets, "|", -1, 0)
        For i = 0 To UBound(aFieldSets)
            aFields = Split(aFieldSets(i), ChrW(166), -1, 0)
            LookupDialog.FieldSetList.AddItem aFields(0)
        Next i
    End If
    
    PopulateSourceDependentOptions
    
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
            End If
            LookupDialog.SearchFieldCombo.AddItem "OCLC No."

        Else
            LookupDialog.SearchFieldCombo.Enabled = False
        End If
        
        LookupDialog.AdditionalFieldsButton.Enabled = False
    Else
        LookupDialog.SearchFieldCombo.Style = 0 'fmStyleDropDownCombo
        LookupDialog.SearchFieldCombo.Enabled = True
        LookupDialog.AdditionalFieldsButton.Enabled = True
        LookupDialog.SearchFieldCombo.AddItem "Keywords"
        LookupDialog.SearchFieldCombo.AddItem "Call No."
        LookupDialog.SearchFieldCombo.AddItem "Title"
        LookupDialog.SearchFieldCombo.AddItem "ISBN"
        LookupDialog.SearchFieldCombo.AddItem "ISSN"
        LookupDialog.SearchFieldCombo.AddItem "MMS ID"
        LookupDialog.SearchFieldCombo.AddItem "Barcode"
        LookupDialog.SearchFieldCombo.ListIndex = 0
    End If
    LookupDialog.SearchFieldCombo.ListIndex = 0
    LookupDialog.ResultTypeCombo.ListIndex = 0

End Sub

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

Function ConstructURL(sBaseURL As String, sQuery1 As String, sSearchType As String) As String
    sURL = sBaseURL & "?operation=searchRetrieve&version=1.2&maximumRecords=" & iMaximumRecords & "&query="
    sQuery1 = Replace(sQuery1, "http://", "")
    sQuery = EncodeURI(sQuery1)
    sIndex = ""
    Select Case sSearchType

    Case "Keywords"
        sIndex = "alma.all_for_ui"
    Case "Call No."
        sIndex = "alma.PermanentCallNumber"
    Case "MMS ID"
        sIndex = "rec.id"
    Case "Title"
        sIndex = "alma.title"
    Case "ISBN"
        sQuery = NormalizeISBN(sQuery1)
        If sQuery = "" Then
            sQuery = "FALSE"
        End If
        sIndex = "alma.isbn"
    Case "ISSN"
        sQuery = NormalizeISSN(sQuery1)
        If sQuery = "" Then
            sQuery = "FALSE"
        End If
        sIndex = "alma.issn"
    Case "Barcode"
        sIndex = "alma.barcode"
        sQuery = Replace(sQuery1, " ", "")
    Case Else
        sIndex = sSearchType
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
    sQuery = Replace(sQuery, "|", "%22+OR+" & sIndex & "+%3D+%22")
    sQuery = Replace(sQuery, "%7C", "%22+OR+" & sIndex & "+%3D+%22")
    sURL = sURL & sIndex
    If sIndex = "alma.PermanentCallNumber" Then
        sURL = sURL & "+all+"
    Else
        sURL = sURL & "+%3D+"
    End If
    sURL = sURL & "%22" + sQuery + "%22"
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
        aFieldMap(i, 0) = sLabel
        aFieldMap(i, 1) = sIndexSet & "." & sIndexCode
    Next i
    
    For i = 0 To UBound(aFieldMap)
        For j = i + 1 To UBound(aFieldMap)
            SearchI = Replace(UCase(aFieldMap(i, 0)), "(", "")
            SearchJ = Replace(UCase(aFieldMap(j, 0)), "(", "")
            If UCase(SearchI > SearchJ) Then
                t1 = aFieldMap(j, 0)
                t2 = aFieldMap(j, 1)
                aFieldMap(j, 0) = aFieldMap(i, 0)
                aFieldMap(j, 1) = aFieldMap(i, 1)
                aFieldMap(i, 0) = t1
                aFieldMap(i, 1) = t2
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
    sQuery = Replace(sQuery, "ï»¿", "") 'BOM
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
    aSearchKeys = Split(sQuery, "|")
    For i = 0 To UBound(aSearchKeys)
        If sCQLQuery <> "" Then
            sCQLQuery = "@or " & sCQLQuery
        End If
        sCQLQuery = sCQLQuery & "@attr 4=1 @attr 1=" & sSearchIndex & " """ & aSearchKeys(i) & """"
    Next i

    zrs = ZOOM_connection_search_pqf(oZConn, sCQLQuery)
    ZOOM_resultset_option_set zrs, "count", iMaximumRecords
    zcount = ZOOM_resultset_size(zrs)
    If zcount > 0 Then
        sAllRecords = "<searchRetrieveResponse xmlns=""http://www.loc.gov/zing/srw/""><records>"
        If zcount > iMaximumRecords Then
            zcount = iMaximumRecords
        End If
        For i = 0 To zcount - 1
            zrec = ZOOM_resultset_record(zrs, i)
            Dim zptr As LongPtr
            Dim zsize As Long
            zptr = ZOOM_record_get(zrec, "xml;charset=marc8,utf8", zsize)
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
    'ZOOM_resultset_destroy (zrs)
End Function

Function Lookup(sQuery1 As String, sCatalogURL As String) As String
    If oXMLHTTP Is Nothing Then
        Initialize
    End If
    
    Dim sSearchType As String
    sSearchType = CStr(LookupDialog.SearchFieldCombo.Value)
    Dim sFormat As String
    sURL = ""
    
    If LookupDialog.ValidateCheckBox.Value Then
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
        sURL = ConstructURL(sCatalogURL, sQuery1, sSearchType)
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
        Set oRegEx = CreateObject("vbscript.regexp")
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
                    Case "020"
                        oRegEx.Pattern = "\[([^\]]*)\],""label"":""Isbn S"""
                        Set oISBNs = oRegEx.Execute(sCurrentRecord)
                        If oISBNs.Count > 0 Then
                            sISBNs = oISBNs(0).Submatches(0)
                            sISBNs = Replace(sISBNs, """,""", Chr(166))
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
                            sOCLCs = Replace(sOCLCs, """,""", Chr(166))
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
                                sCGD = sCGD & Chr(166)
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
End Function

Function CollapseIPLCHoldings(sHoldings)
    sResult = ""
    aHoldingsA = Split(sHoldings, "|")
    For i = 0 To UBound(aHoldingsA)
        aHoldingsB = Split(aHoldingsA(i), Chr(166))
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
    Set oRegEx = CreateObject("vbscript.regexp")
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
    Set oRegEx = CreateObject("vbscript.regexp")
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
    Set oRegEx = CreateObject("vbscript.regexp")
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
    Set oRegEx = CreateObject("vbscript.regexp")
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
    Set oRegEx = CreateObject("vbscript.regexp")
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