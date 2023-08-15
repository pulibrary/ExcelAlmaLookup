Global oRegEx As Object
Global oXMLHTTP As Object
Global oXMLDOM As Object
Global oRegistry As Object
Global sRegString As String
Global aExplainFields As Variant
Global bTerminateLoop As Boolean

'Initialize global objects
Private Sub Initialize()
    On Error GoTo ErrHandler
    Set oRegEx = CreateObject("vbscript.regexp")
    With oRegEx
        .MultiLine = False
        .Global = True
        .IgnoreCase = True
    End With
    
    Set oRegistry = CreateObject("WScript.Shell")
    sRegString = "HKCU\Software\Excel Local Catalog Lookup\"
    Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP")
    Set oXMLDOM = CreateObject("MSXML2.DomDocument")
    oXMLDOM.SetProperty "SelectionLanguage", "XPath"
    Exit Sub
ErrHandler:
    MsgBox ("There was an error connecting to the catalog.  Please try again.")
    End
End Sub

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
    LookupDialog.ResultColumnSpinner.Value = FindLastColumn() + 1
    LookupDialog.LookupRange.Value = Selection.Address
    LookupDialog.Show
End Sub

Function GetRegistryURLs()
    On Error Resume Next
    If oRegistry Is Nothing Then
        Initialize
    End If
    GetRegistryURLs = oRegistry.RegRead(sRegString & "CatalogURL")
    If Err.Number <> 0 Then
        oRegistry.RegWrite sRegString & "CatalogURL", "", "REG_SZ"
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
    oRegistry.RegWrite sRegString & "CatalogURL", sCatalogURLs, "REG_SZ"
End Sub

Function GetFieldSets()
    On Error Resume Next
    If oRegistry Is Nothing Then
        Initialize
    End If
    GetFieldSets = oRegistry.RegRead(sRegString & "FieldSets")
     If Err.Number <> 0 Then
        oRegistry.RegWrite sRegString & "FieldSets", "", "REG_SZ"
        GetFieldSets = ""
    End If
End Function

Function SetFieldSets(sSetString As String)
    If oRegistry Is Nothing Then
        Initialize
    End If
    oRegistry.RegWrite sRegString & "FieldSets", sSetString, "REG_SZ"
End Function

Sub PopulateCombos()
    Dim sCatalogURL As String
    On Error Resume Next
    sCatalogURLs = GetRegistryURLs()
    If Err.Number = 0 Then
        LookupDialog.CatalogURLBox.Clear
        aCatalogURLs = Split(sCatalogURLs, "|")
        For i = 0 To UBound(aCatalogURLs)
            LookupDialog.CatalogURLBox.AddItem aCatalogURLs(i)
        Next i
        LookupDialog.CatalogURLBox.ListIndex = 0
        'LookupDialog.CatalogURLBox.Text = sCatalogURL
    End If

    sFieldSets = GetFieldSets()
    If Err.Number = 0 Then
        LookupDialog.FieldSetList.Clear
        aFieldSets = Split(sFieldSets, "|")
        For i = 0 To UBound(aFieldSets)
            aFields = Split(aFieldSets(i), "¦")
            LookupDialog.FieldSetList.AddItem aFields(0)
        Next i
    End If

    LookupDialog.SearchFieldCombo.Clear
    LookupDialog.ResultTypeCombo.Clear
            
    LookupDialog.SearchFieldCombo.AddItem "Keywords"
    LookupDialog.SearchFieldCombo.AddItem "Call No."
    LookupDialog.SearchFieldCombo.AddItem "Title"
    LookupDialog.SearchFieldCombo.AddItem "ISBN"
    LookupDialog.SearchFieldCombo.AddItem "ISSN"
    LookupDialog.SearchFieldCombo.AddItem "MMS ID"
    
    LookupDialog.ResultTypeCombo.AddItem "True/False"
    LookupDialog.ResultTypeCombo.AddItem "MMS ID"
    LookupDialog.ResultTypeCombo.AddItem "ISBN"
    LookupDialog.ResultTypeCombo.AddItem "Title"
    LookupDialog.ResultTypeCombo.AddItem "*Call No."
    LookupDialog.ResultTypeCombo.AddItem "*Location/DB Name"
    LookupDialog.ResultTypeCombo.AddItem "Language code"
    LookupDialog.ResultTypeCombo.AddItem "*Coverage"
    LookupDialog.ResultTypeCombo.AddItem "Leader"

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

'Determined the rightmost column containing data
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
    sUrl = sBaseURL & "?operation=searchRetrieve&version=1.2&query="
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
        sIndex = "alma.isbn"
    Case "ISSN"
        sQuery = NormalizeISSN(sQuery1)
        sIndex = "alma.issn"
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
    sUrl = sUrl & sIndex
    If sIndex = "alma.PermanentCallNumber" Then
        sUrl = sUrl & "+all+"
    Else
        sUrl = sUrl & "+%3D+"
    End If
    sUrl = sUrl & "%22" + sQuery + "%22"
    If Not LookupDialog.IncludeSuppressed Then
        sUrl = sUrl & "+AND+alma.mms_tagSuppressed=false"
    End If
    ConstructURL = sUrl

End Function


Function GetAllFields()
    If oXMLHTTP Is Nothing Then
        Initialize
    End If
    Dim sCatalogURL As String
    sCatalogURL = CStr(LookupDialog.CatalogURLBox.Text)
    invalidURL = False
    If Left(sCatalogURL, 4) <> "http" Then
        invalidURL = True
    End If
    If Not invalidURL Then
        sExplainURL = sCatalogURL & "?version=1.2&operation=explain"
        With oXMLHTTP
            .Open "GET", sExplainURL, True
            .send
            Do While .readyState <> 4
                DoEvents
            Loop
            sResponse = .responseText
            If .Status <> 200 Or InStr(sResponse, "explainResponse") = 0 Then
                invalidURL = True
            End If
        End With
    End If
    If invalidURL Then
        MsgBox ("Cannot access catalog.  Please confirm the Alma URL is correct")
        GetAllFields = Null
        Exit Function
    End If
    oXMLDOM.SetProperty "SelectionNamespaces", "xmlns:xr='http://www.loc.gov/zing/srw/' " & _
        "xmlns:xpl='http://explain.z3950.org/dtd/2.0/' " & _
        "xmlns:ns='http://explain.z3950.org/dtd/2.1/'"
    oXMLDOM.LoadXML (sResponse)
    sFields = ""
    Set aFields = oXMLDOM.SelectNodes("xr:explainResponse/xr:record/xr:recordData/" & _
        "xpl:explain/xpl:indexInfo/xpl:index")
    Dim aFieldMap() As Variant
    ReDim aFieldMap(aFields.Length - 1, 2)
    For i = 0 To aFields.Length - 1
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

Function Lookup(sQuery1 As String, sCatalogURL As String) As String
    If oXMLHTTP Is Nothing Then
        Initialize
    End If
    
    Dim sSearchType As String
    sSearchType = CStr(LookupDialog.SearchFieldCombo.Value)
    Dim sFormat As String
      
    
    sUrl = ConstructURL(sCatalogURL, sQuery1, sSearchType)
    sResponse = ""

    With oXMLHTTP
        .Open "GET", sUrl, True
        .send
    
        Do While .readyState <> 4
            DoEvents
        Loop
        sResponse = .responseText
    End With
    Lookup = sResponse
End Function

Function ExtractField(sResultTypeAll As String, sResultXML As String) As String
    aResultFields = Split(sResultTypeAll, "|")
    iResultTypes = UBound(aResultFields)
    oXMLDOM.SetProperty "SelectionNamespaces", "xmlns:sr='http://www.loc.gov/zing/srw/' " & _
        "xmlns:marc='http://www.loc.gov/MARC21/slim'"
    oXMLDOM.LoadXML (sResultXML)
    Set aRecords = oXMLDOM.SelectNodes("sr:searchRetrieveResponse/sr:records/sr:record/sr:recordData/marc:record")

    iRecords = aRecords.Length

    If iRecords = 0 Then
        ExtractField = "FALSE"
        Exit Function
    ElseIf sResultType = "exists" Then
        ExtractField = "TRUE"
        Exit Function
    End If
    
    ExtractField = ""
       
    'Iterate through results, compile result string
    For i = 0 To iRecords - 1
        If oXMLDOM.parseError.ErrorCode = 0 Then
           sRecord = ""
           For h = 0 To UBound(aResultFields)
              sResultType = aResultFields(h)
              sResultFilter = ""
              iFilterPos = InStr(1, sResultType, "#")
              If iFilterPos > 0 Then
                sResultFilter = Mid(sResultType, iFilterPos + 1)
                sResultType = Left(sResultType, iFilterPos - 1)
              End If
              sBibPrefix = "marc:datafield"
               If sResultType = "000" Then
                  sBibPrefix = "marc:leader"
               ElseIf Left(sResultType, 2) Like "00" Then
                  sBibPrefix = "marc:controlfield"
               End If
               Select Case sResultType
                  Case "exists"
                     ExtractField = "TRUE "
                  Case "000" To "999z", "AVA" To "AVAz", "AVD" To "AVDz", "AVE" To "AVEz"
                     Dim oFieldList As IXMLDOMNodeList
                     If sResultType = "000" Then
                       Set oFieldList = aRecords(i).SelectNodes(sBibPrefix)
                       sRecord = oFieldList.Item(0).XML
                       oRegEx.Pattern = "<[^>]*>"
                       sRecord = oRegEx.Replace(sRecord, "")
                     ElseIf sResultType Like "###" Then
                       Set oFieldList = aRecords(i).SelectNodes(sBibPrefix & "[@tag='" & sResultType & "']")
                       For j = 0 To oFieldList.Length - 1
                         If sRecord <> "" Then
                            sRecord = sRecord & Chr(166)
                         End If
                         sRecord = sRecord & oFieldList.Item(j).XML
                       Next j
                    
                       oRegEx.Pattern = "<subfield code=.6.>[^<]*</subfield>"
                       sRecord = oRegEx.Replace(sRecord, "")
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
                    ElseIf sResultType Like "00#(*,*)" Then
                       sField = Left(sResultType, 3)
                       iSubStart = Mid(sResultType, 5, InStr(1, sResultType, ",") - 5)
                       iSubLen = Mid(sResultType, InStr(1, sResultType, ",") + 1)
                       iSubLen = Left(iSubLen, Len(iSubLen) - 1)
                       Set oFieldList = aRecords(i).SelectNodes(sBibPrefix & "[@tag='" & sField & "']")
                       If Not oFieldList Is Nothing Then
                         sRecord = oFieldList.Item(0).XML
                         oRegEx.Pattern = "<[^>]*>"
                         sRecord = oRegEx.Replace(sRecord, " ")
                         oRegEx.Pattern = "^\s+"
                         sRecord = oRegEx.Replace(sRecord, "")
                         sRecord = Mid(sRecord, iSubStart + 1, iSubLen)
                       End If
                     ElseIf sResultType Like "###-880" Then
                        sMainField = Left(sResultType, 3)
                        Set oFieldList = aRecords(i).SelectNodes(sBibPrefix & "[@tag='880'][marc:subfield[@code='6' and starts-with(text(),'" & sMainField & "')]]")
                        For j = 0 To oFieldList.Length - 1
                          If sRecord <> "" Then
                            sRecord = sRecord & Chr(166)
                          End If
                          sRecord = sRecord & oFieldList.Item(j).XML
                        Next j
                        oRegEx.Pattern = "<subfield code=.6.>[^<]*</subfield>"
                        sRecord = oRegEx.Replace(sRecord, "")
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
                        For j = 0 To oFieldList.Length - 1
                           If sRecord <> "" And Right(sRecord, 1) <> Chr(166) Then
                             sRecord = sRecord & Chr(166)
                           End If
                           Set oSubfieldList = oFieldList.Item(j).SelectNodes("marc:subfield" & sSubfieldQuery)
                           For k = 0 To oSubfieldList.Length - 1
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
                       For j = 0 To oFieldList.Length - 1
                          If sRecord <> "" And Right(sRecord, 1) <> Chr(166) Then
                            sRecord = sRecord & Chr(166)
                          End If
                          Set oSubfieldList = oFieldList.Item(j).SelectNodes("marc:subfield" & sSubfieldQuery)
                          For k = 0 To oSubfieldList.Length - 1
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
                      sRecord = Trim(oRegEx.Replace(sRecord, Chr(166)))
                      If sResultFilter <> "" Then
                         sRecordFiltered = ""
                         aResults = Split(sRecord, Chr(166))
                         For j = 0 To UBound(aResults)
                            If InStr(1, aResults(j), sResultFilter) > 0 Then
                                If sRecordFiltered <> "" Then
                                    sRecordFiltered = sRecordFiltered & Chr(166)
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
        'Remove spaces around broken-bar delimiter (multiple hits in the same record)
    Next i
    If Len(ExtractField) > 0 Then
        ExtractField = Left(ExtractField, Len(ExtractField) - 1)
        ExtractField = Replace(ExtractField, Chr(10), "")
        ExtractField = Replace(ExtractField, Chr(13), "")
    Else
        If sResultType = "exists" Then
            ExtractField = "FALSE"
        Else
            ExtractField = "TRUE"
        End If
    End If
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
            NormalizeISBN = Left(sQuery, 10)
        End If
    Else
        NormalizeISBN = Left(sQuery, 13)
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

Function GenerateCheckDigit(sISXN As String)
    iChecksum = 0
    If Len(sISXN) <> 8 And Len(sISXN) <> 10 And Len(sISXN) <> 13 Then
        GenerateCheckDigit = sISXN
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
        GenerateCheckDigit = sISXN
    End If
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