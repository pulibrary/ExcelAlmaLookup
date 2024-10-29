Attribute VB_Name = "LookupDialog"
Attribute VB_Base = "0{B0A3DCBB-C34D-420F-BEBD-E00851586544}{CF25E99D-C97B-45FA-A9AC-A1803470DEB7}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub AdditionalFieldsButton_Click()
    Catalog.aExplainFields = Catalog.GetAllFields()
    AdditionalFieldsDialog.FilterBox.Value = ""
    If Not IsNull(aExplainFields) Then
        AdditionalFieldsDialog.SRUFields.List = aExplainFields
        AdditionalFieldsDialog.Show
    End If
End Sub

Private Sub AddResultButton_Click()
    With LookupDialog
        If ResultTypeCombo.Value <> "" Then
            If ResultTypeList.ListIndex > -1 Then
                ResultTypeList.AddItem ResultTypeCombo.Value, ResultTypeList.ListIndex
            Else
                ResultTypeList.AddItem ResultTypeCombo.Value
            End If
        End If
    End With
End Sub

Private Sub AddURLButton_Click()
    If Catalog.bIsAlma Then
        aFieldMap = Catalog.GetAllFields()
        If Not IsNull(aFieldMap) Then
            Catalog.SetRegistryURLsFromCombo
        End If
    Else
        Catalog.SetRegistryURLsFromCombo
    End If
End Sub

Private Sub CancelButton_Click()
    LookupDialog.Hide
    End
End Sub

Private Sub CatalogURLBox_Change()
    Catalog.sAuth = ""
    sCatalogAuths = GetRegistryAuths()
    aCatalogAuths = Split(sCatalogAuths, "|", -1, 0)
    For i = 0 To UBound(aCatalogAuths)
        aURLAuth = Split(aCatalogAuths(i), ChrW(166), -1, 0)
        If aURLAuth(0) = LookupDialog.CatalogURLBox.Text Then
            sAuth = aURLAuth(1)
            Exit For
        End If
    Next i
    Catalog.bIsAlma = True
    If InStr(1, LookupDialog.CatalogURLBox, "source:") = 1 Then
        Catalog.bIsAlma = False
    End If
    Catalog.PopulateSourceDependentOptions
End Sub

Private Sub ClearCredentialsButton_Click()
    Catalog.ClearRegistryAuth
    Catalog.sAuth = ""
End Sub

Private Sub DeleteSetButton_Click()
    If LookupDialog.FieldSetList.ListIndex < 0 Then
        MsgBox ("Please select a set name")
        Exit Sub
    End If
    sSelectedSet = LookupDialog.FieldSetList.Value
    sSelectedIndex = LookupDialog.FieldSetList.ListIndex
    LookupDialog.FieldSetList.RemoveItem sSelectedIndex
    sFieldSets = Catalog.GetFieldSets()
    aFieldSets = Split(sFieldSets, "|", -1, 0)
    sNewSets = ""
    For i = 0 To UBound(aFieldSets)
        If InStr(1, aFieldSets(i), sSelectedSet + ChrW(166)) = 0 Then
            If i > 0 Then
                sNewSets = sNewSets & "|"
            End If
            sNewSets = sNewSets & aFieldSets(i)
        End If
    Next i
    Catalog.SetFieldSets (sNewSets)
    LookupDialog.FieldSetList.ListIndex = -1
    Catalog.RedrawButtons
End Sub

Private Sub FieldSetList_Change()
    RedrawButtons
End Sub

Private Sub FieldSetList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If LookupDialog.FieldSetList.ListIndex > -1 Then
        sSelectedSet = LookupDialog.FieldSetList.Value
        bSuccess = LoadSet(CStr(sSelectedSet))
    End If
End Sub

Private Function LoadSet(sSetName As String) As Boolean
    LookupDialog.ResultTypeList.Clear
    sFieldSets = Catalog.GetFieldSets()
    aFieldSets = Split(sFieldSets, "|", -1, 0)
    For i = 0 To UBound(aFieldSets)
        If InStr(1, aFieldSets(i), sSetName & ChrW(166)) > 0 Then
            aFields = Split(aFieldSets(i), ChrW(166), -1, 0)
            For j = 1 To UBound(aFields)
                LookupDialog.ResultTypeList.AddItem aFields(j)
            Next j
        End If
    Next i
End Function

Private Sub HelpButton_Click()
    ThisWorkbook.FollowHyperlink Catalog.sRepoURL & "#readme"
End Sub

Private Sub IgnoreHeaderCheckbox_Click()
    If LookupDialog.IgnoreHeaderCheckbox.Value = True Then
        LookupDialog.GenerateHeaderCheckBox.Enabled = True
    Else
        LookupDialog.GenerateHeaderCheckBox.Enabled = False
    End If
End Sub

Private Sub LoadSetButton_Click()
    If LookupDialog.FieldSetList.ListIndex < 0 Then
        MsgBox ("Please select a set name")
        Exit Sub
    End If
    sSelectedSet = LookupDialog.FieldSetList.Value
    bSuccess = LoadSet(CStr(sSelectedSet))
    
End Sub

Private Sub MoveDownButton_Click()
    With LookupDialog.ResultTypeList
    Index = .ListIndex
    If Index < .ListCount - 1 Then
        t = .List(Index)
        .List(Index) = .List(Index + 1)
        .List(Index + 1) = t
        .Selected(Index + 1) = True
    End If
    End With
End Sub

Private Sub MoveUpButton_Click()
    With LookupDialog.ResultTypeList
    Index = .ListIndex
    If Index > 0 Then
        t = .List(Index)
        .List(Index) = .List(Index - 1)
        .List(Index - 1) = t
        .Selected(Index - 1) = True
    End If
    End With
End Sub

Private Sub NewSetButton_Click()
    sNewName = InputBox("Enter the Name of the New Set", "New Set")
    If sNewName = "" Then
        Exit Sub
    End If
    If InStr(1, sNewName, "|") > 0 Or InStr(1, sNewName, ChrW(166)) Then
        MsgBox ("Set name cannot contain vertical bar characters")
        Exit Sub
    End If
    bSuccess = SaveSet(CStr(sNewName))
    If bSuccess Then
        LookupDialog.FieldSetList.AddItem sNewName
    End If
    LookupDialog.FieldSetList.ListIndex = LookupDialog.FieldSetList.ListCount - 1
End Sub

Private Sub OKButton_Click()
    If Catalog.bIsAlma Then
        aFieldMap = Catalog.GetAllFields()
        If IsNull(aFieldMap) Then
            Exit Sub
        End If
    End If
    Dim sCatalogURL As String
    sCatalogURL = CStr(LookupDialog.CatalogURLBox.Text)
    If sCatalogURL = "source:worldcat" Then
        bSuccess = Catalog.Z3950Connect(sCatalogURL)
        If Not bSuccess Then
            Exit Sub
        End If
    End If
    Catalog.SetRegistryURLsFromCombo
    iResultColumn = LookupDialog.ResultColumnSpinner.Value
    If LookupDialog.ResultTypeList.ListCount = 0 Then
        AddResultButton_Click
    End If
    
    'Disable ISO Holdings if result types do not require them
    If Catalog.bIsoholdEnabled = True Then
        Catalog.bIsoholdEnabled = False
        For i = 0 To LookupDialog.ResultTypeList.ListCount - 1
            Dim sResType As String
            sResType = LookupDialog.ResultTypeList.List(i)
            If Left(sResType, 2) = "**" Then
                Catalog.bIsoholdEnabled = True
                Exit For
            End If
        Next i
    End If
    'Validate selected range, truncate to part containing actual data
    
    Set oSourceRange = Workbooks(Catalog.sFileName).Worksheets(Catalog.sSheetName).Range(LookupRange.Value)
    If IsObject(oSourceRange) Then
        LookupDialog.Hide
        With oSourceRange
            iRowCount = .Rows.Count
            iSourceColumn = .Cells(1, 1).Column
            iFirstSourceRow = .Cells(1, 1).Row
            If LookupRange.Value Like "*#*" Then
                iLastSourceRow = iFirstSourceRow + iRowCount - 1
            Else
                iLastSourceRow = .Range("A999999").End(xlUp).Row
            End If
            If iFirstSourceRow + .Rows.Count - 1 < iLastSourceRow Then
                iLastSourceRow = iFirstSourceRow + .Rows.Count - 1
            End If
        End With
        iStartIndex = 1
        'If LookupDialog.IgnoreHeaderCheckbox.Value = True Then
        '    iStartIndex = 2
        '    iRowCount = iRowCount - 1
        'End If
        Catalog.bTerminateLoop = False
        iTotal = iLastSourceRow - iFirstSourceRow + 1
        SearchingDialog.ProgressLabel = "Row 1 of " & iTotal
        SearchingDialog.Show
        'Iterate through rows, look up in catalog
        For i = iStartIndex To iTotal
            If Catalog.bTerminateLoop = True Then
                Exit For
            End If
            If Not oSourceRange.Rows(i).EntireRow.Hidden Then
                SearchingDialog.ProgressLabel = "Row " & i & " of " & iTotal
                Application.ScreenUpdating = False
                Dim sSearchString As String
                sSearchString = oSourceRange.Cells(i, 1).Value
                sSearchString = Replace(sSearchString, Chr(160), " ")
                sSearchString = Trim(sSearchString)
                If sSearchString <> "" Then
                    If sSearchString = "FALSE" Then
                        sResultRec = ""
                        sResultHold = ""
                    Else
                        sResultRec = Catalog.Lookup(sSearchString, sCatalogURL)
                        iHoldingsStart = InStr(2, sResultRec, "<?xml")
                        If iHoldingsStart > 0 Then
                            sResultHold = Mid(sResultRec, iHoldingsStart)
                            sResultRec = Left(sResultRec, iHoldingsStart - 1)
                        End If
                    End If
                    For j = 0 To LookupDialog.ResultTypeList.ListCount - 1
                        Dim stype As String
                        stype = LookupDialog.ResultTypeList.List(j)
                        stype = Replace(stype, "*", "")
                        If i = 1 And LookupDialog.IgnoreHeaderCheckbox.Value = True Then
                            sResult = ""
                            If LookupDialog.GenerateHeaderCheckBox.Value = True Then
                                sResult = stype
                            End If
                            GoTo NextRow
                        End If
                        If stype = "MMS ID" Or stype = "Catalog ID" Or _
                            (LookupDialog.CatalogURLBox.Value = "source:worldcat" And stype = "OCLC No.") Then
                            stype = "001"
                        ElseIf stype = "ISBN" Then
                            stype = "020"
                        ElseIf stype = "ISSN" Then
                            stype = "022"
                        ElseIf stype = "Title" Then
                            stype = "245"
                        ElseIf stype = "OCLC No." Then
                            stype = "035$a#(OCoLC)"
                        ElseIf stype = "Call No." Then
                            stype = "AVA$d"
                        ElseIf stype = "Location/DB Name" Then
                            stype = "AVA$bj|AVE$lm"
                        ElseIf stype = "Language code" Then
                            stype = "008(35,3)"
                        ElseIf stype = "Coverage" Then
                            stype = "AVA$t|AVE$s"
                        ElseIf InStr(1, stype, "Leader") = 1 Or InStr(1, stype, "LDR") Then
                            stype = Replace(stype, "Leader", "000")
                            stype = Replace(stype, "LDR", "000")
                        ElseIf stype = "True/False" Then
                            stype = "exists"
                        ElseIf stype = "ReCAP Holdings" Then
                            stype = "recap"
                        ElseIf stype = "BorrowDirect Holdings" Then
                            stype = "999$sp"
                        ElseIf stype = "WorldCat Holdings" Then
                            stype = "948$c"
                        End If
                        If sResultRec = "" Then
                            sResult = ""
                        ElseIf sResultRec = "INVALID" Then
                            sResult = "INVALID"
                        Else
                            If stype = "Barcode" Then
                                sResult = ExtractField(stype, CStr(sResultHold), True)
                            ElseIf stype = "Item Location" Or stype = "Item Enum/Chron" Or stype = "Shelf Locator" Then
                                sSearchType = CStr(LookupDialog.SearchFieldCombo.Value)
                                sBarcode = ""
                                If sSearchType = "Barcode" Or sSearchType = "alma.barcode" Then
                                    sResult = ExtractField(stype, CStr(sResultHold), True, sSearchString)
                                Else
                                    sResult = ExtractField(stype, CStr(sResultHold), True)
                                End If
                            Else
                                sResult = ExtractField(stype, CStr(sResultRec), False)
                                If sResult = "ERROR:InvalidRecap" Then
                                    MsgBox ("ReCAP queries do not support the result type: """ & LookupDialog.ResultTypeList.List(j) & """")
                                    SearchingDialog.Hide
                                    LookupDialog.Show
                                    Exit Sub
                                End If
                            End If
                            sResult = Trim(sResult)
                            iExtraBars = (Len(sResult) - Len(Replace(sResult, "|", ""))) - _
                                (Len(sSearchString) - Len(Replace(sSearchString, "|", "")))
                            If Right(sResult, 1) = "|" And iExtraBars <> 0 Then
                                sResult = Left(sResult, Len(sResult) - 1)
                            End If
                        End If
                        If sResult = "" Then
                            sResult = False
                        End If
NextRow:
                        oSourceRange.Cells(i, iResultColumn - iSourceColumn + 1 + j).NumberFormat = "@"
                        oSourceRange.Cells(i, iResultColumn - iSourceColumn + 1 + j).Value = sResult
                    Next j
                End If
                If ActiveWorkbook.Name = Catalog.sFileName And ActiveSheet.Name = Catalog.sSheetName Then
                    minRow = ActiveWindow.VisibleRange.Row
                    maxRow = minRow + ActiveWindow.VisibleRange.Rows.Count
                    If iFirstSourceRow + i <= minRow + 1 Or iFirstSourceRow + i >= maxRow - 1 Then
                        ActiveWindow.SmallScroll down:=(iFirstSourceRow + i - (maxRow + minRow) / 2) + 1
                    End If
                    Application.ScreenUpdating = True
                End If
                DoEvents
            End If
        Next i
        Application.ScreenUpdating = True
        SearchingDialog.Hide
    Else
        MsgBox ("Invalid Range Selected")
    End If
    If Catalog.bTerminateLoop Then
        LookupDialog.Show
        If UserPassForm.RememberCheckbox.Value = False Then
            UserPassForm.UserNameBox.Value = ""
            UserPassForm.PasswordBox.Value = ""
        End If
    Else
        LookupDialog.ResultTypeList.Clear
    End If
    End
End Sub


Private Sub OtherSourcesButton_Click()
    OtherSourcesDialog.Show
End Sub

Private Sub RemoveResultButton_Click()
    With LookupDialog.ResultTypeList
        If .ListIndex > -1 Then
            .RemoveItem (.ListIndex)
        End If
    End With
    RedrawButtons
End Sub

Private Sub RemoveURLButton_Click()
    If LookupDialog.CatalogURLBox.ListCount < 2 Then
        MsgBox ("Please add another URL before removing the last one")
        Exit Sub
    End If
    
    sCatalogURL = LookupDialog.CatalogURLBox.Value
    For i = 0 To LookupDialog.CatalogURLBox.ListCount - 1
        If sCatalogURL = LookupDialog.CatalogURLBox.List(i) Then
            LookupDialog.CatalogURLBox.RemoveItem (i)
            LookupDialog.CatalogURLBox.ListIndex = 0
            Exit For
        End If
    Next i
    Catalog.SetRegistryURLsFromCombo
End Sub

Private Sub ResultColumnSpinner_Change()
    LookupDialog.ResultColumnInput.Value = Catalog.ColumnLetterConvert(LookupDialog.ResultColumnSpinner.Value)
End Sub

Private Sub ResultTypeCombo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        AddResultButton_Click
        KeyCode = 0
    End If
End Sub

Private Sub ResultTypeList_Change()
    RedrawButtons
End Sub

Function SaveSet(sSetName As String) As Boolean
    If LookupDialog.ResultTypeList.ListCount = 0 Then
        MsgBox ("Please add at least one result type to the set")
        SaveSet = False
        Exit Function
    End If
    sSetString = sSetName
    For i = 0 To LookupDialog.ResultTypeList.ListCount - 1
        sSetString = sSetString & ChrW(166) & LookupDialog.ResultTypeList.List(i)
    Next i
    
    sAllSets = Catalog.GetFieldSets()
    aAllSets = Split(sAllSets, "|", -1, 0)
    sNewSets = ""
    bSetFound = False
    For i = 0 To UBound(aAllSets)
        If i > 0 Then
                sNewSets = sNewSets & "|"
        End If
        If InStr(1, aAllSets(i), sSetName + ChrW(166)) > 0 Then
            sNewSets = sNewSets & sSetString
            bSetFound = True
        Else
            sNewSets = sNewSets & aAllSets(i)
        End If
    Next i
    If Not bSetFound Then
        If sNewSets <> "" Then
            sNewSets = sNewSets & "|"
        End If
        sNewSets = sNewSets & sSetString
    End If
    Catalog.SetFieldSets (sNewSets)
    SaveSet = True
End Function

Private Sub SaveSetButton_Click()
    If LookupDialog.FieldSetList.ListIndex < 0 Then
        MsgBox ("Please select a set name")
        Exit Sub
    End If
    bSuccess = SaveSet(LookupDialog.FieldSetList.Value)
End Sub
