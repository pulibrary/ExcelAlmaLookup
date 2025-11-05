Attribute VB_Name = "LookupDialog"
Attribute VB_Base = "0{253DE61E-115B-4C18-886D-4F1412E1F2A3}{7773D981-29EF-4C89-84E7-98C490B61DB3}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
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

Private Sub AddSearchButton_Click()
    If SearchFieldCombo.Value = "" Or OperatorCombo.Value = "" Or SearchValueBox.Value = "" Then
        MsgBox ("Please make sure the search type, operator, and value are all filled in.")
    End If
    SearchListBox.AddItem ""
    iIndex = SearchListBox.ListCount - 1
    SearchListBox.List(iIndex, 0) = BooleanCombo.Value
    SearchListBox.List(iIndex, 1) = SearchFieldCombo.Value
    SearchListBox.List(iIndex, 2) = OperatorCombo.Value
    SearchListBox.List(iIndex, 3) = SearchValueBox.Value
    Catalog.PopulateBooleanCombo
End Sub

Private Sub AddURLButton_Click()
    If Catalog.bIsAlma Then
        aFieldMap = Catalog.GetAllFields()
        If Not IsNull(aFieldMap) Then
            Catalog.AddURLtoRegistry CatalogURLBox.Value
        End If
    Else
        Catalog.AddURLtoRegistry CatalogURLBox.Value
    End If
    Catalog.PopulateCombos
End Sub

Private Sub CancelButton_Click()
    Catalog.SaveSearchParams
    Catalog.SaveResultList
    LookupDialog.Hide
    End
End Sub

Private Sub CatalogURLBox_Change()
    If CatalogURLBox.Value <> "" Then
        Catalog.sAuth = ""
        iSelected = Catalog.GetSourceRegIndex(CatalogURLBox.Value)
        sAuth = GetSetting(Catalog.sRegistryDir, "Sources", "AUTH" & Format(iSelected, "000"), "")
        SaveSetting Catalog.sRegistryDir, "Sources", "SELECTED", CatalogURLBox.Value
        Catalog.bIsAlma = True
        Catalog.bIsWorldCat = False
        If InStr(1, LookupDialog.CatalogURLBox, "source:") = 1 Then
            If Catalog.bIsAlma = True Then
                Catalog.bIsAlma = False
                Catalog.aExplainFields = Empty
            End If
            If LookupDialog.CatalogURLBox = "source:worldcat" Then
                Catalog.bIsWorldCat = True
            End If
        End If
        Catalog.PopulateSourceDependentOptions
        Catalog.PopulateBooleanCombo
    End If
End Sub

Private Sub ClearCredentialsButton_Click()
    Catalog.ClearRegistryAuth (CatalogURLBox.Value)
    Catalog.sAuth = ""
End Sub

Private Sub DeleteSetButton_Click()
    iSetIndex = LookupDialog.FieldSetList.ListIndex
    If iSetIndex < 0 Then
        MsgBox ("Please select a set name")
        Exit Sub
    End If
    
    sSelectedSet = LookupDialog.FieldSetList.List(iSetIndex)
    Catalog.DeleteFieldSet sSelectedSet
    LookupDialog.FieldSetList.RemoveItem iSetIndex
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
    
    iMax = GetSetting(Catalog.sRegistryDir, "FieldSets", "MAXALL", -1)
    For i = 0 To iMax
        sRegName = GetSetting(Catalog.sRegistryDir, "FieldSets", "NAME" & Format(i, "000"), "")
        If sRegName = sSetName Then
            iSetMax = GetSetting(Catalog.sRegistryDir, "FieldSets", "MAX" & Format(i, "000"), -1)
            For j = 0 To iSetMax
                sField = GetSetting(Catalog.sRegistryDir, "FieldSets", "FIELD" & Format(i, "000") & "-" & Format(j, "000"))
                LookupDialog.ResultTypeList.AddItem sField
            Next j
            LoadSet = True
            Exit Function
        End If
    Next i
    LoadSet = False
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
    bError = False
    sNewName = InputBox("Enter the Name of the New Set", "New Set")
    If sNewName = "" Then
        bError = True
    End If
    If Not bError And InStr(1, sNewName, "|") > 0 Or InStr(1, sNewName, ChrW(166)) Then
        MsgBox ("Set name cannot contain vertical bar characters")
        bError = True
    End If
    If Not bError Then
        iMax = GetSetting(Catalog.sRegistryDir, "FieldSets", "MAXALL", -1)
        For i = 0 To iMax
            sRegName = GetSetting(Catalog.sRegistryDir, "FieldSets", "NAME" & Format(i, "000"), "")
            If sRegName = sNewName Then
                MsgBox ("Set name already exists")
                bError = True
            End If
        Next i
    End If
    If Not bError Then
        bSuccess = Catalog.SaveFieldSet(sNewName)
        If bSuccess Then
            LookupDialog.FieldSetList.AddItem sNewName
            LookupDialog.FieldSetList.ListIndex = LookupDialog.FieldSetList.ListCount - 1
        End If
    End If
End Sub

Private Sub OKButton_Click()
    If Catalog.bIsAlma Then
        aFieldMap = Catalog.GetAllFields()
        If IsNull(aFieldMap) Then
            Exit Sub
        End If
    End If
    
    If Not ValidateMaxResults Then
        Exit Sub
    End If
    
    Dim sCatalogURL As String
    sCatalogURL = CStr(LookupDialog.CatalogURLBox.Text)
    If Catalog.bIsWorldCat Then
        bSuccess = Catalog.Z3950Connect(sCatalogURL)
        If Not bSuccess Then
            Exit Sub
        End If
    End If
    
    Catalog.AddURLtoRegistry (sCatalogURL)
    Catalog.SaveSearchParams
    Catalog.SaveResultList
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
    
    If IsObject(oSourceRange) Then
        LookupDialog.Hide
        iRowCount = oSourceRange.Rows.Count
        iSourceColumn = oSourceRange.Cells(1, 1).Column
        iFirstSourceRow = oSourceRange.Cells(1, 1).Row
        iLastSourceRow = Catalog.GetLastPopulatedRow(oSourceRange)
        iStartIndex = 1
        Catalog.bTerminateLoop = False
        iTotal = iLastSourceRow - iFirstSourceRow + 1
        SearchingDialog.ProgressLabel = "Row 1 of " & iTotal
        SearchingDialog.Show
        'Iterate through rows, look up in catalog
        bIgnoreHeader = LookupDialog.IgnoreHeaderCheckbox.Value
        For i = iStartIndex To iTotal
            If Catalog.bTerminateLoop = True Then
                Exit For
            End If
            If Not oSourceRange.Rows(i).EntireRow.Hidden Then
                SearchingDialog.ProgressLabel = "Row " & i & " of " & iTotal
                Application.ScreenUpdating = False
                Dim sSearchString As String
                Set oSearchRow = oSourceRange.Rows(i).EntireRow
                If i = iStartIndex And bIgnoreHeader Then
                    sResultRec = "HEADER"
                Else
                    sResultRec = Catalog.Lookup(oSearchRow, sCatalogURL)
                End If
                If sResultRec <> "" Then
                    iHoldingsStart = InStr(2, sResultRec, "<?xml")
                    If iHoldingsStart > 0 Then
                        sResultHold = Mid(sResultRec, iHoldingsStart)
                        sResultRec = Left(sResultRec, iHoldingsStart - 1)
                    End If
                    For j = 0 To LookupDialog.ResultTypeList.ListCount - 1
                        Dim stype As String
                        stype = LookupDialog.ResultTypeList.List(j)
                        While Left(stype, 1) = "*"
                            stype = Mid(stype, 2)
                        Wend
                        If i = iStartIndex And bIgnoreHeader Then
                            sResult = ""
                            If LookupDialog.GenerateHeaderCheckBox.Value = True Then
                                sResult = stype
                            End If
                            GoTo NextRow
                        End If
                        If stype = "MMS ID" Or stype = "Catalog ID" Or _
                            (Catalog.bIsWorldCat And stype = "OCLC No.") Then
                            stype = "001"
                        ElseIf stype = "LCCN" Then
                            stype = "010"
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
                        ElseIf stype = "ReCAP CGD" Then
                            stype = "recap_cgd"
                        ElseIf stype = "BorrowDirect Holdings" Then
                            stype = "999$sp"
                        ElseIf stype = "WorldCat Holdings" Then
                            stype = "948$ch"
                        ElseIf Catalog.bIsWorldCat And stype = "Holdings Count" Then
                            stype = "948$ch#"
                        End If
                        If sResultRec = "" Then
                            sResult = ""
                        ElseIf sResultRec = "INVALID" Then
                            sResult = "INVALID"
                        ElseIf sResultRec = "TOO MANY HOLDINGS" Then
                            sResult = "TOO MANY HOLDINGS"
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
                Else
                    For j = 0 To LookupDialog.ResultTypeList.ListCount - 1
                        oSourceRange.Cells(i, iResultColumn - iSourceColumn + 1 + j).NumberFormat = "@"
                        oSourceRange.Cells(i, iResultColumn - iSourceColumn + 1 + j).Value = "FALSE"
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
    
End Sub

Private Sub OperatorCombo_Change()
    If OperatorCombo.Value = "empty" Then
        SearchValueBox.Value = "is empty"
        SearchValueBox.Enabled = False
    ElseIf SearchValueBox.Enabled = False Then
        SearchValueBox.Value = ""
        SearchValueBox.Enabled = True
    End If
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

Private Sub RemoveSearchButton_Click()
    With LookupDialog.SearchListBox
        iSelected = .ListIndex
        If iSelected > -1 Then
            .RemoveItem (iSelected)
        End If
        If .ListIndex = -1 Then
            LookupDialog.RemoveSearchButton.Enabled = False
            LookupDialog.BooleanCombo.Enabled = False
            LookupDialog.BooleanCombo.Value = ""
        End If
        If .ListCount > 0 Then
            .List(0, 0) = ""
        End If
    End With
    Catalog.PopulateBooleanCombo
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
    Catalog.RemoveURLfromRegistry (sCatalogURL)
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


Private Sub SaveSetButton_Click()
    If LookupDialog.FieldSetList.ListIndex < 0 Then
        MsgBox ("Please select a set name")
        Exit Sub
    End If
    bSuccess = Catalog.SaveFieldSet(LookupDialog.FieldSetList.Value)
End Sub
Private Sub SearchFieldCombo_Change()
    sField = LookupDialog.SearchFieldCombo.Value
    Catalog.PopulateBooleanCombo
    If sField = "Other fields..." Then
        AdditionalFieldsDialog.FilterBox.Value = ""
        If Catalog.bIsAlma Then
            If IsNull(Catalog.aExplainFields) Or IsEmpty(Catalog.aExplainFields) Then
                Catalog.aExplainFields = Catalog.GetAllFields()
            End If
            If Not IsNull(Catalog.aExplainFields) Then
                AdditionalFieldsDialog.SRUFields.List = Catalog.aExplainFields
            Else
                MsgBox ("Cannot access catalog.  Please confirm the Alma URL is correct.")
                Exit Sub
            End If
        ElseIf Catalog.bIsWorldCat Then
            AdditionalFieldsDialog.SRUFields.List = Catalog.aOCLCSearchKeys
        End If
        AdditionalFieldsDialog.Show
    End If
End Sub

Private Sub SearchFieldCombo_AfterUpdate()
    Catalog.PopulateOperatorCombo
    Catalog.PopulateBooleanCombo
End Sub

Private Sub SearchListBox_Click()
    If LookupDialog.SearchListBox.Enabled Then
        If LookupDialog.SearchListBox.ListIndex > -1 Then
            LookupDialog.RemoveSearchButton.Enabled = True
        Else
            LookupDialog.RemoveSearchButton.Enabled = False
        End If
    End If
End Sub


Private Sub SearchValueBox_Change()
    If LookupDialog.SearchListBox.Enabled Then
        If LookupDialog.SearchValueBox.Value <> "" Then
            LookupDialog.AddSearchButton.Enabled = True
        Else
            LookupDialog.AddSearchButton.Enabled = False
        End If
    End If
End Sub

Private Sub UserForm_Activate()
    If InStr(1, LookupDialog.CatalogURLBox.Value, "http") = 1 Then
        Catalog.aExplainFields = Catalog.GetAllFields()
    End If
End Sub