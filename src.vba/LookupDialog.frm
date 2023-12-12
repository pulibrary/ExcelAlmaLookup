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
    aFieldMap = Catalog.GetAllFields()
    If Not IsNull(aFieldMap) Then
        Catalog.SetRegistryURLsFromCombo
    End If
End Sub

Private Sub CancelButton_Click()
    LookupDialog.Hide
End Sub

Private Sub CatalogURLBox_Change()
    Catalog.sAuth = ""
    sCatalogAuths = GetRegistryAuths()
    aCatalogAuths = Split(sCatalogAuths, "|")
    For i = 0 To UBound(aCatalogAuths)
        aURLAuth = Split(aCatalogAuths(i), "¦")
        If aURLAuth(0) = LookupDialog.CatalogURLBox.Text Then
            sAuth = aURLAuth(1)
            Exit For
        End If
    Next i
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
    aFieldSets = Split(sFieldSets, "|")
    sNewSets = ""
    For i = 0 To UBound(aFieldSets)
        If InStr(1, aFieldSets(i), sSelectedSet + "¦") = 0 Then
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
    aFieldSets = Split(sFieldSets, "|")
    For i = 0 To UBound(aFieldSets)
        If InStr(1, aFieldSets(i), sSetName & "¦") > 0 Then
            aFields = Split(aFieldSets(i), "¦")
            For j = 1 To UBound(aFields)
                LookupDialog.ResultTypeList.AddItem aFields(j)
            Next j
        End If
    Next i
End Function

Private Sub HelpButton_Click()
    ThisWorkbook.FollowHyperlink Catalog.sRepoURL & "#readme"
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
    If InStr(1, sNewName, "|") > 0 Or InStr(1, sNewName, Chr(166)) Then
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
    aFieldMap = Catalog.GetAllFields()
    If IsNull(aFieldMap) Then
        Exit Sub
    End If
    Dim sCatalogURL As String
    sCatalogURL = CStr(LookupDialog.CatalogURLBox.Text)
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
    Set oSourceRange = Range(LookupRange.Value)
    
    If IsObject(oSourceRange) Then
        LookupDialog.Hide
        With oSourceRange
            iRowCount = oSourceRange.Rows.Count
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
        If LookupDialog.IgnoreHeaderCheckbox.Value = True Then
            iFirstSourceRow = iFirstSourceRow + 1
            iRowCount = iRowCount - 1
        End If
        Catalog.bTerminateLoop = False
        iTotal = iLastSourceRow - iFirstSourceRow + 1
        SearchingDialog.ProgressLabel = "Row 1 of " & iTotal
        SearchingDialog.Show
        'Iterate through rows, look up in catalog
        For i = iFirstSourceRow To iLastSourceRow
            If Catalog.bTerminateLoop = True Then
                Exit For
            End If
            If Not Rows(i).EntireRow.Hidden Then
                SearchingDialog.ProgressLabel = "Row " & (i - iFirstSourceRow + 1) & " of " & iTotal
                Application.ScreenUpdating = False
                Dim sSearchString As String
                sSearchString = Cells(i, iSourceColumn).Value
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
                        Dim sType As String
                        sType = LookupDialog.ResultTypeList.List(j)
                        sType = Replace(sType, "*", "")
                        If sType = "MMS ID" Then
                            sType = "001"
                        ElseIf sType = "ISBN" Then
                            sType = "020"
                        ElseIf sType = "Title" Then
                            sType = "245"
                        ElseIf sType = "Call No." Then
                            sType = "AVA$d"
                        ElseIf sType = "Location/DB Name" Then
                            sType = "AVA$bj|AVE$lm"
                        ElseIf sType = "Language code" Then
                            sType = "008(35,3)"
                        ElseIf sType = "Coverage" Then
                            sType = "AVA$t|AVE$s"
                        ElseIf sType = "Leader" Then
                            sType = "000"
                        ElseIf sType = "True/False" Then
                            sType = "exists"
                        End If
                        If sResultRec = "" Then
                            sResult = ""
                        Else
                            If sType = "Barcode" Then
                                sResult = ExtractField(sType, CStr(sResultHold), True)
                            Else
                                sResult = ExtractField(sType, CStr(sResultRec), False)
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
                        Cells(i, iResultColumn + j).NumberFormat = "@"
                        Cells(i, iResultColumn + j).Value = sResult
                    Next j
                End If
                minRow = ActiveWindow.VisibleRange.Row
                maxRow = minRow + ActiveWindow.VisibleRange.Rows.Count
                If i <= minRow + 1 Or i >= maxRow - 1 Then
                    ActiveWindow.SmallScroll down:=(i - (maxRow + minRow) / 2) + 1
                End If
                Application.ScreenUpdating = True
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

Private Sub RecapCheckBox_Click()
    PopulateCombos
    RedrawButtons
    LookupDialog.ResultTypeList.Clear
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
        sSetString = sSetString & "¦" & LookupDialog.ResultTypeList.List(i)
    Next i
    
    sAllSets = Catalog.GetFieldSets()
    aAllSets = Split(sAllSets, "|")
    sNewSets = ""
    bSetFound = False
    For i = 0 To UBound(aAllSets)
        If i > 0 Then
                sNewSets = sNewSets & "|"
        End If
        If InStr(1, aAllSets(i), sSetName + "¦") > 0 Then
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