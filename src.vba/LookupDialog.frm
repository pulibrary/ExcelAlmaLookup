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

Private Sub CancelButton_Click()
    LookupDialog.Hide
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

Private Sub OKButton_Click()
    
    
    Dim sCatalogURL As String
    sCatalogURL = CStr(LookupDialog.CatalogURLBox.Text)
    
    If sCatalogURL = "" Then
        MsgBox ("Please enter a URL for the catalog (e.g. 'https://catalog.princeton.edu')")
        Exit Sub
    Else
        CreateObject("WScript.Shell").RegWrite "HKCU\Software\Excel Local Catalog Lookup\CatalogURL", sCatalogURL, "REG_SZ"
    End If

    iResultColumn = LookupDialog.ResultColumnSpinner.Value
    If LookupDialog.ResultTypeList.ListCount = 0 Then
        AddResultButton_Click
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
                iLastSourceRow = .Range("A65534").End(xlUp).Row
            End If
            If iFirstSourceRow + .Rows.Count - 1 < iLastSourceRow Then
                iLastSourceRow = iFirstSourceRow + .Rows.Count - 1
            End If
        End With
        
        'Iterate through rows, look up in catalog
        For i = iFirstSourceRow To iLastSourceRow
            Application.ScreenUpdating = False
            Dim sSearchString As String
            sSearchString = Cells(i, iSourceColumn).Value
            If sSearchString <> "" Then
                sResultRec = Catalog.Lookup(sSearchString, sCatalogURL)
                For j = 0 To LookupDialog.ResultTypeList.ListCount - 1
                    Dim sType As String
                    sType = LookupDialog.ResultTypeList.List(j)
                    If sType = "MMS ID" Then
                        sType = "001"
                    ElseIf sType = "ISBN" Then
                        sType = "020"
                    ElseIf sType = "Title" Then
                        sType = "245"
                    ElseIf sType = "Call No." Then
                        sType = "AVA$d"
                    ElseIf sType = "Location(s)" Then
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
                    sResult = ExtractField(sType, CStr(sResultRec))
                    sResult = Trim(sResult)
                    iExtraBars = (Len(sResult) - Len(Replace(sResult, "|", ""))) - _
                        (Len(sSearchString) - Len(Replace(sSearchString, "|", "")))
                    'Debug.Print iExtraBars & " " & sSearchString & " " & sResult
                    If Right(sResult, 1) = "|" And iExtraBars <> 0 Then
                        sResult = Left(sResult, Len(sResult) - 1)
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
        Next i
    Else
        MsgBox ("Invalid Range Selected")
    End If
    LookupDialog.ResultTypeList.Clear
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

Private Sub ResultColumnSpinner_Change()
    LookupDialog.ResultColumnInput.Value = Catalog.ColumnLetterConvert(LookupDialog.ResultColumnSpinner.Value)
End Sub


Private Sub ResultsGroup_Click()

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