Attribute VB_Name = "AdditionalFieldsDialog"
Attribute VB_Base = "0{EC6688E6-4DA8-4E0A-8610-29EEA530C753}{42AED507-D364-4031-8E6B-6F8EAC0AD77F}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub PopulateSearchField()
    i = AdditionalFieldsDialog.SRUFields.ListIndex
    If i = -1 Then
        MsgBox ("No field is selected")
    Else
        With LookupDialog.SearchFieldCombo
            .AddItem AdditionalFieldsDialog.SRUFields.List(i, 1), .ListCount - 1
            .ListIndex = .ListCount - 2
        End With
    End If
    Catalog.PopulateOperatorCombo
End Sub

Private Sub CancelAdditionalField_Click()
    AdditionalFieldsDialog.Hide
End Sub

Private Sub FilterBox_Change()
    sFilterText = LCase(AdditionalFieldsDialog.FilterBox.Value)
    AdditionalFieldsDialog.SRUFields.Clear
    If sFilterText = "" Then
        AdditionalFieldsDialog.SRUFields.List = Catalog.aExplainFields
        Exit Sub
    End If
    iFilterCount = 0
    For i = 0 To UBound(Catalog.aExplainFields)
        If InStr(1, LCase(aExplainFields(i, 0) & "|" & aExplainFields(i, 1)), sFilterText) > 0 Then
            AdditionalFieldsDialog.SRUFields.AddItem
            AdditionalFieldsDialog.SRUFields.List(iFilterCount, 0) = aExplainFields(i, 0)
            AdditionalFieldsDialog.SRUFields.List(iFilterCount, 1) = aExplainFields(i, 1)
            iFilterCount = iFilterCount + 1
        End If
    Next i
End Sub

Private Sub SelectAdditionalField_Click()
    PopulateSearchField
    AdditionalFieldsDialog.Hide
End Sub

Private Sub SRUFields_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    PopulateSearchField
    AdditionalFieldsDialog.Hide
End Sub