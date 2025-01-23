Attribute VB_Name = "AdditionalFieldsDialog"
Attribute VB_Base = "0{9508A44C-2F3B-446F-9BDF-0F8DB2BD2AEB}{7F5DA57A-A81B-4BCF-B7CC-9C079F088481}"
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
        sIndex = AdditionalFieldsDialog.SRUFields.List(i, 1)
        LookupDialog.SearchFieldCombo.Value = sIndex
        LookupDialog.SearchFieldCombo.AddItem sIndex
    End If
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