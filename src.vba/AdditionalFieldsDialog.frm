Attribute VB_Name = "AdditionalFieldsDialog"
Attribute VB_Base = "0{FA48B1CA-3DDF-43FF-955F-366A6ECEA367}{A4911A58-2E3F-455B-977D-7DD0CC4F7E71}"
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
            If Catalog.bIsAlma Then
                .AddItem AdditionalFieldsDialog.SRUFields.List(i, 1), .ListCount - 1
            ElseIf LookupDialog.CatalogURLBox = "source:worldcat" Then
                .AddItem AdditionalFieldsDialog.SRUFields.List(i, 0), .ListCount - 1
            End If
            .ListIndex = .ListCount - 2
        End With
    End If
    Catalog.PopulateOperatorCombo
End Sub

Private Sub CancelAdditionalField_Click()
    LookupDialog.SearchFieldCombo.ListIndex = 0
    AdditionalFieldsDialog.Hide
End Sub

Private Sub FilterBox_Change()
    sFilterText = LCase(AdditionalFieldsDialog.FilterBox.Value)
    AdditionalFieldsDialog.SRUFields.Clear
    Dim aIndexFields As Variant
    
    If Catalog.bIsAlma Then
        aIndexFields = Catalog.aExplainFields
    ElseIf Catalog.bIsWorldCat Then
        aIndexFields = Catalog.aOCLCSearchKeys
    End If
    If sFilterText = "" Then
        AdditionalFieldsDialog.SRUFields.List = aIndexFields
        Exit Sub
    End If
    iFilterCount = 0
    For i = 0 To UBound(aIndexFields)
        If InStr(1, LCase(aIndexFields(i, 0) & "|" & aIndexFields(i, 1)), sFilterText) > 0 Then
            AdditionalFieldsDialog.SRUFields.AddItem
            AdditionalFieldsDialog.SRUFields.List(iFilterCount, 0) = aIndexFields(i, 0)
            AdditionalFieldsDialog.SRUFields.List(iFilterCount, 1) = aIndexFields(i, 1)
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