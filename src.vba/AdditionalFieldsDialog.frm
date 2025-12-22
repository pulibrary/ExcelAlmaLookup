Attribute VB_Name = "AdditionalFieldsDialog"
Attribute VB_Base = "0{E5F35E4B-4B52-42B6-93A2-CF4BA9BF850A}{30279A8B-DFA1-4E6E-BA25-25304F8A3F46}"
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