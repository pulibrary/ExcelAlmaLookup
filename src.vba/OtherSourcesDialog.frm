Attribute VB_Name = "OtherSourcesDialog"
Attribute VB_Base = "0{242E5854-D1E5-410C-BC90-9CAEEDAB54C2}{5FFC3E6C-E22D-44A5-B9F5-244D67D8C8DF}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub PopulateURLField()
    i = OtherSourcesDialog.OtherSourcesListBox.ListIndex
    If i = -1 Then
        MsgBox ("No source is selected")
    Else
        sIndex = OtherSourcesDialog.OtherSourcesListBox.List(i, 0)
        LookupDialog.CatalogURLBox.Value = sIndex
    End If
End Sub
Private Sub CancelButton_Click()
    OtherSourcesDialog.Hide
End Sub


Private Sub SelectButton_Click()
    PopulateURLField
    OtherSourcesDialog.Hide
End Sub

Private Sub OtherSourcesListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    PopulateURLField
    OtherSourcesDialog.Hide
End Sub