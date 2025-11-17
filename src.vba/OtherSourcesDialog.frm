Attribute VB_Name = "OtherSourcesDialog"
Attribute VB_Base = "0{9C23F97B-2F27-4390-9C5D-4D81151A1CA0}{18C8B4EA-6BE1-460A-BDBA-25DA18F62ACE}"
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