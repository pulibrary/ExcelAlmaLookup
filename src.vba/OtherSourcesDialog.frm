Attribute VB_Name = "OtherSourcesDialog"
Attribute VB_Base = "0{D4152999-EE08-4247-8FD0-1F4825022CA6}{FD2AB2F0-1F39-4A72-8155-51D1F407C4AB}"
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