Attribute VB_Name = "OtherSourcesDialog"
Attribute VB_Base = "0{F3BB48B4-B517-4B67-9A98-3C2C0DDC0383}{09A56E2A-4B28-4731-8349-D2ECA4A3A8DB}"
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