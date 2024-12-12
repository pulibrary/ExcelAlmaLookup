Attribute VB_Name = "OtherSourcesDialog"
Attribute VB_Base = "0{10F93E85-A413-40DC-B121-8CE12F0D8E42}{07339DE5-75CB-4A17-B29A-5876811599D8}"
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