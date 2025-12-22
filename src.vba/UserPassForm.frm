Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{F0DB0641-1CDD-4C09-A8D6-AB9B0CAF178C}{524FE1F3-FEC7-4E98-A57F-CB797F2CB2F9}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bKeepTryingURL = False
    UserPassForm.Hide
End Sub

Private Sub LoginButton_Click()
    UserPassForm.Hide
End Sub