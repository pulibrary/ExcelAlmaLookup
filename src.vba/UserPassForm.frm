Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{4DA4C497-92F3-4BDA-A4F8-AE8048E02F88}{5204F4E3-9731-4B0A-892C-3D7CEC048CD8}"
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