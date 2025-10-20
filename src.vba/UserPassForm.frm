Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{B67A60C2-A288-4F36-886E-54F8E1B4B721}{93510F67-1D8C-48BB-8E17-E2C613132EB6}"
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