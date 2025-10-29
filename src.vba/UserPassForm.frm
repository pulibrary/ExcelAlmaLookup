Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{AF3C76FD-AF2C-4B77-A9AB-A65B3222147B}{3C90FFC3-78B0-4D8C-8D4D-69C3576E9888}"
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