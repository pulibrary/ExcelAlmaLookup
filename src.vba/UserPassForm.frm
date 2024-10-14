Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{5AE6FA8E-4A3E-49CA-983F-FC1CFFAC9004}{BDDB0E2F-0C8F-4186-94CF-4E818F828630}"
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