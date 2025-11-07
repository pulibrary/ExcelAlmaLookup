Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{2C576B8A-94FC-43F9-98B5-1894A583B9AF}{1B736438-A155-4BEA-A91F-36C71358101C}"
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