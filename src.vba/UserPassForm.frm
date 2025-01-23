Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{595340A9-79D8-4FDB-A74D-76C6C8523B45}{3C455B55-D07F-4C51-A740-6C5C46CC99FC}"
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