Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{C788974F-D535-424F-9664-2F7BD43C7F34}{C7D89E41-948B-419B-AA00-C45824EFC0C3}"
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