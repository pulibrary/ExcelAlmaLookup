Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{B5672717-CC16-4EF5-8C89-42AB87C97617}{D188C8BD-B5EA-44DB-9D8A-C1A0EBBB121F}"
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