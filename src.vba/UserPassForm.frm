Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{3A29B6C6-3019-417C-88B6-1B78BD9851AD}{8D0B1F01-1165-4A29-9F48-ECEE2F4A4F3C}"
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