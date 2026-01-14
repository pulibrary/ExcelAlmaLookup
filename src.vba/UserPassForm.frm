Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{BB106FAF-7839-4D98-9CF7-5CD30E64530C}{F0A0A166-D6AD-4F7D-A445-468E7E556088}"
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