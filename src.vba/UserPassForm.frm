Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{D6FD280A-B5FA-485E-BFB4-61EB5C5BDB69}{7CBD0757-8D85-4601-B0EF-89241B2DD7A2}"
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