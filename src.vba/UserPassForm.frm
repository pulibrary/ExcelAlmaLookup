Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{6F776CDC-C8D2-4D6E-9955-D8027699D959}{DD0B689A-D11C-44B9-936D-282DA78D7D52}"
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