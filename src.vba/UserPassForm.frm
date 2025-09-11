Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{2AB75FDB-7B38-4FED-974B-E49319545E84}{6A187907-205E-4EB2-8FF9-111CBE32B0D1}"
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