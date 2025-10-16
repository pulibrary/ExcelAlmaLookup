Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{8D82CC20-697F-4D26-9E50-47FF541002DF}{C55EC390-BDBF-48B1-9E29-78D862D7505B}"
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