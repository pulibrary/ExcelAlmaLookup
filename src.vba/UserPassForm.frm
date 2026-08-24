Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{10344DC0-CEE5-43B9-A941-1B531F9A80B2}{4B40036C-E3FA-4DB4-B698-345F5E9F1C46}"
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