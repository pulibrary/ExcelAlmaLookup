Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{9AE0C57C-9D21-49B7-9230-58C4F260E8D5}{05282309-3FB6-4D33-8EB0-708B185BE850}"
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