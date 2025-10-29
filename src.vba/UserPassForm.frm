Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{866713DF-9A58-4916-B0BF-BD3AE7C97DD7}{37D39188-8165-4E85-8883-FFA94E49C7EC}"
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