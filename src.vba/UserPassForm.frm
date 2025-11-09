Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{514DFDE4-8807-4DDC-9565-0041951E37BB}{18338D86-79EA-42E5-921F-0EE97EB28128}"
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