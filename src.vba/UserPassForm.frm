Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{70669C1D-8A3E-46F6-A79A-A12E75EC7E16}{CEDF8C59-17F7-488C-8875-D92AEDBABCE3}"
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