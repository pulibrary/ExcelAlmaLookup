Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{8CEE2961-E9FD-4357-908A-292B0077D8D4}{A13F909C-A680-4844-9C2D-0A9D8F132CC0}"
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