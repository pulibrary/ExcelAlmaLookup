Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{C18808AC-A414-4C18-A972-85741573AFB4}{23D3D420-B567-4E25-A105-24E2E7F2F663}"
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