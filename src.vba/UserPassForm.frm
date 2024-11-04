Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{26C754C2-76CE-4185-BE99-D45C3933527B}{629BE6BF-24FC-4AA8-819D-E4D504F74981}"
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