Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{3AC374EA-30EC-4A6A-871A-EAD1E1121DD2}{687B2906-A853-484E-AF38-47101878B8B6}"
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