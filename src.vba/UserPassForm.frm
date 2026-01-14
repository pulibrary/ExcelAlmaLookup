Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{E616AF91-DFEF-41E8-80C8-62EFA15C4191}{72D0D002-D6FF-4415-9ABC-854612C9C8C2}"
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