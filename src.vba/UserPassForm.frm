Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{8F741E11-EA4D-41D0-8DE7-547519155987}{1A71D21D-473A-420E-AF64-55E70914CE5E}"
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