Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{F0F267C3-337E-4ED1-9463-6EA96B82E01A}{3289E727-BA5F-4537-B407-2D10257C3A51}"
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