Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{4771FB3D-BDA7-4213-AFA0-0C81D6B2E3C7}{F240E14E-D767-4D65-960D-9978BA354A83}"
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