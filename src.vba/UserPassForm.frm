Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{ADFABBC6-DA2E-4F8E-BD98-0051EDEBD913}{5C4DE414-EE1C-417D-BAC0-17B19E0C353C}"
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