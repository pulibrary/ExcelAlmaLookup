Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{7E658608-5CE3-4296-834C-F5650555F278}{E5C14E3E-1F0B-4B5D-9DF4-CAF55CFED18F}"
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