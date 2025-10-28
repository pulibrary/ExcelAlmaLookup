Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{35CD8BD3-82C8-442D-840A-EDF0F288584D}{580DBCF5-9D04-4047-B54E-927EFE99BAE2}"
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