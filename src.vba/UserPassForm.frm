Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{214815F7-4695-4C94-B829-3EA39BBC8BD9}{2EC08098-298A-4963-AA7D-9958A385C804}"
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