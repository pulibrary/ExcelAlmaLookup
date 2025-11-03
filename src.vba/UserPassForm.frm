Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{7264B087-61BF-47BC-A625-D983A8337D3C}{7B61155D-41B4-4874-B408-1BD602B64281}"
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