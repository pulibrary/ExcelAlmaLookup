Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{41521B2B-F8BA-41DE-A9F1-FAF4C38C147A}{00FB3352-FA08-4C7F-95E2-D8E9A0B87A6A}"
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