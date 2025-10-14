Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{2FC8C762-7B20-4FBB-B128-D5AF24A1FE62}{BF91A846-6DF1-43B7-853F-39359DD86487}"
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