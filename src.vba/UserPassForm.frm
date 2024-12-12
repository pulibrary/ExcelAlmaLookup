Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{8C8DC843-EE4C-4B36-B0A7-E008B8EC7F67}{C8C1503E-CC31-467C-9784-186E61B3A0AB}"
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