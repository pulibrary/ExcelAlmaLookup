Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{EEB52D4F-000A-431A-B79F-EB0C320A47A9}{6BD66FED-3F05-4B29-879F-0378017580E8}"
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