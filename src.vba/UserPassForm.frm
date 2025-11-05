Attribute VB_Name = "UserPassForm"
Attribute VB_Base = "0{73E66ACC-4768-4FAC-933D-65FE9E14910D}{002CA137-105A-4133-B65F-FD5574A1C766}"
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