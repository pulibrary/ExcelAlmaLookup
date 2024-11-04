Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{000AB31D-B2C9-4D1C-ADDA-45C810528F37}{6BE559E1-68FA-4FA9-8099-EC1729C29653}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub