Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{29138481-553D-45E5-823A-217782CB2BA9}{2C4BA852-E26F-43D7-B568-55C4516199DB}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub