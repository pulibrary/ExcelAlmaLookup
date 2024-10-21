Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{3DEC4C6F-8ED6-4BF4-B0B0-414CCF8B906B}{A6CBD480-9CD0-4BFF-B660-52F8590A36BD}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub