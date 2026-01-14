Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{DCD77FB3-24F4-409C-BC5B-DC371F305572}{5A6C456A-F3F1-4316-876A-289E14E689CD}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub