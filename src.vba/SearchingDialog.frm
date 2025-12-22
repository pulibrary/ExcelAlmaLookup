Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{CD02A35B-8714-402B-B569-323DA6122EF9}{271CC0D7-A40A-49BD-A137-A2FD7D8C9113}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub