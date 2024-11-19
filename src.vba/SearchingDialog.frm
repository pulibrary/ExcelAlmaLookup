Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{DAC5A3CD-7657-4EA2-B1E6-63A715977798}{02B38624-78C3-44E1-A717-3F98F4623EDD}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub