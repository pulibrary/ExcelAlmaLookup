Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{3BB2AB4C-35F1-45CD-B48C-0E26829C9225}{BFF424B0-D75E-4AAA-96CB-33DCEB17184E}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub