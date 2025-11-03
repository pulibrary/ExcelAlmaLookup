Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{545EE37F-9E1C-4E74-8BDE-F355DD241664}{BF125AB9-F672-4DA6-BF10-EC1390BC5FD1}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub