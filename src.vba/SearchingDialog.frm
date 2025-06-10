Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{1E1CC6D6-7596-47CD-8B56-78C6B5ACE139}{51EA2428-620D-47A8-94BF-6FD81DCE21DE}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub