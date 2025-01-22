Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{E74F17F6-B2C5-4063-8FF6-A5E210473BD7}{8B91880E-D545-4DEA-9EE3-C301C5E5FFB9}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub