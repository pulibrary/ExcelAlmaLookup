Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{4A38C425-6F39-477D-AE1F-EB5034B1E365}{C265D129-4BA4-404D-873D-19C93062BE57}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub