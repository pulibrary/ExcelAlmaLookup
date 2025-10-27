Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{70D861E9-4346-4E52-AE45-A1438D9F20E2}{12DE07CF-5294-4940-9786-17B864724C9A}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub