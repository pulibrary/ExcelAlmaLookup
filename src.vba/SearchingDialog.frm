Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{234B1E6B-8F17-4037-804C-84EE6518D119}{B6FFC59A-6214-489F-A79C-D9492D1D508E}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub