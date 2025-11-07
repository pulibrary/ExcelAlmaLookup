Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{AB570172-1F2A-4791-94B6-713AD8AE0B10}{71879888-731B-456C-8CCD-50A97FDC26E9}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub