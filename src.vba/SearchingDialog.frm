Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{7F7E9CFE-7E64-40BC-B586-626CC8DD0F1D}{A46ABC8E-188C-4C6A-A519-B4FB44B1F56B}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub