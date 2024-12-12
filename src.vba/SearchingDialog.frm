Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{A1570804-E3B2-4B47-8C3D-AD5FF46DB5CD}{46184C9C-3763-4DF0-8C81-8D7BFD0DB5AF}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub