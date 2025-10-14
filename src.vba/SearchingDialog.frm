Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{E1D99AD2-A3B1-4E77-A88E-2E8EC17636A2}{87CC196E-EFC4-4638-8BB5-0FDD91D48706}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub