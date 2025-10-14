Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{F2481787-B4FF-4292-B667-4EE96127D92F}{3926FE0B-D7AF-4E4B-A641-7BD9127C89E7}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub