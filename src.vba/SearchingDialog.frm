Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{B7E4CC54-0955-4521-A514-9ED72FFAA33E}{FD6E63F9-8148-46CF-B09F-8D175A9264C1}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub