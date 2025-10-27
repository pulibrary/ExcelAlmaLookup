Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{4F592199-0144-49E2-9B39-30F7E0E99758}{99BF3CD2-7866-4CDA-B1C7-B1B2C4E1447F}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub