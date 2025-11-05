Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{F5B835CE-8200-4E22-8682-1AA9026F1B46}{17A09123-0E1D-4F2A-BA71-845BE01CC535}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub