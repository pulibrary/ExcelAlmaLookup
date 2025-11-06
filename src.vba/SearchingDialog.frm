Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{2E816CDF-03C7-46BB-B41E-C5818C2521ED}{89EA1BE2-5681-4AD6-AAA6-1659D910A7B0}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub