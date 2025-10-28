Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{25401421-98A3-4F01-A581-66077370AFF8}{9F043030-2C19-4A1B-8F77-9284C9AEF536}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub