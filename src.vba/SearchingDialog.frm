Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{B30AED2B-1ABB-44C9-BE6B-FD313EDBE9BB}{1D49CE73-BE87-4C2B-BB19-F4617CA05B5B}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub