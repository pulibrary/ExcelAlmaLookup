Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{D83556FE-F14F-4164-8DA4-A5F9784A2FDE}{363DD391-AF1D-4173-BD15-1A65BFBFD2D6}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub