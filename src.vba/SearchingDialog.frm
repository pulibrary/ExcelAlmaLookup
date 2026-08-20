Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{D370FDD7-FEA8-49FA-9E84-B6C6F7500BD0}{9F5A4C30-7521-4B25-A20E-ADC023089BBB}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub