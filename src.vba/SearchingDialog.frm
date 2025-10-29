Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{AE2FBCEE-84D5-4B14-ACBC-378F8E03B5E9}{F88D3F80-CAD9-419A-99D3-1D62A2CDE9E4}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub