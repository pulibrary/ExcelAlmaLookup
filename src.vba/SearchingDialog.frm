Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{FEC24736-80DC-428B-B3F3-1B4A6FA9D5DE}{CFA50007-7915-4890-96A9-8437E68B7903}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub