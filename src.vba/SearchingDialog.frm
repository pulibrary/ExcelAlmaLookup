Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{309717CC-088F-4841-B694-4BDC0B2272DB}{E13810FB-9916-4635-AC8B-705AD6E27827}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub