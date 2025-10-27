Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{BCCDDFCE-4089-45C6-BC86-DA9DE42ED227}{C9C6AFC0-A158-4EE7-A291-278F4C254701}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub