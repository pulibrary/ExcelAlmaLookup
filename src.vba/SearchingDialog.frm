Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{8E8E018C-6195-4236-9DEC-1BEC0516CC1B}{07FE2A58-E54D-4A50-9DDA-04A691A69102}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub