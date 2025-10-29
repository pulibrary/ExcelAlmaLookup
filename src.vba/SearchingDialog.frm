Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{DE36E20D-63A5-4286-A1CB-B023ABEF9112}{0DDBF805-AE31-47BE-B863-D53C6ECE4DC9}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub