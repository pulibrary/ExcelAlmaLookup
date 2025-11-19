Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{21164AEE-D731-4987-9266-61F8D9188D83}{0C2684C0-B08D-4644-BF4B-125072871937}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub