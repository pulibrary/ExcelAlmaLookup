Attribute VB_Name = "SearchingDialog"
Attribute VB_Base = "0{2F8DA7AB-A6B2-4764-8518-C736BAC1BCE8}{F9D47B0A-A61E-43AD-B5DB-726D343EF98B}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
End Sub