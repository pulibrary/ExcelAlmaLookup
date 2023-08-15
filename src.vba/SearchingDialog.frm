Private Sub CancelButton_Click()
    Catalog.bTerminateLoop = True
    SearchingDialog.Hide
    LookupDialog.Show
End Sub