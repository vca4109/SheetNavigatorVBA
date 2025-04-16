Option Explicit

Public WithEvents Button As MSForms.CommandButton

Private Sub Button_Click()
    Dim sheetName As String
    sheetName = Button.Tag ' Get the stored sheet name
    
    On Error Resume Next
    Sheets(sheetName).Activate  ' Switch to the selected sheet
    Unload SheetNavigator  ' Close the form after clicking
    On Error GoTo 0
End Sub

