Option Explicit

Dim btnHandlers() As New clsButtonHandler  ' Array for handling button clicks

Private Sub SheetFrame_Click()

End Sub

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim btn As MSForms.CommandButton
    Dim i As Integer
    Dim topPos As Integer
    
    ' Clear existing buttons
    Me.SheetFrame.Controls.Clear
    
    ' Resize the handlers array
    ReDim btnHandlers(1 To ThisWorkbook.Sheets.Count)
    
    topPos = 10 ' Initial position for first button
    
    ' Loop through each sheet and create a button
    For i = 1 To ThisWorkbook.Sheets.Count
        Set ws = ThisWorkbook.Sheets(i)
        
        ' Create a button inside the Frame
        Set btn = Me.SheetFrame.Controls.Add("Forms.CommandButton.1", "btn" & i, True)
        With btn
            .Caption = ws.Name  ' Set button text to sheet name
            .Left = 5
            .Top = topPos
            .Width = 190
            .Height = 25
            .Tag = ws.Name  ' Store sheet name
            .Font.Size = 10
        End With
        
        ' Assign the event handler from our class module
        Set btnHandlers(i) = New clsButtonHandler
        Set btnHandlers(i).Button = btn
        
        topPos = topPos + 30  ' Move the next button lower
    Next i
    
    ' Adjust Frame's ScrollHeight dynamically
    Me.SheetFrame.ScrollHeight = topPos + 10
End Sub

