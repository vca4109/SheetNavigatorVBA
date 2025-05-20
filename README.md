SheetNavigator — A Dynamic Excel Sheet Switcher (VBA UserForm)

The SheetNavigator is a VBA-based navigation tool that dynamically generates clickable buttons for each worksheet in your Excel workbook, allowing you to instantly jump between sheets via a sleek, scrollable UserForm interface.

Components Breakdown:
1) UserForm Code: SheetNavigator. This is the heart of the tool. It:
   > Creates buttons for each worksheet when the form is initialized.
   > Adds these buttons to a scrollable Frame (SheetFrame) within the form.
   > Associates each button with an event handler (from a class module).
   > Dynamically sizes the scroll area of the Frame based on the number of sheets.
2) Private Sub UserForm_Initialize()
   ' Loops through all sheets and creates a clickable button for each.
   End Sub
3) Class Module Code: clsButtonHandler. This class allows each generated button to trigger a unique click event.
   > On button click:
   > Activates the corresponding sheet.
   > Closes the form.

    Private Sub Button_Click()
        Sheets(Button.Tag).Activate
        Unload SheetNavigator
    End Sub

4) Standard Module Code. This module contains a single subroutine to launch the SheetNavigator form.

    Sub OpenSheetNavigator()
      SheetNavigator.Show
    End Sub

How It Works;
1) When you run OpenSheetNavigator, it opens the form.
2) The form scans all worksheets and creates labeled buttons for each.

Clicking a button activates the corresponding sheet and closes the form.
✨ Features
📌 Auto-generated navigation for all sheets
🖱️ One-click access to any sheet
📜 Scrollable interface for workbooks with many sheets
🧩 Built-in Class Module support for individual button click events

Usage: To trigger the navigation interface:
    
    Sub OpenSheetNavigator()
        SheetNavigator.Show
    End Sub

