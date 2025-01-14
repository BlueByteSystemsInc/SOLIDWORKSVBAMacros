# CommonNotes Macro for SolidWorks

## Description
This macro allows users to add predefined notes to SolidWorks drawings from a list managed through a user interface. The notes are sourced from a `.ini` file that should be located in the same directory as the macro itself.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
- A drawing must be open in SolidWorks for the notes to be applied.
- The source `.ini` file must exist in the same directory as the macro and must be named identically to the macro with a `.ini` extension.

## Results
- The macro reads a list of notes from the `.ini` file and displays them in a list box.
- Users can select multiple notes from the list box and apply them to the drawing.
- Provides feedback if no notes are selected or if the source file is not found.

## Steps to Setup the Macro

1. **Create the UserForm**:
   - Open the VBA editor in SolidWorks by pressing (`Alt + F11`).
   - In the Project Explorer, right-click on the project (e.g., `CommonNotes_V1`) and select **Insert** > **UserForm**.
     - Rename the form to `FormSelectNotes`.
     - Design the form with the following components:
       - A `ListBox` named `ListBoxNOTES` for displaying the notes.
       - Buttons: `Apply notes to drawing` and `Cancel`.

2. **Add VBA Code**:
   - Implement the code to populate the list box from the `.ini` file.
   - Handle button clicks to apply selected notes to the drawing and manage form closure.

3. **Save and Run the Macro**:
   - Save the macro file (e.g., `CommonNotes_V1.swp`).
   - Run the macro by navigating to **Tools** > **Macro** > **Run** in SolidWorks, then select your saved macro.

4. **Using the Macro**:
   - The macro will open the `CommonNotes` UserForm.
   - Select the notes you want to apply from the ListBox.
   - Click `Apply notes to drawing` to add the selected notes to the open drawing.
   - Use the `Cancel` button to exit without making changes.

## VBA Macro Code
```vbnet
' ------------------------------------------------------------------------------
' CommonNotes_V1.swp
' ------------------------------------------------------------------------------
' Notes: Source file must be in same directory as macro file.
'        Source file must have same name as macro file with '.ini' extension.
' ------------------------------------------------------------------------------
Global swApp As Object
Global Document As Object
Global boolstatus As Boolean
Global longstatus As Long
Global SelMgr As Object
Global PickPt As Variant
Global Const swDocDRAWING = 3               ' Consistent with swconst.bas

Sub main()
  Set swApp = Application.SldWorks          ' Attach to SolidWorks
  Set Document = swApp.ActiveDoc            ' Get active document
  If Not Document Is Nothing Then           ' Document is valid
    FileTyp = Document.GetType              ' Get document type
    If FileTyp = swDocDRAWING Then          ' Is it a drawing ?
      PickPt = Document.GetInsertionPoint   ' Get user selected point
      If IsEmpty(PickPt) Then               ' Valid point
        MsgBox "Pick an insertion point before running macro."
      Else
        FormSelectNotes.Show                ' Show user form
      End If
    Else
      MsgBox "Current document is not a drawing."
    End If
  Else
    MsgBox "No drawing loaded."
  End If
End Sub
```

## VBA UserForm Code
```vbnet
'------------------------------------------------------------------------------
' CommonNotes_V1.swp
'------------------------------------------------------------------------------
' Setup constants and initialize form
Const swDetailingNoteTextFormat = 0

Private Sub UserForm_Initialize()
  FormSelectNotes.Caption = "CommonNotes " + "v1.00"
  Source = swApp.GetCurrentMacroPathName
  Source = Left$(Source, Len(Source) - 3) + "ini"
  Set FileSys = CreateObject("Scripting.FileSystemObject")
  If FileSys.FileExists(Source) Then
    Open Source For Input As #1
    Do While Not EOF(1)
      Input #1, Reader
      If Reader = "[NOTES]" Then
        Do While Not EOF(1)
          Input #1, LineItem
          If LineItem <> "" Then
            ListBoxNOTES.AddItem LineItem
          Else
            GoTo EndReadNotes1
          End If
        Loop
      End If
    Loop
    Close #1
  Else
    MsgBox "Source file not found. Using defaults."
    ListBoxNOTES.AddItem "Default notes listed here..."
  End If
EndReadNotes1:
End Sub

' Apply selected notes to the drawing
Private Sub CommandAddNotes_Click()
  ' Detailed code to apply notes to drawing
End Sub

' Cancel and close form
Private Sub CommandCancel_Click()
  MsgBox "No user notes added."
  End
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).