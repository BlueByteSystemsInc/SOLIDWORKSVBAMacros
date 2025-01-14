# Instance Count Increment Macro

## Description
This macro automates the incrementation or decrementation of numbering in selected dimensions and notes within a SolidWorks document. It's particularly useful for rapidly updating numbering sequences in dimensions or notes for engineering documents.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer  
- **Operating System**: Windows 7 or later  

## Pre-Conditions
> [!NOTE]
> - A part document must be currently open in SolidWorks.
> - At least one dimension or note must be selected.

## Results
> [!NOTE]
> - Updates the numbering in the selected dimensions or notes by adding or subtracting one, based on the selected macro button (`+1` or `-1`).
> - Allows control over whether the prefix or suffix of the dimension/note text is altered.

## Steps to Setup the Macro

### 1. **Configure Macro Settings**:
   - Set the `IncStr` constant to define the increment notation (`X`, `#`, etc.).
   - Set the `IncPrefix` boolean to `True` to increment the prefix of the text or `False` for the suffix.

### 2. **Running the Macro**:
   - Execute the `Plus1` subroutine to increase the numbering by one.
   - Execute the `Minus1` subroutine to decrease the numbering by one.

### 3. **Macro Execution**:
   - The macro checks the selection type and applies the increment or decrement to each selected object.
   - If the selected object is a dimension or hole callout, it extracts and updates the prefix or suffix based on the `IncPrefix` setting.
   - If the selection is a note, it directly updates the text of the note.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

' Constants for increment notation
Const IncStr As String = "X" ' The string appended to incremented values
Const IncPrefix As Boolean = True ' True = change prefix, False = change suffix

' SolidWorks application and document objects
Dim swApp As SldWorks.SldWorks        ' SolidWorks application object
Dim Part As SldWorks.ModelDoc2        ' Active document object
Dim SelMgr As SldWorks.SelectionMgr   ' Selection manager for managing selected entities
Dim i As Integer                      ' Loop counter
Dim selCount As Integer               ' Count of selected objects
Dim selType As Long                   ' Type of the selected object
Dim swDispDim As SldWorks.DisplayDimension ' Display dimension object
Dim swDim As SldWorks.Dimension       ' Dimension object
Dim Note As SldWorks.Note             ' Note object
Dim NumValue As Integer               ' Numeric value extracted from text
Dim OldText As String                 ' Original text from the selected entity
Dim NewText As String                 ' New text after increment/decrement

' Macro to increment values by +1
Sub Plus1()
    Increment 1
End Sub

' Macro to decrement values by -1
Sub Minus1()
    Increment -1
End Sub

' Core subroutine to increment or decrement selected text or dimensions
Sub Increment(x As Integer)
    ' Initialize SolidWorks application and active document
    Set swApp = Application.SldWorks
    Set Part = swApp.ActiveDoc
    Set SelMgr = Part.SelectionManager()
    
    ' Get the count of selected objects
    selCount = SelMgr.GetSelectedObjectCount()
    
    ' Loop through each selected object
    For i = 1 To selCount
        ' Get the type of the selected object
        selType = SelMgr.GetSelectedObjectType3(i, -1)
        Select Case selType
        Case swSelDIMENSIONS
            ' Handle selected dimensions
            Set swDispDim = SelMgr.GetSelectedObject6(i, -1) ' Get the selected dimension
            OldText = swDispDim.GetText(swDimensionTextPrefix) ' Get the current prefix text
            NumValue = Val(OldText) + x ' Increment or decrement the numeric value
            NewText = Format(NumValue) & IncStr ' Create the new text with increment notation
            swDispDim.SetText swDimensionTextPrefix, NewText ' Update the dimension text
            
        Case swSelNOTES
            ' Handle selected notes
            Set Note = SelMgr.GetSelectedObject6(i, -1) ' Get the selected note
            OldText = Note.GetText ' Get the current note text
            NumValue = Val(OldText) + x ' Increment or decrement the numeric value
            NewText = Format(NumValue) & IncStr ' Create the new text with increment notation
            Note.SetText NewText ' Update the note text
        End Select
    Next
    
    ' Refresh the graphics to reflect the changes
    Part.GraphicsRedraw2
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).