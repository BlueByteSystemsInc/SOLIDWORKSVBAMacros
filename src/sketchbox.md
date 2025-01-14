# SketchBox Macro for SolidWorks

## Description
This macro creates a rectangle centered about the origin on a preselected plane in SolidWorks. It automates the process of sketching a rectangle and ensures that dimensions are appropriately assigned.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer  
- **Operating System**: Windows 7 or later  

## Pre-Conditions
> [!NOTE]
> - A plane or face must be preselected before running the macro.
> - SolidWorks must have an active document.

## Results
> [!NOTE]
> - A centered rectangle will be drawn on the selected plane or face.
> - A construction line connecting opposite corners of the rectangle will be added and constrained to the origin.
> - Dimensions for the width and height of the rectangle will be created.

## Steps to Setup the Macro

### 1. **Select a Plane or Face**:
   - In SolidWorks, preselect the plane or face where the rectangle will be sketched.

### 2. **Run the Macro**:
   - Execute the macro. The rectangle and associated features will be created automatically.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

Dim swApp As Object                          ' SolidWorks application object
Dim Part As Object                           ' Active document object
Dim SelMgr As Object                         ' Selection manager for the active document
Dim boolstatus As Boolean                    ' Boolean status for operations
Dim longstatus As Long, longwarnings As Long ' Long status for warnings/errors
Dim Feature As Object                        ' Feature object
Dim CurSelCount As Long                      ' Count of selected items

Sub main()

    ' Initialize SolidWorks application and active document
    Set swApp = Application.SldWorks
    Set Part = swApp.ActiveDoc

    ' Ensure there is an active document
    If Part Is Nothing Then
        MsgBox "No active document found. Please open a part or assembly and try again.", vbCritical, "Error"
        Exit Sub
    End If

    ' Initialize the selection manager
    Set SelMgr = Part.SelectionManager

    ' Disable input dimensions on creation
    swApp.SetUserPreferenceToggle swInputDimValOnCreate, False

    ' Check if a plane or face is preselected
    CurSelCount = SelMgr.GetSelectedObjectCount
    If CurSelCount = 0 Then
        MsgBox "Please preselect a plane or face before running the macro.", vbExclamation, "No Selection"
        Exit Sub
    End If

    ' Insert a new sketch on the selected plane or face
    boolstatus = Part.Extension.SelectByID2("", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
    Part.InsertSketch2 True
    Part.ClearSelection2 True

    ' Create a rectangle centered about the origin
    Part.SketchRectangle -0.037, 0.028, 0, 0.015, -0.019, 0, True

    ' Clear selection and add a diagonal construction line
    Part.ClearSelection2 True
    Dim Line As Object
    Set Line = Part.CreateLine2(-0.037, -0.019, 0, 0.015, 0.028, 0)
    Line.ConstructionGeometry = True

    ' Add midpoint constraints to ensure the rectangle is centered
    boolstatus = Part.Extension.SelectByID2("Point1@Origin", "EXTSKETCHPOINT", 0, 0, 0, True, 0, Nothing, 0)
    Part.SketchAddConstraints "sgATMIDDLE"
    Part.ClearSelection2 True

    ' Add dimensions to the rectangle
    boolstatus = Part.Extension.SelectByID2("Line1", "SKETCHSEGMENT", -0.001, 0.027, 0, False, 0, Nothing, 0)
    Dim Annotation As Object
    Set Annotation = Part.AddDimension2(-0.0004, 0.045, 0) ' Horizontal dimension
    Part.ClearSelection2 True

    boolstatus = Part.Extension.SelectByID2("Line2", "SKETCHSEGMENT", -0.030, 0.001, 0, False, 0, Nothing, 0)
    Set Annotation = Part.AddDimension2(-0.061, -0.001, 0) ' Vertical dimension
    Part.ClearSelection2 True

    ' Re-enable input dimensions on creation
    swApp.SetUserPreferenceToggle swInputDimValOnCreate, True

    ' Inform the user that the macro is complete
    MsgBox "Rectangle sketch created successfully.", vbInformation, "Success"

End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).