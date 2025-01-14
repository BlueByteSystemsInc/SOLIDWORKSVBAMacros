# SolidWorks Macro to Remove Toolbox Status from Part

## Description
This macro allows you to remove the Toolbox status from a part document in SolidWorks. Toolbox parts in SolidWorks have specific properties that differentiate them from regular parts. This script modifies the `ToolboxPartType` property to remove the Toolbox designation from the active part document.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - Ensure a part document is open in SolidWorks.
> - The part document must have been originally created as a Toolbox part.

## Results
> [!NOTE]
> - The macro will remove the Toolbox status from the currently active part document.
> - After execution, the part will no longer be recognized as a Toolbox part by SolidWorks.

## Steps to Setup the Macro

### 1. **Prepare SolidWorks**:
   - Open a part document that was originally created as a Toolbox part in SolidWorks.

### 2. **Run the Macro**:
   - Execute the `main` subroutine.
   - Confirm the success message indicating the removal of Toolbox status.

### 3. **Verify Changes**:
   - Check the part properties in SolidWorks to confirm that it no longer has Toolbox attributes.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

' Declare SolidWorks application and document variables
Dim swApp As SldWorks.SldWorks
Dim part As SldWorks.ModelDoc2
Dim modelDocExt As SldWorks.ModelDocExtension

Sub main()
    ' Initialize the SolidWorks application object
    Set swApp = Application.SldWorks

    ' Get the active document in SolidWorks
    Set part = swApp.ActiveDoc

    ' Check if an active document is open
    If part Is Nothing Then
        MsgBox "No active document found. Please open a part document and try again.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Check if the active document is a part
    If part.GetType <> swDocPART Then
        MsgBox "This macro only works on part documents. Please open a part document and try again.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Access the model document extension to modify properties of the part
    Set modelDocExt = part.Extension

    ' Remove the toolbox status by setting ToolboxPartType to 0
    modelDocExt.ToolboxPartType = 0

    ' Notify the user that the operation was successful
    MsgBox "The part is no longer recognized as a Toolbox part.", vbInformation, "Operation Successful"
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).