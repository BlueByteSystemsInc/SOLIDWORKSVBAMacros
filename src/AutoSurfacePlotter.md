# Automatic Surface Plotter in SolidWorks

## Description
This macro allows users to plot functions in Cartesian, cylindrical, or spherical coordinates in SolidWorks. It automatically verifies that the active document is a part or assembly, then opens a user interface for input. This tool is ideal for quickly generating surfaces based on mathematical functions.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - The active document must be a part or assembly.
> - This macro launches a form interface for entering plot parameters (e.g., Cartesian, cylindrical, or spherical function details).

## Results
> [!NOTE]
> - Plots a mathematical function surface in the active SolidWorks part or assembly.
> - If the document is not a part or assembly, the macro will alert the user and exit.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

Public swApp As SldWorks.SldWorks
Public swDoc As SldWorks.ModelDoc2

' Main subroutine
Sub main()
    Err.Clear
    Set swApp = Application.SldWorks
    Set swDoc = swApp.ActiveDoc
    
    ' Check for valid active document (part or assembly)
    If swDoc Is Nothing Then
       swApp.SendMsgToUser "No active part or assembly"
       Exit Sub
    ElseIf swDoc.GetType <> swDocPART And swDoc.GetType <> swDocASSEMBLY Then
       swApp.SendMsgToUser "Active document must be a part or assembly"
       Exit Sub
    End If
    
    ' Show main form for function plotting
    frmMain.Show False
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).