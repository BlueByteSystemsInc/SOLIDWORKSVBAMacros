# ISO & Shaded With Edges Macro in SolidWorks

## Description
This macro sets the active view display mode to **Shaded with Edges** in SolidWorks, changes the view orientation to **Isometric**, zooms to fit, saves the part silently, and then closes it. This tool is useful for quickly adjusting and saving a part’s display settings in a standardized format.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - An active part document must be open with at least one body.
> - The macro should be executed in SolidWorks with the part open.

## Results
> [!NOTE]
> - Sets the view mode to **Shaded with Edges**.
> - Changes the view orientation to **Isometric**.
> - Zooms to fit the part in the window.
> - Saves the part silently and closes it.


## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

Public Enum swViewDisplayMode_e
    swViewDisplayMode_Wireframe = 1
    swViewDisplayMode_HiddenLinesRemoved = 2
    swViewDisplayMode_HiddenLinesGrayed = 3
    swViewDisplayMode_Shaded = 4
    swViewDisplayMode_ShadedWithEdges = 5   ' Only valid for a part
End Enum

Sub main()
    Const nNewDispMode As Long = swViewDisplayMode_e.swViewDisplayMode_ShadedWithEdges

    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swModView As SldWorks.ModelView
    Dim bRet As Boolean
    Dim swError As Long
    Dim swWarnings As Long

    On Error GoTo ErrorHandler ' Set up error handling

    ' Initialize SolidWorks application and model
    Set swApp = Application.SldWorks
    If swApp Is Nothing Then
        MsgBox "SolidWorks application not found.", vbCritical
        Exit Sub
    End If

    Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then
        MsgBox "No active document found. Please open a document and try again.", vbCritical
        Exit Sub
    End If

    Set swModView = swModel.ActiveView
    If swModView Is Nothing Then
        MsgBox "Unable to access model view.", vbCritical
        Exit Sub
    End If

    ' Set display mode to Shaded with Edges
    swModView.DisplayMode = nNewDispMode
    Debug.Assert nNewDispMode = swModView.DisplayMode

    ' Change view to Isometric and zoom to fit
    bRet = swModel.ShowNamedView2("*Isometric", 7)
    If Not bRet Then
        MsgBox "Failed to set view orientation to Isometric.", vbExclamation
    End If

    swModel.ViewZoomtofit2

    ' Force rebuild to apply changes
    swModel.ForceRebuild3 False

    ' Save the document silently and close it
    bRet = swModel.Save3(swSaveAsOptions_e.swSaveAsOptions_Silent, swError, swWarnings)
    If Not bRet Then
        MsgBox "Error saving document. Error code: " & swError & ", Warnings: " & swWarnings, vbExclamation
    End If

    swApp.CloseDoc swModel.GetPathName

    Exit Sub

ErrorHandler:
    MsgBox "An unexpected error occurred: " & Err.Description, vbCritical
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).