# Rotate All Bodies in Active Part Along X-Axis in SolidWorks

## Description
This macro rotates all bodies in the active part document along the X-axis by a specified angle (in degrees), in either a positive or negative direction. It’s ideal for adjusting the orientation of all bodies within a part file for alignment or repositioning needs.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - The active document must be a part document containing at least one body.
> - The macro prompts the user to enter an angle (in degrees) for the rotation.

## Results
> [!NOTE]
> - All bodies in the part are rotated along the X-axis by the specified angle.
> - The macro clears selections and sets the rotation transformation based on the entered angle.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

Sub SelectOrigin(swApp As SldWorks.SldWorks, swModel As SldWorks.ModelDoc2, nSelMark As Long)
    On Error GoTo ErrorHandler

    Dim swFeat As SldWorks.Feature
    Dim bRet As Boolean
    Set swFeat = swModel.FirstFeature

    Do While Not swFeat Is Nothing
        If "OriginProfileFeature" = swFeat.GetTypeName Then
            bRet = swFeat.Select2(True, nSelMark)
            If Not bRet Then
                MsgBox "Failed to select the origin feature.", vbExclamation
                Exit Sub
            End If
            Exit Do
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop

    Exit Sub

ErrorHandler:
    MsgBox "Error selecting origin: " & Err.Description, vbCritical
End Sub

Sub SelectBodies(swApp As SldWorks.SldWorks, swModel As SldWorks.ModelDoc2, vBodyArr As Variant)
    On Error GoTo ErrorHandler

    Dim swSelMgr As SldWorks.SelectionMgr
    Dim swSelData As SldWorks.SelectData
    Dim vBody As Variant
    Dim swBody As SldWorks.Body2
    Dim bRet As Boolean

    If IsEmpty(vBodyArr) Then
        MsgBox "No bodies found in the part.", vbExclamation
        Exit Sub
    End If

    For Each vBody In vBodyArr
        Set swBody = vBody
        Set swSelMgr = swModel.SelectionManager
        Set swSelData = swSelMgr.CreateSelectData
        swSelData.Mark = 1
        bRet = swBody.Select2(True, swSelData)
        If Not bRet Then
            MsgBox "Failed to select body.", vbExclamation
            Exit For
        End If
    Next vBody

    Exit Sub

ErrorHandler:
    MsgBox "Error selecting bodies: " & Err.Description, vbCritical
End Sub

Sub main()
    On Error GoTo ErrorHandler

    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swPart As SldWorks.PartDoc
    Dim vBodyArr As Variant
    Dim swFeatMgr As SldWorks.FeatureManager
    Dim swFeat As SldWorks.Feature
    Dim bRet As Boolean

    Set swApp = Application.SldWorks
    If swApp Is Nothing Then
        MsgBox "SolidWorks application not found.", vbCritical
        Exit Sub
    End If

    Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then
        MsgBox "No active document found. Please open a part document.", vbCritical
        Exit Sub
    End If

    If swModel.GetType <> swDocPART Then
        MsgBox "Active document is not a part. Please open a part document.", vbExclamation
        Exit Sub
    End If

    Set swPart = swModel
    Set swFeatMgr = swModel.FeatureManager

    ' Clear any existing selection
    swModel.ClearSelection2 True

    ' Get all bodies in the part
    vBodyArr = swPart.GetBodies(swAllBodies)
    SelectBodies swApp, swModel, vBodyArr

    ' Select origin for rotation
    SelectOrigin swApp, swModel, 8

    ' Prompt user for rotation angle
    Dim X As Double
    X = InputBox("Enter angle in degrees for rotation along X-axis:", "Rotation Angle")
    If Not IsNumeric(X) Then
        MsgBox "Invalid input. Please enter a numeric value.", vbExclamation
        Exit Sub
    End If
    X = X * 0.0174532925 ' Convert degrees to radians

    ' Rotate bodies in X-axis direction
    Set swFeat = swFeatMgr.InsertMoveCopyBody2(0, 0, 0, 0, 0, 0, 0, 0, 0, X, False, 1)
    If swFeat Is Nothing Then
        MsgBox "Failed to rotate bodies.", vbExclamation
    End If

    ' Clear selection
    swModel.ClearSelection2 True

    Exit Sub

ErrorHandler:
    MsgBox "An unexpected error occurred: " & Err.Description, vbCritical
End Sub

```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).