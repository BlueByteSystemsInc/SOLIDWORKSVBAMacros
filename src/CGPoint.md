# Create Center of Gravity Point in SolidWorks

## Description
This macro creates a 3D sketch point at the Center of Gravity (CoG) of the active part or assembly document in SolidWorks. It can be used to quickly identify the center of mass location within a part or assembly for analysis and design purposes.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - The active document must be a part or assembly file.
> - The part or assembly must contain valid geometry to calculate the center of gravity.
> - Ensure the part or assembly is open and active before running the macro.

## Results
> [!NOTE]
> - A 3D sketch will be created with a point located at the Center of Gravity.
> - The new sketch will be named "CenterOfGravity" in the feature tree.
> - An error message will be displayed if there is no geometry to process or if the document type is not valid.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

Dim swApp As Object                 ' SolidWorks application object
Dim Part As Object                  ' Active document object (part or assembly)
Dim boolstatus As Boolean           ' Boolean status variable
Dim longstatus As Long              ' Long status variable for capturing operation results
Dim Annotation As Object            ' Annotation object for any annotations added (not used here)
Dim Gtol As Object                  ' Geometric tolerance object (not used here)
Dim DatumTag As Object              ' Datum tag object (not used here)
Dim FeatureData As Object           ' Feature data object for manipulating feature details (not used here)
Dim Feature As Object               ' Feature object for creating/manipulating features (not used here)
Dim Component As Object             ' Component object for assemblies (not used here)

' Main subroutine to create the Center of Gravity point in a 3D sketch
Sub main()
    Dim mp As Variant                ' Array to hold the mass properties (center of gravity coordinates)
    Dim PlaneObj As Object           ' Plane object (not used here)
    Dim PlaneName As String          ' Name of the plane (not used here)
    Dim SketchObj As Object          ' Sketch object for creating the 3D sketch (not used here)
    Dim Version As String            ' SolidWorks version (not used here)

    ' Error handling block to capture unexpected issues
    On Error GoTo errhandlr

    ' Initialize SolidWorks application
    Set swApp = Application.SldWorks
    

    ' Check if SolidWorks application is available
    If swApp Is Nothing Then
        MsgBox "SolidWorks application not found. Please ensure SolidWorks is installed and running.", vbCritical, "SolidWorks Not Found"
        Exit Sub
    End If

    ' Get the currently active document
    Set Part = swApp.ActiveDoc

    ' Check if there is an active document open in SolidWorks
    If Part Is Nothing Then
        MsgBox "No active document found. Please open a part or assembly and try again.", vbCritical, "No Active Document"
        Exit Sub
    End If

    ' Check if the active document is a drawing (GetType = 3 corresponds to drawing)
    If Part.GetType = 3 Then
        MsgBox "This macro only works on parts or assemblies. Please open a part or assembly and try again.", vbCritical, "Invalid Document Type"
        Exit Sub
    End If

    ' Enable adding objects directly to the database without showing in the UI
    Part.SetAddToDB True

    ' Get the mass properties of the active part or assembly
    ' mp array holds center of mass coordinates (mp(0) = X, mp(1) = Y, mp(2) = Z)
    mp = Part.GetMassProperties

    ' Check if mass properties are valid (in case the part has no geometry)
    If Not IsArray(mp) Or UBound(mp) < 2 Then
        MsgBox "No geometry found in the part or assembly. Cannot calculate center of gravity.", vbCritical, "Invalid Geometry"
        Exit Sub
    End If

    ' Insert a new 3D sketch
    Part.Insert3DSketch

    ' Create a point at the center of gravity coordinates
    Part.CreatePoint2 mp(0), mp(1), mp(2)

    ' Exit the sketch mode
    Part.InsertSketch

    ' Rename the newly created feature to "CenterOfGravity" in the feature tree
    Part.FeatureByPositionReverse(0).Name = "CenterOfGravity"

    ' Successfully exit the subroutine
    Exit Sub

' Error handling block
errhandlr:
    MsgBox "An error occurred. No valid geometry found to process.", vbCritical, "Error"
    Exit Sub

End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).

