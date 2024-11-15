# Precise Bounding Box, Weight Calculation, and Sketch Creation

## Description
This macro computes precise bounding box values based on the part's geometry tessellation, calculates the gross and real weight based on the assigned material density, and exports these values as custom properties. Additionally, it can draw a 3D sketch representing the bounding box dimensions, allowing users to visualize the bounding box directly in the part.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - The active document must be a SolidWorks part file (*.sldprt).
> - Ensure the part has a material assigned to calculate weight accurately.
> - This macro provides an option to draw a 3D sketch of the bounding box, which is created with precise tessellation-based dimensions.

## Results
> [!NOTE]
> - Calculates and displays the part's bounding box dimensions (length, width, height).
> - Adds custom properties for bounding box dimensions, gross weight, and real weight.
> - Optionally creates a 3D sketch displaying the bounding box dimensions around the part.
> - Outputs bounding box dimensions and weight values in custom properties for easy reference.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

Dim swApp As SldWorks.SldWorks
Dim Part As SldWorks.ModelDoc2
Dim Height As Variant, Width As Variant, Length As Variant
Dim pesoH As Variant, pesoW As Variant, pesoL As Variant
Dim GrossWeight As Variant, ndensity As Double, nMass As Double, gWeight As Double
Dim Corners As Variant, retval As Boolean
Dim ConfigName As String
Dim SwConfig As SldWorks.Configuration
Dim swSketchPt(15) As SldWorks.SketchPoint
Dim swSketchSeg(12) As SldWorks.SketchSegment
Dim Xmax As Variant, Ymax As Variant, Zmax As Variant
Dim Xmin As Variant, Ymin As Variant, Zmin As Variant
Const swDocPart = 1, swDocASSEMBLY = 2

' Function to get maximum of four values
Function GetMax(Val1 As Double, Val2 As Double, Val3 As Double, Val4 As Double) As Double
    GetMax = Application.WorksheetFunction.Max(Val1, Val2, Val3, Val4)
End Function

' Function to get minimum of four values
Function GetMin(Val1 As Double, Val2 As Double, Val3 As Double, Val4 As Double) As Double
    GetMin = Application.WorksheetFunction.Min(Val1, Val2, Val3, Val4)
End Function

Sub ProcessTessTriangles(vTessTriangles As Variant, X_max As Double, X_min As Double, Y_max As Double, Y_min As Double, Z_max As Double, Z_min As Double)
    ' Iterate through tessellation triangles to get bounding box dimensions
    Dim i As Long
    For i = 0 To UBound(vTessTriangles) / (1 * 9) - 1
        X_max = GetMax(vTessTriangles(9 * i + 0), vTessTriangles(9 * i + 3), vTessTriangles(9 * i + 6), X_max)
        X_min = GetMin(vTessTriangles(9 * i + 0), vTessTriangles(9 * i + 3), vTessTriangles(9 * i + 6), X_min)
        Y_max = GetMax(vTessTriangles(9 * i + 1), vTessTriangles(9 * i + 4), vTessTriangles(9 * i + 7), Y_max)
        Y_min = GetMin(vTessTriangles(9 * i + 1), vTessTriangles(9 * i + 4), vTessTriangles(9 * i + 7), Y_min)
        Z_max = GetMax(vTessTriangles(9 * i + 2), vTessTriangles(9 * i + 5), vTessTriangles(9 * i + 8), Z_max)
        Z_min = GetMin(vTessTriangles(9 * i + 2), vTessTriangles(9 * i + 5), vTessTriangles(9 * i + 8), Z_min)
    Next i
End Sub

' Additional functions and main subroutine code continue here with calculations
' and drawing of bounding box in 3D sketch...

Sub main()
    ' Initializes SolidWorks application and active part document
    Set swApp = Application.SldWorks
    Set Part = swApp.ActiveDoc
    If Part Is Nothing Or Part.GetType <> swDocPart Then
        MsgBox "This macro only works on a part document (*.sldprt).", vbCritical
        Exit Sub
    End If

    ' Process part geometry to calculate bounding box dimensions
    ' Set user units and calculate gross weight and real weight
    ' Additional functionality continues as per full code provided...

    ' Display final output
    UserForm1.Show
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).