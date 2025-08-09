# Precise Bounding Box and Sketch Creation
<video src="../images/TessBox-EN-MM-INCH.mkv" autoplay muted controls style="width: 100%; border-radius: 12px;"></video>


## Description
This macro computes precise bounding box values based on the part's geometry. Additionally, it can draw a 3D sketch representing the bounding box dimensions, allowing users to visualize the bounding box directly in the part.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - The active document must be a SolidWorks part file (*.sldprt).
## Results
> [!NOTE]
> - Calculates and displays the part's bounding box dimensions (length, width, height).
> - Adds custom properties for bounding box dimensions.
> - Creates a 3D sketch displaying the bounding box dimensions around the part.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).


Option Explicit

Dim swApp As SldWorks.SldWorks

Sub Main()
    ' Get SOLIDWORKS application
    Set swApp = Application.SldWorks
    
    ' Get active part document
    Dim swDoc As SldWorks.partDoc
    Set swDoc = swApp.ActiveDoc
    
    If Not swDoc Is Nothing Then
        ' Get precise bounding box using extreme points
        Dim boundingBox As Variant
        boundingBox = GetBoundingBox(swDoc)
        
        ' Draw 3D sketch of bounding box
        CreateBoundingBoxSketch swDoc, boundingBox
        
        ' Calculate bounding box dimensions
        Dim boxWidth As Double
        Dim boxHeight As Double
        Dim boxDepth As Double
        
        boxWidth = CDbl(boundingBox(3)) - CDbl(boundingBox(0))
        boxHeight = CDbl(boundingBox(4)) - CDbl(boundingBox(1))
        boxDepth = CDbl(boundingBox(5)) - CDbl(boundingBox(2))
        
        ' Update custom properties
        UpdateCustomProperties swDoc, boxWidth, boxHeight, boxDepth
  
        
    Else
        Debug.Print "Error: No active part document."
    End If
End Sub

' Function to Get Bounding Box Using Extreme Points
Function GetBoundingBox(partDoc As SldWorks.partDoc) As Variant
    Dim boundingData(5) As Double
    Dim solidBodies As Variant
    solidBodies = partDoc.GetBodies2(swBodyType_e.swSolidBody, True)
    
    Dim minX As Double, minY As Double, minZ As Double
    Dim maxX As Double, maxY As Double, maxZ As Double
    
    If Not IsEmpty(solidBodies) Then
        Dim i As Integer
        For i = 0 To UBound(solidBodies)
            Dim bodyObj As SldWorks.Body2
            Set bodyObj = solidBodies(i)
            
            Dim coordX As Double, coordY As Double, coordZ As Double
            
            ' Get extreme points
            bodyObj.GetExtremePoint 1, 0, 0, coordX, coordY, coordZ: If i = 0 Or coordX > maxX Then maxX = coordX
            bodyObj.GetExtremePoint -1, 0, 0, coordX, coordY, coordZ: If i = 0 Or coordX < minX Then minX = coordX
            bodyObj.GetExtremePoint 0, 1, 0, coordX, coordY, coordZ: If i = 0 Or coordY > maxY Then maxY = coordY
            bodyObj.GetExtremePoint 0, -1, 0, coordX, coordY, coordZ: If i = 0 Or coordY < minY Then minY = coordY
            bodyObj.GetExtremePoint 0, 0, 1, coordX, coordY, coordZ: If i = 0 Or coordZ > maxZ Then maxZ = coordZ
            bodyObj.GetExtremePoint 0, 0, -1, coordX, coordY, coordZ: If i = 0 Or coordZ < minZ Then minZ = coordZ
        Next
    End If
    
    ' Store bounding box coordinates
    boundingData(0) = minX: boundingData(1) = minY: boundingData(2) = minZ
    boundingData(3) = maxX: boundingData(4) = maxY: boundingData(5) = maxZ
    
    GetBoundingBox = boundingData
End Function

' Subroutine to Draw 3D Sketch Bounding Box
Sub CreateBoundingBoxSketch(modelDoc As SldWorks.ModelDoc2, boundingBox As Variant)
    Dim sketchMgr As SldWorks.SketchManager
    Dim minX As Double, minY As Double, minZ As Double
    Dim maxX As Double, maxY As Double, maxZ As Double
    
    ' Extract bounding box coordinates
    minX = CDbl(boundingBox(0)): minY = CDbl(boundingBox(1)): minZ = CDbl(boundingBox(2))
    maxX = CDbl(boundingBox(3)): maxY = CDbl(boundingBox(4)): maxZ = CDbl(boundingBox(5))
    
    ' Start 3D sketch
    Set sketchMgr = modelDoc.SketchManager
    sketchMgr.Insert3DSketch True
    sketchMgr.AddToDB = True
    
    ' Draw bounding box edges
    Create3DSketchLine sketchMgr, maxX, minY, minZ, maxX, minY, maxZ
    Create3DSketchLine sketchMgr, maxX, minY, maxZ, minX, minY, maxZ
    Create3DSketchLine sketchMgr, minX, minY, maxZ, minX, minY, minZ
    Create3DSketchLine sketchMgr, minX, minY, minZ, maxX, minY, minZ

    Create3DSketchLine sketchMgr, maxX, maxY, minZ, maxX, maxY, maxZ
    Create3DSketchLine sketchMgr, maxX, maxY, maxZ, minX, maxY, maxZ
    Create3DSketchLine sketchMgr, minX, maxY, maxZ, minX, maxY, minZ
    Create3DSketchLine sketchMgr, minX, maxY, minZ, maxX, maxY, minZ
    
    Create3DSketchLine sketchMgr, minX, minY, minZ, minX, maxY, minZ
    Create3DSketchLine sketchMgr, minX, minY, maxZ, minX, maxY, maxZ
    Create3DSketchLine sketchMgr, maxX, minY, minZ, maxX, maxY, minZ
    Create3DSketchLine sketchMgr, maxX, minY, maxZ, maxX, maxY, maxZ
    
    ' Finish 3D sketch
    sketchMgr.AddToDB = False
    sketchMgr.Insert3DSketch True
    
    ' Update Model
    modelDoc.ForceRebuild3 True
    modelDoc.GraphicsRedraw2
End Sub

' Helper Function to Create a 3D Sketch Line
Sub Create3DSketchLine(sketchMgr As SldWorks.SketchManager, x1 As Double, y1 As Double, z1 As Double, x2 As Double, y2 As Double, z2 As Double)
    sketchMgr.CreateLine x1, y1, z1, x2, y2, z2
End Sub

' Subroutine to Update Custom Properties
Sub UpdateCustomProperties(modelDoc As SldWorks.ModelDoc2, width As Double, height As Double, depth As Double)
    Dim customPropMgr As SldWorks.CustomPropertyManager
    Set customPropMgr = modelDoc.Extension.CustomPropertyManager("")
    
    ' Convert dimensions to string format for properties
    Dim widthStr As String
    Dim heightStr As String
    Dim depthStr As String
    
    widthStr = Format(width * 1000, "0.000") ' Convert to mm
    heightStr = Format(height * 1000, "0.000")
    depthStr = Format(depth * 1000, "0.000")
    
    ' Set or update custom properties
    customPropMgr.Add3 "BoundingBoxWidth", swCustomInfoText, widthStr & " mm", swCustomPropertyDeleteAndAdd
    customPropMgr.Add3 "BoundingBoxHeight", swCustomInfoText, heightStr & " mm", swCustomPropertyDeleteAndAdd
    customPropMgr.Add3 "BoundingBoxDepth", swCustomInfoText, depthStr & " mm", swCustomPropertyDeleteAndAdd
End Sub


```
You can download the macro from [here](../images/TessBox-EN-MM-INCH.swp)
## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).