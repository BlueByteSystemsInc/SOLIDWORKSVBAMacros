# Hole Detection Macro

## Description
This macro identifies and processes circular holes on a selected face in a SolidWorks part or assembly document. It calculates the diameter and material thickness of the hole(s), selects and counts the holes, and displays the results. If no valid holes are found, the user is notified. Additionally, the macro measures the time taken to perform the operation.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - The active document must be a part or assembly.
> - Only one face must be selected in the part or assembly.
> - The face must contain circular holes (elliptical holes are ignored).

## Results
> [!NOTE]
> - Displays the diameter of the first hole found, the material thickness, and the total number of holes on the selected face.
> - If no holes are found or if an invalid selection is made, appropriate warnings are displayed.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

' Declare public variables to store hole diameters and thicknesses
Public Diameter(2000) As Single
Public Thickness(2000) As Single

Dim swModel As SldWorks.ModelDoc2

' --------------------------------------------------------------------------
' Function to get the normal vector of a face at the midpoint of the co-edge
' --------------------------------------------------------------------------
Function GetFaceNormalAtMidCoEdge(swCoEdge As SldWorks.CoEdge) As Variant
    Dim swFace As SldWorks.Face2
    Dim swSurface As SldWorks.Surface
    Dim swLoop As SldWorks.Loop2
    Dim varParams As Variant
    Dim varPoint As Variant
    Dim dblMidParam As Double
    Dim dblNormal(2) As Double
    Dim bFaceSenseReversed As Boolean

    varParams = swCoEdge.GetCurveParams

    ' Calculate the midpoint of the co-edge curve
    If varParams(6) > varParams(7) Then
        dblMidParam = (varParams(6) - varParams(7)) / 2 + varParams(7)
    Else
        dblMidParam = (varParams(7) - varParams(6)) / 2 + varParams(6)
    End If
    varPoint = swCoEdge.Evaluate(dblMidParam)

    ' Get the face and surface corresponding to the co-edge
    Set swLoop = swCoEdge.GetLoop
    Set swFace = swLoop.GetFace
    Set swSurface = swFace.GetSurface
    bFaceSenseReversed = swFace.FaceInSurfaceSense
    varParams = swSurface.EvaluateAtPoint(varPoint(0), varPoint(1), varPoint(2))

    ' Adjust the normal vector based on the face's sense
    If bFaceSenseReversed Then
        dblNormal(0) = -varParams(0)
        dblNormal(1) = -varParams(1)
        dblNormal(2) = -varParams(2)
    Else
        dblNormal(0) = varParams(0)
        dblNormal(1) = varParams(1)
        dblNormal(2) = varParams(2)
    End If

    GetFaceNormalAtMidCoEdge = dblNormal
End Function

' --------------------------------------------------------------------------
' Function to get the tangent vector at the midpoint of a co-edge
' --------------------------------------------------------------------------
Function GetTangentAtMidCoEdge(swCoEdge As SldWorks.CoEdge) As Variant
    Dim varParams As Variant
    Dim dblMidParam As Double
    Dim dblTangent(2) As Double

    varParams = swCoEdge.GetCurveParams

    ' Calculate the midpoint of the co-edge curve
    If varParams(6) > varParams(7) Then
        dblMidParam = (varParams(6) - varParams(7)) / 2 + varParams(7)
    Else
        dblMidParam = (varParams(7) - varParams(6)) / 2 + varParams(6)
    End If

    varParams = swCoEdge.Evaluate(dblMidParam)

    ' Retrieve the tangent vector
    dblTangent(0) = varParams(3)
    dblTangent(1) = varParams(4)
    dblTangent(2) = varParams(5)
    GetTangentAtMidCoEdge = dblTangent
End Function

' --------------------------------------------------------------------------
' Function to get the cross product of two vectors
' --------------------------------------------------------------------------
Function GetCrossProduct(varVec1 As Variant, varVec2 As Variant) As Variant
    Dim dblCross(2) As Double
    dblCross(0) = varVec1(1) * varVec2(2) - varVec1(2) * varVec2(1)
    dblCross(1) = varVec1(2) * varVec2(0) - varVec1(0) * varVec2(2)
    dblCross(2) = varVec1(0) * varVec2(1) - varVec1(1) * varVec2(0)
    GetCrossProduct = dblCross
End Function

' --------------------------------------------------------------------------
' Function to check if two vectors are equal within a tolerance
' --------------------------------------------------------------------------
Function VectorsAreEqual(varVec1 As Variant, varVec2 As Variant) As Boolean
    Dim dblMag As Double
    Dim dblDot As Double
    Dim dblUnit1(2) As Double
    Dim dblUnit2(2) As Double

    dblMag = (varVec1(0) * varVec1(0) + varVec1(1) * varVec1(1) + varVec1(2) * varVec1(2)) ^ 0.5
    dblUnit1(0) = varVec1(0) / dblMag: dblUnit1(1) = varVec1(1) / dblMag: dblUnit1(2) = varVec1(2) / dblMag
    dblMag = (varVec2(0) * varVec2(0) + varVec2(1) * varVec2(1) + varVec2(2) * varVec2(2)) ^ 0.5
    dblUnit2(0) = varVec2(0) / dblMag: dblUnit2(1) = varVec2(1) / dblMag: dblUnit2(2) = varVec2(2) / dblMag
    dblDot = dblUnit1(0) * dblUnit2(0) + dblUnit1(1) * dblUnit2(1) + dblUnit1(2) * dblUnit2(2)
    dblDot = Abs(dblDot - 1#)

    ' Compare within a tolerance
    If dblDot < 0.0000000001 Then '1.0e-10
        VectorsAreEqual = True
    Else
        VectorsAreEqual = False
    End If
End Function

' --------------------------------------------------------------------------
' Function to select hole edges on a face and calculate hole dimensions
' --------------------------------------------------------------------------
Sub SelectHoleEdges(swFace As SldWorks.Face2, swSelData As SldWorks.SelectData)
    Dim swThisLoop As SldWorks.Loop2
    Dim swThisCoEdge As SldWorks.CoEdge
    Dim swPartnerCoEdge As SldWorks.CoEdge
    Dim varThisNormal As Variant
    Dim varPartnerNormal As Variant
    Dim varCrossProduct As Variant
    Dim varTangent As Variant
    Dim vEdgeArr As Variant
    Dim swEdge As SldWorks.Edge
    Dim swCurve As SldWorks.Curve
    Dim vCurveParam As Variant
    Dim i As Integer
    Dim index As Integer
    Dim bRet As Boolean
    Dim pi As Single

    pi = 3.14159265359
    index = 0
    
    ' Get the first loop in the face
    Set swThisLoop = swFace.GetFirstLoop

    Do While Not swThisLoop Is Nothing
        ' Hole is inner loop and has only one edge (circular or elliptical)
        If swThisLoop.IsOuter = False And 1 = swThisLoop.GetEdgeCount Then
            Set swThisCoEdge = swThisLoop.GetFirstCoEdge
            Set swPartnerCoEdge = swThisCoEdge.GetPartner

            varThisNormal = GetFaceNormalAtMidCoEdge(swThisCoEdge)
            varPartnerNormal = GetFaceNormalAtMidCoEdge(swPartnerCoEdge)

            ' Check if the normals of the faces are not equal
            If Not VectorsAreEqual(varThisNormal, varPartnerNormal) Then
                ' Calculate cross product and tangent vector
                varCrossProduct = GetCrossProduct(varThisNormal, varPartnerNormal)
                varTangent = GetTangentAtMidCoEdge(swThisCoEdge)

                ' If cross product and tangent vector are equal, process the hole
                If VectorsAreEqual(varCrossProduct, varTangent) Then
                    vEdgeArr = swThisLoop.GetEdges
                    Set swEdge = vEdgeArr(0)
                    Set swCurve = swEdge.GetCurve
                    vCurveParam = swEdge.GetCurveParams2

                    ' Ignore elliptical holes, only process circular ones
                    If swCurve.IsCircle Then
                        ' Select the edge and calculate diameter
                        bRet = swEdge.Select4(True, swSelData)
                        Diameter(index) = Round(swCurve.GetLength2(vCurveParam(6), vCurveParam(7)) * 1000# / pi, 2)
                    End If
                End If
            End If
        End If
        Set swThisLoop = swThisLoop.GetNext ' Move to next loop
        index = index + 1
    Loop
End Sub

' --------------------------------------------------------------------------
' Main subroutine to process the selected face and count holes
' --------------------------------------------------------------------------
Sub main()
    Dim swApp As SldWorks.SldWorks
    Dim swSelMgr As SldWorks.SelectionMgr
    Dim swSelData As SldWorks.SelectData
    Dim swFace As SldWorks.Face2
    Dim objCount As Long
    Dim TimeStart As Single
    Dim TimeEnd As Single
    
    ' Initialize SolidWorks application and check active document
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    ' Error handling: Check if there is an active document and it's not a drawing
    If swModel Is Nothing Then
        MsgBox "No document is opened!", vbExclamation, "Warning !"
        Exit Sub
    End If
    If swModel.GetType = swDocDRAWING Then
        MsgBox "This is not a part or assembly document!", vbExclamation, "Warning !"
        Exit Sub
    End If

    ' Get the selection manager and ensure only one face is selected
    Set swSelMgr = swModel.SelectionManager
    If swSelMgr.GetSelectedObjectCount > 1 Then
        MsgBox "You can only select one face!", vbExclamation, "Warning !"
        Exit Sub
    End If
    If swSelMgr.GetSelectedObjectCount < 1 Then
        MsgBox "You have not selected a face!", vbExclamation, "Warning !"
        Exit Sub
    End If
    If swSelMgr.GetSelectedObjectType2(1) <> swSelFACES Then
        MsgBox "You did not select a face!", vbExclamation, "Warning !"
        Exit Sub
    End If
    
    ' Start the timer for performance measurement
    TimeStart = Timer

    ' Process the selected face
    Set swFace = swSelMgr.GetSelectedObject5(1)
    Set swSelData = swSelMgr.CreateSelectData
    swModel.ClearSelection2 True
    SelectHoleEdges swFace, swSelData

    ' Get the count of selected hole edges
    objCount = swSelMgr.GetSelectedObjectCount

    ' If no holes are found, show an informational message
    If objCount = 0 Then
        MsgBox "Zero hole found on selected face!", vbInformation, "Zero hole"
        Exit Sub
    End If

    ' End the timer and display results
    TimeEnd = Timer
    MsgBox _
    "Hole Diameter : " & Diameter(0) & " mm" & vbCrLf & vbCrLf & _
    "Material thickness : " & Thickness(0) & " mm" & vbCrLf & vbCrLf & _
    "Number of Holes : " & objCount & vbCrLf & vbCrLf & _
    "Time taken : " & Round((TimeEnd - TimeStart), 2) & " Seconds", , _
    "Hole Counting Macro V0.2"
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).