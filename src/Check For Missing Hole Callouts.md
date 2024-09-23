# Check For Missing Hole Callouts

## Description
This macro checks a SOLIDWORKS drawing to ensure all Hole Wizard features in the referenced part have corresponding hole callouts. It iterates through each view, retrieves Hole Wizard features, and compares them with those in the drawing. If any callouts are missing, it notifies the user, ensuring proper documentation of all holes.

## System Requirements
- **SOLIDWORKS Version**: SOLIDWORKS 2014 or newer
- **Operating System**: Windows 10 or later

## Pre-Conditions
> [!NOTE]
> - Active SOLIDWORKS Drawing: A drawing document (.SLDDRW) must be open and active in SOLIDWORKS.
> - Referenced Part Document: The drawing must contain at least one view that references a part document (.SLDPRT) with Hole Wizard features.
> - Hole Wizard Features: The referenced part must have holes created using the Hole Wizard feature, not manual cut-extrudes or other methods.
> - Proper Naming: Hole Wizard features and hole callouts should have consistent naming conventions if applicable.

## Results
> [!NOTE]
> - Hole Callout Verification: The macro will analyze the drawing and identify any Hole Wizard features in the > - > - referenced part that do not have corresponding hole callouts in the drawing views.
> - User Notification: If missing hole callouts are found, the macro will display messages listing the specific Hole Wizard features that lack callouts.
> - Design Checker Update: The macro will set custom check results in the SOLIDWORKS Design Checker, highlighting the failed items (missing hole callouts).

## VBA Macro Code
```vbnet
' All rights reserved to Blue Byte Systems Inc.
' Blue Byte Systems Inc. does not provide any warranties for macros.
' This macro compares the Hole Wizard features in a drawing with hole callouts in a view.

Sub main()
    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swView As SldWorks.View
    Dim swDrawingDoc As SldWorks.DrawingDoc
    Dim swPart As SldWorks.ModelDoc2
    Dim sModelName As String
    Dim swFeature As SldWorks.Feature
    Dim totalFeatures As Long
    Dim featureName As String
    Dim i As Long
    Dim featureType As String
    Dim totalHoleWzd As Long
    Dim holeWizardFeatures(50) As String
    Dim swDisplayDimension As SldWorks.DisplayDimension
    Dim attachedEntityArr As Variant
    Dim swEntity As SldWorks.Entity
    Dim swAnnotation As SldWorks.Annotation
    Dim swEdge As SldWorks.Edge
    Dim faceEntities As Variant
    Dim swFace1 As SldWorks.Face2
    Dim swFace2 As SldWorks.Face2
    Dim swHoleWzdFeature As SldWorks.Feature
    Dim holeCalloutFeatures(50) As String
    Dim missingHoleCallouts(50) As String
    Dim missingCount As Long
    Dim comparisonCount As Long
    Dim featureCheck As Boolean
    Dim errorCode As Long
    Dim failedItemsArr() As String

    ' Initialize SOLIDWORKS application
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set swDrawingDoc = swModel
    Set swView = swDrawingDoc.GetFirstView

    ' Get the first display dimension in the view
    Set swDisplayDimension = swView.GetFirstDisplayDimension()
    Set swPart = swView.ReferencedDocument

    ' Loop through views until a part is found
    Do While swPart Is Nothing
        Set swView = swView.GetNextView
        Set swPart = swView.ReferencedDocument
    Loop

    ' Start processing the hole callouts
    missingCount = 0
    Do While Not swView Is Nothing
        Set swDisplayDimension = swView.GetFirstDisplayDimension()
        Do While Not swDisplayDimension Is Nothing
            ' Check if the dimension is a hole callout
            If swDisplayDimension.IsHoleCallout <> False Then
                Set swAnnotation = swDisplayDimension.GetAnnotation
                attachedEntityArr = swAnnotation.GetAttachedEntities3
                Set swEntity = attachedEntityArr(0)
                Set swEdge = swEntity
                faceEntities = swEdge.GetTwoAdjacentFaces2()
                Set swFace1 = faceEntities(0)
                Set swFace2 = faceEntities(1)

                ' Get the feature associated with the face
                Set swHoleWzdFeature = swFace1.GetFeature
                If swHoleWzdFeature.GetTypeName = "HoleWzd" Then
                    holeCalloutFeatures(missingCount) = swHoleWzdFeature.Name
                    missingCount = missingCount + 1
                Else
                    Set swHoleWzdFeature = swFace2.GetFeature
                    holeCalloutFeatures(missingCount) = swHoleWzdFeature.Name
                    missingCount = missingCount + 1
                End If
            End If
            Set swDisplayDimension = swDisplayDimension.GetNext
        Loop
        Set swView = swView.GetNextView
    Loop

    ' Count total Hole Wizard features in the referenced model
    totalFeatures = swPart.GetFeatureCount
    totalHoleWzd = 0

    For i = totalFeatures To 1 Step -1
        Set swFeature = swPart.FeatureByPositionReverse(totalFeatures - i)
        If Not swFeature Is Nothing Then
            featureName = swFeature.Name
            featureType = swFeature.GetTypeName
            If featureType = "HoleWzd" Then
                If swFeature.IsSuppressed = False Then
                    holeWizardFeatures(totalHoleWzd) = featureName
                    totalHoleWzd = totalHoleWzd + 1
                End If
            End If
        End If
    Next

    ' Compare Hole Wizard features with hole callout features
    comparisonCount = 0

    For i = 0 To totalHoleWzd
        featureCheck = False
        For comparisonCount = 0 To missingCount
            If holeWizardFeatures(i) = holeCalloutFeatures(comparisonCount) Then
                featureCheck = True
            End If
        Next comparisonCount

        ' Store missing features
        If featureCheck = False Then
            missingHoleCallouts(comparisonCount) = holeWizardFeatures(i)
            comparisonCount = comparisonCount + 1
        End If
    Next

    ' If any features are missing, report them
    If comparisonCount > 0 Then
        ReDim Preserve failedItemsArr(1 To 2, 1 To comparisonCount) As String
        For i = 0 To comparisonCount - 1
            failedItemsArr(1, i + 1) = missingHoleCallouts(i)
            MsgBox "YOU HAVE MISSED THE FOLLOWING FEATURE: " & missingHoleCallouts(i)
        Next
        Dim dcApp As Object
        Set dcApp = swApp.GetAddInObject("SWDesignChecker.SWDesignCheck")
        errorCode = dcApp.SetCustomCheckResult(False, failedItemsArr)
    End If
End Sub

```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).