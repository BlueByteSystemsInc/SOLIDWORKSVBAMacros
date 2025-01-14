# Sweep Length Calculation Macro for SolidWorks

## Description
This macro calculates the total length of all sweep features in an active SolidWorks part document. It adds a custom property to the part that stores the total length, making it accessible for documentation or further calculations.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - A SolidWorks part document must be open.
> - The part should contain one or more sweep features.

## Results
> [!NOTE]
> - Calculates the cumulative length of all sweep features.
> - Adds a custom property to the part document that lists the total length.

## Steps to Setup the Macro

1. **Create the VBA Modules**:
   - Open the SolidWorks VBA editor by pressing (`Alt + F11`).
   - Insert a new module into your project and copy the provided macro code into this module.

2. **Run the Macro**:
   - Ensure that a part with sweep features is open.
   - Run the macro from within SolidWorks by navigating to **Tools** > **Macro** > **Run**, then select your saved macro file.

3. **Using the Macro**:
   - The macro automatically processes all sweep features in the active document and calculates their total length.
   - The total length is stored as a custom property in the part document for easy reference.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

' SolidWorks VBA macro to calculate the total length of all sweep features in a part

Option Explicit

' Main subroutine to calculate the total length of sweep features in a part
Sub main()
    ' Declare variables for SolidWorks application, part document, and features
    Dim swApp As SldWorks.SldWorks                  ' SolidWorks application object
    Dim swPart As SldWorks.PartDoc                  ' Active part document
    Dim FeatureV As Variant                         ' Variant for iterating over features
    Dim myFeature As SldWorks.Feature               ' Feature object
    Dim Features As Variant                         ' Array of features in the part
    Dim swSweepFtData As SldWorks.SweepFeatureData  ' Sweep feature data object
    Dim swSweepFt As SldWorks.Feature               ' Sweep feature object
    Dim Path As Object                              ' Path object for sweep
    Dim swSweepPathSk As SldWorks.Sketch            ' Sketch object representing the sweep path
    Dim swSketchSeg As SldWorks.SketchSegment       ' Sketch segment object
    Dim swModelDocExt As SldWorks.ModelDocExtension ' ModelDocExtension object for custom properties
    Dim swCustProp As SldWorks.CustomPropertyManager ' Custom property manager object
    Dim swConfig As SldWorks.Configuration          ' Active configuration object
    Dim TotalLength As Double                       ' Total length of all sweep features
    Dim Length As Double                            ' Length of the current sweep path

    ' Initialize SolidWorks application and active part document
    Set swApp = Application.SldWorks
    Set swPart = swApp.ActiveDoc

    ' Get all features in the part
    Features = swPart.FeatureManager.GetFeatures(False)
    TotalLength = 0 ' Initialize total length to zero

    ' Process each feature to find sweep features and calculate total length
    For Each FeatureV In Features
        Set myFeature = FeatureV
        If myFeature.GetTypeName2 = "Sweep" Then ' Check if the feature is a sweep
            Set swSweepFt = myFeature
            Set swSweepFtData = swSweepFt.GetDefinition
            Call swSweepFtData.AccessSelections(swPart, Nothing) ' Access the sweep feature data
            
            ' Get the path of the sweep
            Set Path = swSweepFtData.Path
            If Not Path Is Nothing Then
                ' Get the sketch for the sweep path
                Set swSweepPathSk = Path.GetSpecificFeature2
                Dim vSketchSeg As Variant
                vSketchSeg = swSweepPathSk.GetSketchSegments
                Length = 0 ' Initialize length for the current sweep
                For i = 0 To UBound(vSketchSeg)
                    Set swSketchSeg = vSketchSeg(i)
                    If swSketchSeg.ConstructionGeometry = False Then ' Exclude construction geometry
                        Length = swSketchSeg.GetLength + Length ' Accumulate length
                    End If
                Next i
                Length = Length * 1000 ' Convert to millimeters
                Call swSweepFtData.ReleaseSelectionAccess ' Release the selections
            End If
            TotalLength = TotalLength + Length ' Add to total length
        End If
    Next FeatureV

    ' Store the total length in a custom property
    Set swModelDocExt = swPart.Extension
    Set swConfig = swPart.GetActiveConfiguration
    Set swCustProp = swModelDocExt.CustomPropertyManager(swConfig.Name)
    Call swCustProp.Add3("Length", swCustomInfoDouble, TotalLength, swCustomPropertyReplaceValue)

    ' Display the total length
    MsgBox "Total Length = " & TotalLength & "mm"
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).