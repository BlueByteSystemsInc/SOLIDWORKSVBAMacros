# Suppress Threads, Add Mate Reference, Lower Image Quality, and Set BOM to Document Name

## Description
This macro performs the following actions on the active part:
1. **Suppresses the threads** in the model to optimize performance.
2. **Adds a mate reference** to the largest face of the part, allowing it to be used easily in assemblies.
3. **Lowers the image quality** (tessellation quality) to reduce the graphical load.
4. **Sets the Bill of Materials (BOM)** to use the document name instead of the configuration name.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - The active document must be a part.
> - The macro assumes that the thread features are labeled appropriately (e.g., "Cut-Sweep1", "Sweep1", "Cut-Extrude1").
> - The part must have valid bodies and faces for mate references.

## Results
> [!NOTE]
> - Thread features will be suppressed to improve performance.
> - A mate reference will be added to the largest face, which simplifies assembly creation.
> - Image quality (tessellation) will be lowered to reduce performance load.
> - The Bill of Materials (BOM) will be set to use the document name instead of the configuration name.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

' Main subroutine to suppress threads, add mate reference, lower image quality, and set BOM to document name
Sub main()
    ' Declare and initialize necessary SolidWorks objects
    Dim swApp As SldWorks.SldWorks                ' SolidWorks application object
    Dim swmodel As SldWorks.ModelDoc2             ' Active document object (part)
    Dim boolstatus As Boolean                     ' Boolean status to capture operation results
    Dim selmgr As SldWorks.SelectionMgr           ' Selection manager object
    Dim swfeatmgr As SldWorks.FeatureManager      ' Feature manager object
    Dim Configmgr As SldWorks.ConfigurationManager' Configuration manager object
    Dim swconfig As SldWorks.Configuration        ' Configuration object
    Dim swfeats As Variant                        ' Array of features in the part
    Dim feat As Variant                           ' Individual feature object
    Dim swBody As SldWorks.Body2                  ' Body object in the part
    Dim swFace As SldWorks.Face2                  ' Face object for mate reference
    Dim edges As Variant                          ' Array of edges for mate reference
    Dim templarge As SldWorks.Face2               ' Temporarily store the largest face
    Dim swEnt As SldWorks.Entity                  ' Entity object for selection
    Dim myFeature As SldWorks.Feature             ' Feature object for mate reference
    Dim i As Long                                 ' Loop counter for iterating through faces

    ' Initialize SolidWorks application and get the active document
    Set swApp = Application.SldWorks
    Set swmodel = swApp.ActiveDoc
    Set selmgr = swmodel.SelectionManager
    Set swfeatmgr = swmodel.FeatureManager
    Set Configmgr = swmodel.ConfigurationManager
    Set swconfig = Configmgr.ActiveConfiguration

    ' Suppressing threads by selecting specific features (e.g., Cut-Sweep, Sweep1, Cut-Extrude1)
    boolstatus = swmodel.Extension.SelectByID2("Cut-Sweep1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
    If selmgr.GetSelectedObjectCount = 0 Then
        boolstatus = swmodel.Extension.SelectByID2("Sweep1", "BODYFEATURE", 0, 0, 0, True, 0, Nothing, 0)
        boolstatus = swmodel.Extension.SelectByID2("Cut-Extrude1", "BODYFEATURE", 0, 0, 0, True, 0, Nothing, 0)
    End If
    swmodel.EditSuppress2    ' Suppress selected features
    swmodel.ClearSelection2 True  ' Clear the selection after suppression

    ' Set the BOM to use the document name instead of the configuration name
    boolstatus = swmodel.EditConfiguration3(swconfig.Name, swconfig.Name, "", "", 32)

    ' Lower image quality (tessellation quality) to reduce performance load
    swmodel.SetTessellationQuality 6  ' Set tessellation quality to lower value (6)

    ' Hide all sketches and planes
    swfeats = swfeatmgr.GetFeatures(False)
    For Each feat In swfeats
        ' Hide reference planes, sketches, and helixes
        If feat.GetTypeName = "RefPlane" Or feat.GetTypeName = "ProfileFeature" Or feat.GetTypeName = "Helix" Then
            feat.Select (True)
            swmodel.BlankRefGeom    ' Hide reference geometry
            swmodel.BlankSketch     ' Hide sketches
        End If
    Next

    ' Adding Mate Reference (only works with basic parts, not screws, washers, etc.)
    Dim vBodies As Variant
    vBodies = swmodel.GetBodies2(swAllBodies, True)
    Set swBody = vBodies(0)         ' Get the first body in the part
    Set swFace = swBody.GetFirstFace
    swmodel.ClearSelection2 True
    Set templarge = swFace          ' Initialize the largest face with the first face

    ' Find the largest face in the body (based on area)
    For i = 1 To swBody.GetFaceCount
        Set swEnt = swFace
        If swFace.GetArea > templarge.GetArea Then
            Set templarge = swFace  ' Update the largest face
        End If
        Set swFace = swFace.GetNextFace   ' Move to the next face
    Next i

    ' Get the edges of the largest face
    Set swEnt = templarge
    edges = templarge.GetEdges

    ' Add mate reference using the first edge of the largest face
    Set myFeature = swmodel.FeatureManager.InsertMateReference2("Mate Reference", edges(1), 0, 0, False, Nothing, 0, 0, False, Nothing, 0, 0)

End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).