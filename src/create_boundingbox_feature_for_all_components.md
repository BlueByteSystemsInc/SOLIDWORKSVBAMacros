# Create Bounding Box for Assembly and Components


>[!WARNING]
> This macro creates the **SOLIDWORKS bounding box feature** and a not custom one. It requires SOLIDWORKS 2018 or newer. We have an alternative macro that uses sketch entities that create a tightest-fit bounding box.

## Macro Description

This macro creates a bounding box for the main assembly and its components within the active document in SOLIDWORKS. It traverses each component in the assembly, checking if a bounding box has already been created for the component. If not, the bounding box is created, and the component is processed.

## VBA Macro Code

```vbnet
Dim swApp As SldWorks.SldWorks
Dim swModel As ModelDoc2
Dim swAssembly As AssemblyDoc
Dim vComponents As Variant
Dim ProcessedFiles As New Collection

Sub main()

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    ' Ensure the active document is an assembly
    Set swAssembly = swModel
    
    ' Create bounding box for the main assembly
    CreateBoundingBox swModel
    
    ' Get all components of the assembly
    vComponents = swAssembly.GetComponents(False)
    
    Dim component
    For Each component In vComponents
        Dim swComponent As Component2
        Set swComponent = component
        
        Dim swComponentModelDoc As ModelDoc2
        Set swComponentModelDoc = swComponent.GetModelDoc2
        
        If Not swComponentModelDoc Is Nothing Then
            ' Check if component already processed
            If ExistsInCollection(ProcessedFiles, swComponentModelDoc.GetTitle()) = False Then
                ' Create bounding box for the component
                CreateBoundingBox swComponentModelDoc
                
                ' Add component to processed list
                ProcessedFiles.Add swComponentModelDoc.GetTitle(), swComponentModelDoc.GetTitle()
                
                ' Output component path for debugging
                Debug.Print swComponentModelDoc.GetPathName()
            End If
        End If
    Next component
End Sub

Sub CreateBoundingBox(ByRef swComponentModelDoc As ModelDoc2)
    
    ' Make the document visible
    swComponentModelDoc.Visible = True

    Dim swFeatureManager As featureManager
    Dim swBoundingBoxFeatureDefinition As BoundingBoxFeatureData
    Dim swBoundingBoxFeature As Feature
    
    ' Access the FeatureManager
    Set swFeatureManager = swComponentModelDoc.featureManager
    
    ' Define the bounding box feature
    Set swBoundingBoxFeatureDefinition = swFeatureManager.CreateDefinition(swConst.swFmBoundingBox)
    
    ' Set options for bounding box creation
    swBoundingBoxFeatureDefinition.ReferenceFaceOrPlane = swConst.swGlobalBoundingBoxFitOptions_e.swBoundingBoxType_BestFit
    swBoundingBoxFeatureDefinition.IncludeHiddenBodies = False
    swBoundingBoxFeatureDefinition.IncludeSurfaces = False
    
    ' Create the bounding box feature
    Set swBoundingBoxFeature = swFeatureManager.CreateFeature(swBoundingBoxFeatureDefinition)
    
    ' Make the document invisible again
    swComponentModelDoc.Visible = False
End Sub

Public Function ExistsInCollection(col As Collection, key As Variant) As Boolean
    On Error GoTo err
    ExistsInCollection = True
    IsObject (col.Item(key))
    Exit Function
err:
    ExistsInCollection = False
End Function

```


## System Requirements
To run this VBA macro, ensure that your system meets the following requirements:

- SOLIDWORKS Version: SOLIDWORKS 2018 or later
- VBA Environment: Pre-installed with SOLIDWORKS (Access via Tools > Macro > New or Edit)
- Operating System: Windows 7, 8, 10, or later

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).