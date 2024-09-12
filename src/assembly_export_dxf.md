# Export Sheet Metal to DXF in SOLIDWORKS

## Macro Description

This VBA macro automates the process of exporting all sheet metal parts from a SOLIDWORKS assembly to DXF files. The macro loops through each component in the assembly, checks if it's a sheet metal part, and exports the flat pattern of the part to a DXF file. The macro ensures that duplicate parts are not processed more than once, preventing redundant exports.

## VBA Macro Code

```vbnet
' ********************************************************************
' DISCLAIMER: 
' This code is provided as-is with no warranty or liability by 
' Blue Byte Systems Inc. The company assumes no responsibility for 
' any issues arising from the use of this code in production.
' ********************************************************************

' Enum for Sheet Metal export options
Enum SheetMetalOptions_e
    ExportFlatPatternGeometry = 1
    IncludeHiddenEdges = 2
    ExportBendLines = 4
    IncludeSketches = 8
    MergeCoplanarFaces = 16
    ExportLibraryFeatures = 32
    ExportFormingTools = 64
    ExportBoundingBox = 2048
End Enum

' Declare SolidWorks and component variables
Dim swApp As SldWorks.SldWorks
Dim swAssemblyDoc As AssemblyDoc
Dim swcomponent As Component2
Dim vcomponents As Variant
Dim processedFiles As New Collection
Dim component

Sub main()

    ' Get the active SolidWorks application
    Set swApp = Application.SldWorks

    ' Get the active document (assembly)
    Set swAssemblyDoc = swApp.ActiveDoc

    ' Ensure that the active document is an assembly
    If swAssemblyDoc Is Nothing Then
        MsgBox "Please open an assembly document.", vbExclamation
        Exit Sub
    End If

    ' Get all components of the assembly
    vcomponents = swAssemblyDoc.GetComponents(False)

    ' Loop through each component in the assembly
    For Each component In vcomponents

        ' Set the component
        Set swcomponent = component

        ' Get the ModelDoc2 (part or assembly) for the component
        Dim swmodel As ModelDoc2
        Set swmodel = swcomponent.GetModelDoc2

        ' Check if the model is valid
        If Not swmodel Is Nothing Then

            ' Check if the model has already been processed
            If ExistsInCollection(processedFiles, swmodel.GetTitle()) = False Then

                ' Export the sheet metal part to DXF
                PrintDXF swmodel

                ' Add the processed file to the collection to avoid duplicates
                processedFiles.Add swmodel.GetTitle(), swmodel.GetTitle()

            End If

        End If

    Next

End Sub

' Function to export a sheet metal part to DXF
Function PrintDXF(ByRef swmodel As ModelDoc2) As String

    ' Check if the document is a part file
    If swmodel.GetType() = SwConst.swDocumentTypes_e.swDocPART Then

        Dim swPart As PartDoc
        Set swPart = swmodel

        ' Get the model path
        Dim modelPath As String
        modelPath = swPart.GetPathName

        ' Define the output DXF path
        Dim OUT_PATH As String
        OUT_PATH = Left(modelPath, Len(modelPath) - 6) ' Remove ".SLDPRT" extension
        OUT_PATH = OUT_PATH + "dxf"

        ' Make the model visible before exporting
        swmodel.Visible = True

        ' Export the sheet metal part to DXF using specified options
        If False = swPart.ExportToDWG2(OUT_PATH, modelPath, swExportToDWG_e.swExportToDWG_ExportSheetMetal, _
                                       True, Empty, False, False, _
                                       SheetMetalOptions_e.ExportFlatPatternGeometry + _
                                       SheetMetalOptions_e.ExportBendLines, Empty) Then
            ' Raise error if export fails
            err.Raise vbError, "", "Failed to export flat pattern"
        End If

        ' Hide the model after exporting
        swmodel.Visible = False

    End If

    ' Print the model path to the debug console
    Debug.Print swmodel.GetPathName()

End Function

' Function to check if an item exists in a collection
Public Function ExistsInCollection(col As Collection, key As Variant) As Boolean
    On Error GoTo err
    ExistsInCollection = True
    IsObject (col.Item(key)) ' Check if the item exists in the collection
    Exit Function
err:
    ExistsInCollection = False ' Return false if the item does not exist
End Function
```

## System Requirements
To run this VBA macro, ensure that your system meets the following requirements:

- SOLIDWORKS Version: SOLIDWORKS 2017 or later
- VBA Environment: Pre-installed with SOLIDWORKS (Access via Tools > Macro > New or Edit)
- Operating System: Windows 7, 8, 10, or later


>[!NOTE] 
> Pre-conditions
> - The active document must be an assembly (.sldasm) in SOLIDWORKS.
> - Ensure that the components contain valid sheet metal parts for export.

>[!NOTE] 
> Post-conditions
> The flat pattern of each sheet metal part will be exported as a DXF file.

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).