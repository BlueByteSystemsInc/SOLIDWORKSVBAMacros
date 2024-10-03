# Traverse Assembly and Save Parts as DXF

## Description
This macro traverses the active assembly and saves all its child components (parts) as DXF files in the specified folder. It recursively traverses through the assembly hierarchy, flattens any sheet metal parts, and exports the flat pattern as a DXF file. This macro is designed to streamline the export process for sheet metal parts.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - The active document must be an assembly.
> - Sheet metal parts should be set up correctly for flattening and exporting.
> - A folder path must be provided where the DXF files will be saved.

## Results
> [!NOTE]
> - All sheet metal parts within the active assembly are exported as DXF files.
> - The DXF files will be saved in the specified folder.
> - The macro will skip any parts that are not sheet metal.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

' --------------------------------------------------------------------------
' Main subroutine to start the process and get user input for save path
' --------------------------------------------------------------------------
Sub main()

    ' Declare necessary SolidWorks objects
    Dim swApp As SldWorks.SldWorks              ' SolidWorks application object
    Dim swModel As SldWorks.ModelDoc2           ' Active document object (assembly)
    Dim savepath As String                      ' User input for the folder path to save DXF files

    ' Initialize SolidWorks application
    Set swApp = CreateObject("SldWorks.Application")

    ' Get the currently active document
    Set swModel = swApp.ActiveDoc

    ' Check if there is an active document open
    If swModel Is Nothing Then
        MsgBox "No active document found. Please open an assembly and try again.", vbCritical, "No Active Document"
        Exit Sub
    End If

    ' Prompt user for the folder path to save DXF files
    savepath = InputBox("Where do you want to save the DXF files?")

    ' Call the traverse function to iterate through components and export DXF files
    traverse swModel, savepath

End Sub

' --------------------------------------------------------------------------
' Recursive function to traverse through the assembly components and export parts
' --------------------------------------------------------------------------
Function traverse(Pathname As ModelDoc2, savepath As String)

    ' Declare necessary variables and objects
    Dim swApp As SldWorks.SldWorks                  ' SolidWorks application object
    Dim swModel As SldWorks.ModelDoc2               ' Model document object
    Dim swRootComp As SldWorks.Component2           ' Root component of the assembly
    Dim swConf As SldWorks.Configuration            ' Configuration of the active assembly
    Dim swConfMgr As SldWorks.ConfigurationManager  ' Configuration manager of the active assembly
    Dim vChildComp As Variant                       ' Array of child components in the assembly
    Dim swChildComp As SldWorks.Component2          ' Child component object
    Dim i As Long                                   ' Loop counter for iterating through child components

    ' Initialize SolidWorks application
    Set swApp = CreateObject("SldWorks.Application")

    ' Set the active model to the passed Pathname parameter
    Set swModel = Pathname

    ' Get the configuration manager and active configuration of the model
    Set swConfMgr = swModel.ConfigurationManager
    Set swConf = swConfMgr.ActiveConfiguration

    ' Get the root component of the assembly
    Set swRootComp = swConf.GetRootComponent3(True)

    ' Get the children components of the root component
    vChildComp = swRootComp.GetChildren

    ' Loop through each child component
    For i = 0 To UBound(vChildComp)
        Set swChildComp = vChildComp(i)

        ' Get the model document of the child component
        Set swModel = swChildComp.GetModelDoc2

        ' If the child component is a part, traverse further or export as DXF
        If Not swModel Is Nothing Then

            ' Check if the component is an assembly (type 2 = swDocASSEMBLY)
            If swModel.GetType = 2 Then
                traverse swModel, savepath ' Recursively traverse through sub-assemblies

            ' If the component is a part, flatten and export as DXF
            Else
                flat swModel, savepath
            End If
        End If
    Next i

End Function

' --------------------------------------------------------------------------
' Function to flatten sheet metal parts and save as DXF
' --------------------------------------------------------------------------
Sub flat(swModel As SldWorks.ModelDoc2, savepath As String)

    ' Declare necessary variables and objects
    Dim swApp As SldWorks.SldWorks                ' SolidWorks application object
    Dim swFeat As SldWorks.Feature                ' Feature object to access flat pattern feature
    Dim swFlat As SldWorks.Feature                ' Flat pattern feature object

    ' Initialize SolidWorks application
    Set swApp = CreateObject("SldWorks.Application")

    ' Get the first feature in the part
    Set swFeat = swModel.FirstFeature

    ' Loop through each feature to find the "FlatPattern" feature
    Do While Not swFeat Is Nothing

        ' Check if the feature is a "FlatPattern" feature
        If swFeat.GetTypeName = "FlatPattern" Then

            ' Un-suppress the flat pattern
            swFeat.Select (True)
            swModel.EditUnsuppress2

            ' Export the part as a DXF file
            dxf swModel, savepath

            ' Re-suppress the flat pattern
            swFeat.Select (True)
            swModel.EditSuppress2
        End If

        ' Move to the next feature in the model
        Set swFeat = swFeat.GetNextFeature
    Loop

End Sub

' --------------------------------------------------------------------------
' Function to export the flat pattern of the part as a DXF file
' --------------------------------------------------------------------------
Public Function dxf(swModel As SldWorks.ModelDoc2, savepath As String)

    ' Declare necessary variables
    Dim swApp As SldWorks.SldWorks                ' SolidWorks application object
    Dim swPart As SldWorks.PartDoc                ' Part document object
    Dim sModelName As String                      ' Model name of the part
    Dim sPathName As String                       ' Path name of the DXF file
    Dim varAlignment As Variant                   ' Alignment data for exporting
    Dim dataAlignment(11) As Double               ' Alignment data array
    Dim varViews As Variant                       ' Views data for exporting
    Dim dataViews(1) As String                    ' Views data array
    Dim options As Long                           ' Export options

    ' Initialize SolidWorks application
    Set swApp = Application.SldWorks
    swApp.ActivateDoc swModel.GetPathName

    ' Check if the part is in the bent state (flat pattern should be unsuppressed)
    If swModel.GetBendState <> 2 Then
        Exit Function
    End If

    ' Get the model name and set the path for DXF file
    sModelName = swModel.GetPathName
    sPathName = savepath & "\" & swModel.GetTitle & ".dxf"

    ' Set alignment and view data for DXF export
    dataAlignment(0) = 0#: dataAlignment(1) = 0#: dataAlignment(2) = 0#
    dataAlignment(3) = 1#: dataAlignment(4) = 0#: dataAlignment(5) = 0#
    dataAlignment(6) = 0#: dataAlignment(7) = 1#: dataAlignment(8) = 0#
    dataAlignment(9) = 0#: dataAlignment(10) = 0#: dataAlignment(11) = 1#
    varAlignment = dataAlignment

    dataViews(0) = "*Current"
    dataViews(1) = "*Front"
    varViews = dataViews

    ' Export the flat pattern of the sheet metal part to DXF file
    options = 13 ' Export options for flat pattern geometry, bend lines, and sketches
    swPart.ExportToDWG sPathName, sModelName, 1, True, varAlignment, False, False, options, Null

    ' Close the part document after exporting
    swApp.CloseDoc (swModel.GetPathName)

End Function
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).