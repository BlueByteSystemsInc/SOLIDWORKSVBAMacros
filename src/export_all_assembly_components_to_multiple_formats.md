# Export Configurations of Assembly Components to Multiple Formats

## Description

This macro automates the process of exporting all configurations of components within an active SOLIDWORKS assembly to multiple file formats. It iterates through each component in the assembly, checks if it has already been processed to avoid duplicates, and then exports each configuration of the component to `.step`, `.igs`, and `.x_t` formats. The exported files are saved in a specified directory with filenames that include the component name and configuration name.

## System Requirements

- **SOLIDWORKS Version**: SOLIDWORKS 2014 or newer
- **Operating System**: Windows 7 or later

## VBA Code

```vbnet
'*********************************************************
' Blue Byte Systems Inc.
' Disclaimer: Blue Byte Systems Inc. provides this macro "as-is" without any warranties.
' Use at your own risk. The company is not liable for any damages resulting from its use.
'*********************************************************

Option Explicit

Dim swApp As SldWorks.SldWorks
Dim swAssembly As SldWorks.AssemblyDoc
Dim swComponent As SldWorks.Component2
Dim vComponents As Variant
Dim processedFiles As Collection
Dim component As Variant

Sub Main()

    Dim errors As Long
    Dim warnings As Long

    ' Initialize the collection to keep track of processed files
    Set processedFiles = New Collection

    ' Get the SOLIDWORKS application object
    Set swApp = Application.SldWorks

    ' Get the active document and ensure it is an assembly
    Set swAssembly = swApp.ActiveDoc
    If swAssembly Is Nothing Then
        MsgBox "No active document found.", vbExclamation, "Error"
        Exit Sub
    End If

    If swAssembly.GetType <> swDocumentTypes_e.swDocASSEMBLY Then
        MsgBox "The active document is not an assembly.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Get all components in the assembly
    vComponents = swAssembly.GetComponents(False)

    ' Iterate through each component
    For Each component In vComponents

        Set swComponent = component

        Dim swModel As SldWorks.ModelDoc2
        Set swModel = swComponent.GetModelDoc2

        If Not swModel Is Nothing Then

            ' Check if the model has already been processed
            If Not ExistsInCollection(processedFiles, swModel.GetTitle()) Then

                ' Save configurations of the model
                SaveConfigurations swModel

                ' Add the model to the processed files collection
                processedFiles.Add swModel.GetTitle(), swModel.GetTitle()

            End If

        End If

    Next component

    MsgBox "Export completed successfully.", vbInformation, "Done"

End Sub

Sub SaveConfigurations(ByRef swModel As SldWorks.ModelDoc2)

    Dim extensions(1 To 3) As String
    extensions(1) = ".step"
    extensions(2) = ".igs"
    extensions(3) = ".x_t"

    swModel.Visible = True

    Dim configurationNames As Variant
    configurationNames = swModel.GetConfigurationNames

    Dim configName As Variant
    For Each configName In configurationNames

        swModel.ShowConfiguration2 configName

        Dim extension As Variant
        For Each extension In extensions

            Dim outputPath As String
            outputPath = "C:\BOM Export\"
            outputPath = outputPath & Left(swModel.GetTitle(), 6) & "_" & configName & extension

            Dim saveSuccess As Boolean
            Dim errors As Long
            Dim warnings As Long

            saveSuccess = swModel.Extension.SaveAs3(outputPath, _
                            swSaveAsVersion_e.swSaveAsCurrentVersion, _
                            swSaveAsOptions_e.swSaveAsOptions_Silent, _
                            Nothing, Nothing, errors, warnings)

            If Not saveSuccess Then
                MsgBox "Failed to save: " & outputPath, vbExclamation, "Error"
            End If

        Next extension

    Next configName

    swModel.Visible = False

End Sub

Function ExistsInCollection(col As Collection, key As Variant) As Boolean
    On Error GoTo ErrHandler
    ExistsInCollection = True
    Dim temp As Variant
    temp = col.Item(key)
    Exit Function
ErrHandler:
    ExistsInCollection = False
End Function
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).