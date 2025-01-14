# Change Part Material Macro

## Description
This macro automates the process of updating the material property of all SolidWorks part files in a specified directory. It opens each part file, changes its material to the specified one, rebuilds the model, and saves the changes.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer  
- **Operating System**: Windows 7 or later  

## Pre-Conditions
> [!NOTE]
> - Ensure that the target directory contains valid SolidWorks part files (`.SLDPRT`).
> - Update the macro with the desired directory path and material name before running.

## Results
> [!NOTE]
> - All part files in the specified directory will have their material updated to the specified material (e.g., "Brass").
> - A confirmation message will display the name of the applied material once the macro completes.

## Steps to Use the Macro

### **1. Configure Folder and Material**
   - Update the `SheetFormat` subroutine in the macro to specify:
     - The target directory (`D:\SW\`) containing the `.SLDPRT` files.
     - The desired material (e.g., `"Brass"`).

### **2. Execute the Macro**
   - Run the macro in SolidWorks. It will iterate through all part files in the specified directory, apply the material, rebuild the model, and save the changes.

### **3. View Results**
   - A message box will display confirming that the material has been updated for all parts.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

' Declare global variables
Dim swApp As SldWorks.SldWorks       ' SolidWorks application object
Dim swModel As ModelDoc2             ' Active document object
Dim swFilename As String             ' Current file path
Dim nErrors As Long                  ' Error count during file open/save
Dim nWarnings As Long                ' Warning count during file open/save
Dim Response As String               ' Directory file response
Dim configName As String             ' Configuration name
Dim databaseName As String           ' Material database name
Dim newPropName As String            ' New material name
Dim bShowConfig As Boolean           ' Show configuration flag

Sub main()

    ' Initialize SolidWorks application
    Set swApp = Application.SldWorks

    ' Specify the folder path containing part files and run the material update
    ' Change the path to your desired folder containing SolidWorks part files
    UpdateMaterial "D:\SW\", ".SLDPRT", True

End Sub

' Subroutine to update material for all files in the specified folder
Sub UpdateMaterial(folder As String, ext As String, silent As Boolean)

    Dim swDocTypeLong As Long         ' Document type identifier

    ' Ensure the file extension is in uppercase
    ext = UCase$(ext)

    ' Identify document type based on extension
    swDocTypeLong = Switch(ext = ".SLDPRT", swDocPART, True, -1)

    ' Exit if not a SolidWorks part file
    If swDocTypeLong = -1 Then Exit Sub

    ' Change to the specified folder
    ChDir folder

    ' Get the first file in the folder
    Response = Dir(folder)
    Do Until Response = ""

        ' Build the full file path
        swFilename = folder & Response

        ' Process files with the specified extension
        If Right(UCase$(Response), 7) = ext Then

            ' Open the file silently
            Set swModel = swApp.OpenDoc6(swFilename, swDocTypeLong, swOpenDocOptions_Silent, "", nErrors, nWarnings)

            ' Apply material if the file is a part
            If swDocTypeLong = swDocPART Then
                ApplyMaterial swModel, "Default", "SolidWorks Materials", "Brass"
            End If

            ' Rebuild, save, and close the document
            swModel.ViewZoomtofit2
            swModel.ForceRebuild3 False
            swModel.Save2 silent
            swApp.CloseDoc swModel.GetTitle

        End If

        ' Move to the next file in the folder
        Response = Dir

    Loop

    ' Notify the user about the material update
    MsgBox "Material updated/changed to Brass", vbOKOnly, "Material Update Complete!"

End Sub

' Subroutine to apply material to the part
Sub ApplyMaterial(swModel As ModelDoc2, configName As String, databaseName As String, newMaterialName As String)

    ' Show the configuration (if applicable)
    bShowConfig = swModel.ShowConfiguration2(configName)

    ' Debug information (optional for development)
    Debug.Print "  Configuration Name: " & configName

    ' Set the material for the specified configuration
    swModel.SetMaterialPropertyName2 configName, databaseName, newMaterialName

End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).