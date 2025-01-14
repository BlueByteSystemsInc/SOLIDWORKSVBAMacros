# Export Flat Pattern Macro

## Description
This macro exports the flat pattern view of all configurations in an open sheet metal part as DXF files. The DXF files are saved in the same directory as the part file with the configuration name appended to the file name.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer  
- **Operating System**: Windows 7 or later  

## Pre-Conditions
> [!NOTE]
> - A sheet metal part must be open in SolidWorks.
> - The document must be a valid part file.

## Results
> [!NOTE]
> - Flat pattern DXF files are created for all configurations in the part file.
> - The files are saved in the same directory as the part file.

## Steps to Use the Macro

### **1. Open a Sheet Metal Part**
   - Ensure the active document in SolidWorks is a sheet metal part.

### **2. Execute the Macro**
   - Run the macro to generate DXF files for each configuration in the part.

### **3. Verify Exported Files**
   - Check the directory containing the part file for the generated DXF files.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

Sub main()

    ' Declare variables for SolidWorks application and active document
    Dim swApp                   As SldWorks.SldWorks
    Dim swModel                 As SldWorks.ModelDoc2
    Dim vConfNameArr            As Variant  ' Array to hold configuration names
    Dim sConfigName             As String   ' Current configuration name
    Dim i                       As Long     ' Loop counter
    Dim bShowConfig             As Boolean  ' Flag for showing configuration
    Dim bRebuild                As Boolean  ' Flag for rebuilding the model
    Dim bRet                    As Boolean  ' Flag for export success
    Dim FilePath                As String   ' File path of the part
    Dim PathSize                As Long     ' Length of the file path
    Dim PathNoExtension         As String   ' File path without extension
    Dim NewFilePath             As String   ' File path for the new DXF file

    ' Initialize SolidWorks application and get the active document
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    ' Check if a document is active
    If swModel Is Nothing Then
        MsgBox "No document is open. Please open a sheet metal part and try again.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Validate if the document is a part
    If swModel.GetType <> swDocPART Then
        MsgBox "This macro only supports sheet metal parts. Please open a sheet metal part and try again.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Get the file path of the active part
    FilePath = swModel.GetPathName
    If FilePath = "" Then
        MsgBox "The part must be saved before running the macro.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Get the list of all configuration names in the part
    vConfNameArr = swModel.GetConfigurationNames

    ' Loop through each configuration
    For i = 0 To UBound(vConfNameArr)
        sConfigName = vConfNameArr(i)  ' Current configuration name

        ' Show the configuration
        bShowConfig = swModel.ShowConfiguration2(sConfigName)

        ' Rebuild the model to ensure the configuration is up-to-date
        bRebuild = swModel.ForceRebuild3(False)

        ' Construct the file path for the DXF file
        PathSize = Strings.Len(FilePath)  ' Get the length of the file path
        PathNoExtension = Strings.Left(FilePath, PathSize - 6)  ' Remove extension from file path
        NewFilePath = PathNoExtension & "_" & sConfigName & ".DXF"  ' Append configuration name and DXF extension

        ' Export the flat pattern as a DXF file
        bRet = swModel.ExportFlatPatternView(NewFilePath, 1)

        ' Check if the export was successful and log the result
        If bRet Then
            Debug.Print "Successfully exported: " & NewFilePath
        Else
            MsgBox "Failed to export flat pattern for configuration: " & sConfigName, vbExclamation, "Export Error"
        End If
    Next i

    ' Notify the user of successful completion
    MsgBox "Flat patterns exported successfully for all configurations.", vbInformation, "Export Complete"

End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).