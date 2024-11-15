# Export Sheet Metal Flat Pattern as DXF for CAM Software

## Description
This macro unfolds a sheet metal part and saves it as a DXF file, configured for use in CAM programs, with flat pattern views exported.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - The active document must be a sheet metal part.
> - The part must have valid configurations with flat pattern views enabled.

## Results
> [!NOTE]
> - The sheet metal part will be unfolded (flattened) and saved as a DXF file.
> - The DXF file will be saved in the same directory as the part with the same name but with a `.DXF` extension.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

' --------------------------------------------------------------------------
' Main subroutine to export sheet metal flat patterns as DXF for CAM software
' --------------------------------------------------------------------------
Sub main()

    ' Declare and initialize necessary SolidWorks objects
    Dim swApp As SldWorks.SldWorks                ' SolidWorks application object
    Dim swModel As SldWorks.ModelDoc2             ' Active model document (part)
    Dim vConfNameArr As Variant                   ' Array to hold configuration names
    Dim sConfigName As String                     ' String to hold the current configuration name
    Dim i As Long                                 ' Loop counter for iterating through configurations
    Dim bShowConfig As Boolean                    ' Boolean to show the configuration
    Dim bRebuild As Boolean                       ' Boolean to rebuild the model
    Dim bRet As Boolean                           ' Boolean to check the success of DXF export
    Dim FilePath As String                        ' String to hold the path of the current file
    Dim PathSize As Long                          ' Length of the current file path
    Dim PathNoExtension As String                 ' File path without extension
    Dim NewFilePath As String                     ' New file path for the DXF file

    ' Initialize SolidWorks application and get the active document (sheet metal part)
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    ' Get all configuration names of the active part
    vConfNameArr = swModel.GetConfigurationNames

    ' Loop through each configuration in the part
    For i = 0 To UBound(vConfNameArr)

        ' Get the name of the current configuration
        sConfigName = vConfNameArr(i)

        ' Show the current configuration
        bShowConfig = swModel.ShowConfiguration2(sConfigName)

        ' Rebuild the model after showing the configuration
        bRebuild = swModel.ForceRebuild3(True)

        ' Get the file path of the current part
        FilePath = swModel.GetPathName

        ' Calculate the size of the path and remove the file extension
        PathSize = Strings.Len(FilePath)
        PathNoExtension = Strings.Left(FilePath, PathSize - 7)

        ' Generate the new file path with the ".DXF" extension
        NewFilePath = PathNoExtension & ".DXF"

        ' Export the flat pattern view as a DXF file
        bRet = swModel.ExportFlatPatternView(NewFilePath, 1)

    Next i

End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).