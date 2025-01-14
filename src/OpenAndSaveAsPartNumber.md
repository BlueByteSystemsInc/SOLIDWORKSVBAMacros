# Open & Save As Part Number Macro

## Description
This macro automates the process of opening SolidWorks part files from a specified directory, retrieving a custom property (part number), and saving the parts to a new location with the part number as the filename. It streamlines file management and ensures consistency in naming conventions based on part properties.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer  
- **Operating System**: Windows 7 or later  

## Pre-Conditions
> [!NOTE]
> - The directories for the source files and the target save location must be set up in the macro.
> - The parts should contain a custom property named "Part Number".

## Results
> [!NOTE]
> - Parts will be saved in the new location with the filename set to their part number.
> - Any existing files with the same name in the target directory will be overwritten.

## Steps to Setup the Macro

### 1. **Configure Source and Target Directories**:
   - Modify the `OpenAndSaveas` subroutine calls to set the source (`X:\123\`) and target (`X:\ABC\`) directories.
   - Ensure that the target directory has write permissions.

### 2. **Run the Macro**:
   - Execute the `main` subroutine.
   - The macro opens each part in the source directory, reads its "Part Number" custom property, and saves it in the target directory with the part number as the new filename.

### 3. **Review Results**:
   - Check the target directory to ensure all files are correctly renamed and saved.
   - Verify that no files were inappropriately overwritten.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

' Declare variables for SolidWorks application, models, and file operations
Dim swApp As SldWorks.SldWorks              ' SolidWorks application object
Dim swModel As ModelDoc2                    ' Active SolidWorks document
Dim swFilename As String                    ' File path and name of the model
Dim nErrors As Long                         ' Variable to store error count
Dim nWarnings As Long                       ' Variable to store warning count
Dim retval As Long                          ' Return value for operations

Dim Response As String                      ' String to store the response from the directory listing
Dim DocName As String                       ' Document name for operations

' Main subroutine to initiate the process
Sub main()
    ' Initialize SolidWorks application object
    Set swApp = Application.SldWorks

    ' Specify the folder path containing files to process
    ' and invoke the OpenAndSaveas subroutine
    OpenAndSaveas "X:\123\", ".SLDPRT", True
End Sub

' Subroutine to open files from a folder and save them with a new name
Sub OpenAndSaveas(folder As String, ext As String, silent As Boolean)

    ' Declare variables for document type, custom properties, and file paths
    Dim swDocTypeLong As Long                    ' SolidWorks document type
    Dim swCustPrpMgr As SldWorks.CustomPropertyManager ' Custom property manager
    Dim Value As String                          ' Value of the custom property
    Dim GetName As String                        ' New file name with path
    Dim swSavePath As String                     ' Destination folder for saving

    ' Determine the document type based on the file extension
    ext = UCase$(ext) ' Convert file extension to uppercase
    swDocTypeLong = Switch( _
        ext = ".SLDPRT", swDocPART, _
        ext = ".SLDDRW", swDocDRAWING, _
        ext = ".SLDASM", swDocASSEMBLY, _
        True, -1 _
    )

    ' Exit if the file is not a recognized SolidWorks document type
    If swDocTypeLong = -1 Then Exit Sub

    ' Change the working directory to the specified folder
    ChDir (folder)

    ' Loop through all files in the folder
    Response = Dir(folder)
    Do Until Response = ""

        ' Construct the full file path for the current file
        swFilename = folder & Response

        ' Check if the file matches the specified extension
        If Right(UCase$(Response), 7) = ext Then
            
            ' Open the document in SolidWorks
            Set swModel = swApp.OpenDoc6(swFilename, swDocTypeLong, swOpenDocOptions_Silent, "", nErrors, nWarnings)
            
            ' For non-drawing files, set the view to Isometric
            If swDocTypeLong <> swDocDRAWING Then
                swModel.ShowNamedView2 "*Isometric", -1
            End If
            
            ' Access the custom property manager
            Set swCustPrpMgr = swModel.Extension.CustomPropertyManager("")
            
            ' Retrieve the "Part Number" custom property value
            swCustPrpMgr.Get3 "Part Number", False, "", Value
            
            ' Define the new save path and file name
            swSavePath = "X:\ABC\"
            GetName = swSavePath & Value & ".sldprt"
            
            ' Save the file with the new name
            swModel.SaveAs (GetName)
            
            ' Close the document after saving
            swApp.CloseDoc swModel.GetTitle
        End If

        ' Get the next file in the directory
        Response = Dir
    Loop

End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).