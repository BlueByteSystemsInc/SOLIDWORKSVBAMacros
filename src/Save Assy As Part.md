# Save Assembly as Part Macro for SolidWorks

## Description
This macro automates the process of saving an active SolidWorks assembly as a single part file containing only the exterior faces. This method is useful for simplifying complex assemblies into more manageable single files for archiving, basic interfacing, or performance optimization.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - A SolidWorks assembly (.SLDASM) must be open.
> - The assembly should contain at least one part.

## Results
> [!NOTE]
> - The active assembly is saved as a new part (.SLDPRT) file containing only the exterior faces.
> - The original assembly remains unchanged.

## Steps to Setup the Macro

1. **Prepare SolidWorks**:
   - Ensure SolidWorks is running with an assembly document open.
   - Verify that the assembly contains at least one part component.

2. **Configure and Run the Macro**:
   - Open the SolidWorks VBA editor by pressing (`Alt + F11`).
   - Insert a new module and copy the provided macro code into this module.
   - Run the macro from within SolidWorks by navigating to **Tools** > **Macro** > **Run**, then select your saved macro file.

3. **Using the Macro**:
   - The macro will automatically process the active document and save it as a new part file with the same name appended with "SLDPRT".

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

' Define module-level SolidWorks application variables
Dim swApp As SldWorks.SldWorks                    ' SolidWorks application object
Dim swModel As SldWorks.ModelDoc2                 ' Active SolidWorks document object
Dim boolstatus As Boolean                         ' Boolean to capture operation success

' Main subroutine
Sub main()
    ' Initialize SolidWorks application and get the active document
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    ' Check if a document is open
    If swModel Is Nothing Then
        MsgBox "No active document found. Please open a file."
        Exit Sub
    End If

    ' Variables for file path manipulation
    Dim FilePath As String                         ' Full path of the current file
    Dim PathSize As Long                           ' Length of the file path
    Dim PathNoExtension As String                 ' File path without extension
    Dim NewFilePath As String                     ' New file path for saving

    ' Extract the current file path and modify it for the new part file
    FilePath = swModel.GetPathName                 ' Get the full path of the active document
    PathSize = Strings.Len(FilePath)              ' Get the length of the file path
    PathNoExtension = Strings.Left(FilePath, PathSize - 6) ' Remove the last 6 characters (e.g., ".SLDASM")
    NewFilePath = PathNoExtension & "SLDPRT"      ' Append "SLDPRT" to create a new file path

    ' Save the assembly as a new part file containing only exterior faces
    boolstatus = swModel.SaveAs3(NewFilePath, 0, 0) ' Save operation with default options

    ' Check the operation success and notify the user
    If boolstatus Then
        MsgBox "Assembly saved as part file successfully at: " & NewFilePath
    Else
        MsgBox "Failed to save assembly as part file. Please check the file path and permissions."
    End If
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).