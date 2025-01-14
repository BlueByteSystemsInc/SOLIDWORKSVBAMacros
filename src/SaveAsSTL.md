# Save as STL Macro for SolidWorks

## Description
This macro automates the process of saving the active SolidWorks document as an STL file, preserving the original document format and renaming the file extension accordingly.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - A SolidWorks document (not a drawing) must be open.
> - The document should not be a drawing file to proceed with the STL save operation.

## Results
> [!NOTE]
> - The active document is saved as an STL file in the same directory with the same file name but with an ".stl" extension.
> - A message box informs the user of the save location.

## Steps to Setup the Macro

1. **Open Your Document**:
   - Ensure that a SolidWorks document other than a drawing is open and active.

2. **Load and Run the Macro**:
   - Open the SolidWorks VBA editor by pressing (`Alt + F11`).
   - Insert a new module and paste the provided macro code.
   - Run the macro directly from the VBA editor or from within SolidWorks under **Tools** > **Macro** > **Run**.

3. **Using the Macro**:
   - Once run, if the conditions are met, the document will be saved as an STL file.
   - A confirmation message will display the path where the STL file was saved.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

' Main subroutine to save the active document as an STL file
Sub main()
    ' Declare and initialize SolidWorks application and the active document
    Dim swApp As SldWorks.SldWorks               ' SolidWorks application object
    Set swApp = Application.SldWorks
    Dim Part As ModelDoc2                        ' Active document object
    Set Part = swApp.ActiveDoc

    ' Check if a document is loaded
    If Part Is Nothing Then
        MsgBox "No document is open.", vbExclamation, "Error"
        Exit Sub ' Exit if no document is active
    End If

    ' Get the full file path of the active document
    Dim Path As String
    Path = Part.GetPathName

    ' Exit if the active document is a drawing
    If Part.GetType = swDocDRAWING Then
        MsgBox "The macro does not support drawing documents.", vbCritical, "Operation Aborted"
        Exit Sub
    End If

    ' Replace the current file extension with ".stl"
    Dim Extension As String
    Extension = Mid(Path, InStrRev(Path, ".")) ' Extract the current file extension
    Path = Replace(Path, Extension, ".stl")   ' Replace the extension with ".stl"

    ' Save the document as an STL file
    Dim longstatus As Long                     ' Status of the save operation
    longstatus = Part.SaveAs3(Path, 0, 0)      ' Save the file with STL extension

    ' Check the status of the save operation and inform the user
    If longstatus = 0 Then
        MsgBox "Failed to save the file as STL.", vbCritical, "Error"
    Else
        MsgBox "Saved as STL: " & Path, vbInformation, "Success"
    End If
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).