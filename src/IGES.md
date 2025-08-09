# Export Active Document as IGES File
<video src="../images/Save_as_IGES.mkv" autoplay muted controls style="width: 100%; border-radius: 12px;"></video>
## Description
This macro exports the active SolidWorks document (part or assembly) as an IGES file (.igs) to the same directory where the original file is saved. The exported IGES file will have the same name as the active document but with the `.igs` extension. This macro is useful for quickly saving parts and assemblies in IGES format for compatibility with other CAD software.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - The active document must be a part or assembly.
> - This macro does not work for drawing files.
> - Ensure the document is saved before running the macro, as the file will be exported in the same directory.

## Results
> [!NOTE]
> - The macro will save the active document as an `.igs` file in the same directory.
> - A message box will appear confirming the location of the saved file.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Dim swApp As SldWorks.SldWorks
Dim Part As ModelDoc2
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long

Sub main()

    ' Initialize SolidWorks application and active document
    Set swApp = Application.SldWorks
    Set Part = swApp.ActiveDoc
    
    ' Exit if no document is active
    If Part Is Nothing Then Exit Sub
    
    ' Get the path of the active document
    Dim Path As String
    Path = Part.GetPathName
    
    ' Exit if the active document is a drawing (since IGES export is not supported for drawings)
    If Part.GetType = swDocDRAWING Then
        Exit Sub
    End If
    
    ' Prepare the path for the IGES file by replacing the extension
    Dim Extension As String
    Extension = Mid(Path, InStrRev(Path, "."))
    Path = Replace(Path, Extension, ".igs")
    Extension = ".igs"

    ' Export the file as IGES
    longstatus = Part.SaveAs3(Path, 0, 0)

    ' Notify the user about the saved file location
    MsgBox "Saved " & Path, vbInformation
End Sub
```
You can download the macro from [here](../images/Save_as_IGES.swp)

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).