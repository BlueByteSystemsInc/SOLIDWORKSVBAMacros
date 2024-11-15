# Export Active Document as PARASOLID File

## Description
This macro exports the active SolidWorks document (part or assembly) as a PARASOLID file (.X_T) to the same directory where the original file is saved. The exported file will automatically be named with the same name as the active document but with the `.X_T` extension. This macro is especially useful for quickly saving parts and assemblies as PARASOLID files for use in other software or for sharing with colleagues.


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
> - The macro will save the active document as a `.X_T` file in the same directory.
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
    
    ' Exit if the active document is a drawing (since PARASOLID export is not supported for drawings)
    If Part.GetType = swDocDRAWING Then
        Exit Sub
    End If
    
    ' Prepare the path for the PARASOLID file by replacing the extension
    Dim Extension As String
    Extension = Mid(Path, InStrRev(Path, "."))
    Path = Replace(Path, Extension, ".X_T")
    Extension = ".X_T"

    ' Export the file as PARASOLID
    longstatus = Part.SaveAs3(Path, 0, 0)

    ' Notify the user about the saved file location
    MsgBox "Saved " & Path, vbInformation
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).
