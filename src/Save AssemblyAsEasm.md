# Save Assembly as eDrawings (.easm) File

## Description
This macro saves the currently active assembly as an eDrawings (.easm) file in the same folder as the original assembly file. It automatically assigns the eDrawings file the same name as the assembly and checks if the file already exists in the directory to avoid overwriting.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - The active document must be an assembly.
> - The assembly should be saved to ensure the path can be derived correctly.
> - Ensure that the assembly and its components are not read-only or open in another application.

## Results
> [!NOTE]
> - The active assembly is saved as an eDrawings (.easm) file in the same folder as the assembly file.
> - The eDrawings file is saved with the same name as the assembly.
> - If a file with the same name already exists, the macro appends a counter value to the filename to avoid overwriting.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

' Declare necessary SolidWorks and macro variables
Dim swApp As SldWorks.SldWorks                    ' SolidWorks application object
Dim swModel As SldWorks.ModelDoc2                 ' Active document object (assembly)
Dim Name As String                           ' File name for the eDrawings file
Dim Error As Long, Warnung As Long                ' Error and warning variables
Dim nRetval As Long                               ' Return value for file saving operations
Dim return_value As Boolean                       ' Boolean variable to check saving status
Dim OpenDoc As Object                             ' Object for the currently active document
Dim ZählerInt As Integer                         ' Counter to avoid overwriting files
Dim ZählerStr As String                          ' String representation of the counter
Dim fertig As Boolean                             ' Boolean variable to indicate completion
Dim Pfad As String                                ' Path for saving the eDrawings file
Dim Originalname As String                        ' Original file name of the assembly
Dim Speicherort As String                         ' Full path for saving the eDrawings file

' --------------------------------------------------------------------------
' Main subroutine to save the assembly as an eDrawings (.easm) file
' --------------------------------------------------------------------------
Sub main()

    ' Initialize SolidWorks application and get the active document
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set OpenDoc = swApp.ActiveDoc()

    ' Check if there is an active document open
    If swModel Is Nothing Then
        MsgBox "No active document found. Please open an assembly and try again.", vbCritical, "No Active Document"
        Exit Sub
    End If

    ' Check if the active document is an assembly
    If swModel.GetType <> swDocASSEMBLY Then
        MsgBox "This macro only works on assemblies. Please open an assembly and try again.", vbCritical, "Invalid Document Type"
        Exit Sub
    End If

    ' Initialize counter and variables
    ZählerInt = 1
    fertig = False
    Originalname = OpenDoc.GetTitle  ' Get the original name of the active document

    ' Loop to determine a unique file name if a file with the same name exists
    Do
        ' Get the full path of the currently active document
        Speicherort = OpenDoc.GetPathName

        ' Extract the base name of the document (remove extension)
        Name = OpenDoc.GetTitle
        Name = Left(Name, (Len(Name) - 7))

        ' Convert counter to string and append to the file name if necessary
        ZählerStr = Str(ZählerInt)
        ZählerStr = Right(ZählerStr, (Len(ZählerStr) - 1))
        Pfad = Name + ".easm"
        Speicherort = Left(Speicherort, Len(Speicherort) - Len(Originalname)) + Pfad

        ' Check if a file with the same name already exists in the directory
        Pfad = Dir(Speicherort, vbNormal)
        If Pfad = "" Then
          fertig = True
        Else
          ZählerInt = ZählerInt + 1
        End If

    Loop Until fertig = True  ' Repeat until a unique file name is found

    ' Save the assembly as an eDrawings file (.easm) in the specified location
    return_value = swModel.SaveAs4(Speicherort, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Error, Warnung)
    If return_value = True Then
        nRetval = swApp.SendMsgToUser2("Problems saving file.", swMbWarning, swMbOk)
    End If

End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).
