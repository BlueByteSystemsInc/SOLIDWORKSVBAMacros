# Save Part as eDrawings File (.eprt) Macro

## Description
This macro saves the active SolidWorks part as an **eDrawings Part (.eprt)** file in the same directory as the original part file. If a file with the same name already exists, the macro increments a counter to avoid overwriting, creating a unique file name as necessary.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - The active document must be a part file.
> - Ensure that SolidWorks is open with a part file loaded before running this macro.

## Results
> [!NOTE]
> - The part will be saved as a `.eprt` file in the same directory as the original part file.
> - If a file with the same name already exists, the macro will increment a counter to avoid overwriting the existing file.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Dim swModel      As SldWorks.ModelDoc2
Dim Dateiname    As String
Dim Error        As Long
Dim Warnung      As Long
Dim nRetval      As Long
Dim reurn_value  As Boolean
Dim OpenDoc      As Object
Dim ZählerInt    As Integer
Dim ZählerStr    As String
Dim fertig       As Boolean
Dim Pfad         As String
Dim Originalname As String
Dim Speicherort  As String

Sub main()
    ' Initialize SolidWorks application and active document
    Set swApp = CreateObject("SldWorks.Application")
    Set swModel = swApp.ActiveDoc
    Set OpenDoc = swApp.ActiveDoc()
    
    ' Set initial variables
    ZählerInt = 1
    fertig = False
    Originalname = OpenDoc.GetTitle
    
    ' Start loop to check for existing files and increment file name if needed
    Do
        Speicherort = OpenDoc.GetPathName
        Name = OpenDoc.GetTitle
        Name = Left(Name, (Len(Name) - 7)) ' Remove file extension from original name
        ZählerStr = Str(ZählerInt)
        ZählerStr = Right(ZählerStr, (Len(ZählerStr) - 1))
        Pfad = Name + ".eprt"
        Speicherort = Left(Speicherort, Len(Speicherort) - Len(Originalname)) + Pfad
        Pfad = Dir(Speicherort, vbNormal)
        
        ' If no file exists with this name, exit loop
        If Pfad = "" Then
            fertig = True
        Else
            ZählerInt = ZählerInt + 1 ' Increment counter for file name
        End If
    Loop Until fertig = True
    
    ' Save the file as eDrawings Part (.eprt)
    return_value = swModel.SaveAs4(Speicherort, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Error, Warnung)
    
    ' Display message if there was a problem saving the file
    If reurn_value = True Then
        nRetval = swApp.SendMsgToUser2("Problems saving file.", swMbWarning, swMbOk)
    End If
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).