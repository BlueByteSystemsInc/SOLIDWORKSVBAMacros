# Save Assembly as Part (All Components) Macro

## Description
This macro automates the process of saving a SolidWorks assembly as a single part file that includes all components. It is particularly useful for simplifying assemblies for external use or reducing file size.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer  
- **Operating System**: Windows 7 or later  

## Pre-Conditions
> [!NOTE]
> - An assembly document must be open in SolidWorks.
> - The assembly must contain at least one part.

## Results
> [!NOTE]
> - The assembly will be saved as a single part file with all components included.

## Steps to Setup the Macro

### 1. **Prepare the Assembly**:
   - Open the desired assembly in SolidWorks.

### 2. **Run the Macro**:
   - Execute the macro. The assembly will be saved as a part file with all components included. The new file will be created in the same directory as the original assembly with the `.SLDPRT` extension.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

' Declare SolidWorks objects and variables
Dim swApp               As SldWorks.SldWorks               ' SolidWorks application object
Dim swModel             As SldWorks.ModelDoc2              ' Active document object
Dim swModelDocExt       As SldWorks.ModelDocExtension      ' Extension object for model operations
Dim FilePath            As String                          ' Full file path of the active document
Dim PathSize            As Long                            ' Length of the file path string
Dim PathNoExtension     As String                          ' File path without extension
Dim NewFilePath         As String                          ' Path for the new part file
Dim nErrors             As Long                            ' Variable to store errors during save operation
Dim nWarnings           As Long                            ' Variable to store warnings during save operation

' Main subroutine to perform the save operation
Sub main()

    ' Initialize SolidWorks application and active document
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set swModelDocExt = swModel.Extension

    ' Get the full path of the active assembly file
    FilePath = swModel.GetPathName
    PathSize = Strings.Len(FilePath)

    ' Remove the extension from the file path to prepare for the new part file
    PathNoExtension = Strings.Left(FilePath, PathSize - 6)

    ' Create the new file path with ".SLDPRT" extension
    NewFilePath = PathNoExtension & "SLDPRT"

    ' Set the user preference to save the assembly as a part with all components included
    swApp.SetUserPreferenceIntegerValue swSaveAssemblyAsPartOptions, swSaveAsmAsPart_AllComponents

    ' Save the assembly as a part file
    swModelDocExt.SaveAs NewFilePath, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, nErrors, nWarnings

    ' Provide feedback to the user
    If nErrors = 0 And nWarnings = 0 Then
        MsgBox "Assembly successfully saved as part: " & NewFilePath, vbInformation, "Success"
    Else
        MsgBox "Save operation completed with errors or warnings." & vbCrLf & _
               "Errors: " & nErrors & vbCrLf & "Warnings: " & nWarnings, vbExclamation, "Attention"
    End If

End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).