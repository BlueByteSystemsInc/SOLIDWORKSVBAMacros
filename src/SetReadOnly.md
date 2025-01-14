# Toggle Read-Only Status of SolidWorks Files Macro

## Description
This macro allows users to toggle the read-only attribute for all files currently loaded in SolidWorks. This is especially useful for preventing unintentional edits or when transitioning files to a protected status in a collaborative environment.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - SolidWorks must be open with one or more documents loaded.

## Results
> [!NOTE]
> - The macro will prompt the user to set all loaded files to read-only or revert them to normal status.
> - Changes will affect all files currently open in the SolidWorks session.

## Steps to Setup the Macro

### 1. **Prepare SolidWorks**:
   - Open the documents you want to modify the read-only status for in SolidWorks.

### 2. **Run the Macro**:
   - Execute the `main` subroutine.
   - Respond to the prompt depending on whether you want to enable or disable read-only status.

### 3. **Review Changes**:
   - Verify the read-only status of the files after running the macro to ensure it has been applied or removed as expected.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

' Declare variables for SolidWorks application and loaded models
Dim swApp As SldWorks.SldWorks          ' SolidWorks application object
Dim swModel As SldWorks.ModelDoc2       ' SolidWorks model object
Dim vModels As Variant                  ' Array to store all loaded models
Dim count As Long                       ' Count of loaded models
Dim index As Long                       ' Index for iterating through models
Dim PathName As String                  ' Path of the model file

' Main subroutine
Sub main()
    ' Initialize SolidWorks application object
    Set swApp = Application.SldWorks

    ' Prompt user with a question about setting files to Read-Only
    Dim Message As String
    Dim Style As VbMsgBoxStyle
    Dim Title As String
    Dim Response As VbMsgBoxResult
    
    count = swApp.GetDocumentCount  ' Get the count of loaded models
    Message = "Set the files loaded in SolidWorks to read-only?" & vbNewLine & vbNewLine & _
              "Files loaded in the current SolidWorks session: " & count
    Style = vbYesNo + vbQuestion + vbDefaultButton2
    Title = "      -= SolidWorks 2014 =-       "
    
    ' Display the message box to the user
    Response = MsgBox(Message, Style, Title)

    ' If the user selects Yes
    If Response = vbYes Then
        MsgBox "Starting Read-Only setting" & vbNewLine & vbNewLine & _
               "Files loaded in the current SolidWorks session: " & count, vbOKOnly, "SolidWorks 2014"
        
        ' Retrieve all loaded models
        vModels = swApp.GetDocuments
        For index = LBound(vModels) To UBound(vModels)
            Set swModel = vModels(index)            ' Access each model
            Path_File = swModel.GetPathName         ' Get the file path of the model
            
            ' Set the file to Read-Only
            SetAttr Path_File, vbReadOnly
        Next index
        
    ' If the user selects No
    Else
        MsgBox "        Removing the read-only attribute        ", vbOKOnly, "SolidWorks 2014"
        
        ' Retrieve all loaded models
        vModels = swApp.GetDocuments
        For index = LBound(vModels) To UBound(vModels)
            Set swModel = vModels(index)            ' Access each model
            Path_File = swModel.GetPathName         ' Get the file path of the model
            
            ' Remove the Read-Only attribute
            SetAttr Path_File, vbNormal
        Next index
        
        Exit Sub
    End If

End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).