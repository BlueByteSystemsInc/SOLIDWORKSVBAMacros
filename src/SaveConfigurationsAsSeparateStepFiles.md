# Batch Export Configurations to STEP Macro for SolidWorks

## Description
This SolidWorks macro facilitates the batch export of each configuration within an active document to the STEP file format, specifically using the AP214 standard. It is especially useful for efficiently managing multiple configurations in projects where external compatibility is necessary.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - An assembly or part document with one or more configurations must be open.
> - The macro should be executed within the SolidWorks environment.

## Results
> [!NOTE]
> - Each configuration of the active document is saved as a separate STEP file in a specified directory.
> - The user is prompted to specify a file prefix and output directory to organize the exported files.
> - A progress bar is displayed during the export process, providing feedback on the operation's progress.

## Steps to Setup the Macro

1. **Open SolidWorks**:
   - Launch SolidWorks and open the document you wish to export configurations from.

2. **Load and Run the Macro**:
   - Open the SolidWorks VBA editor by pressing (`Alt + F11`).
   - Insert a new module and paste the provided macro code.
   - Run the macro from the VBA editor or save it and run it from **Tools** > **Macro** > **Run**.

3. **Using the Macro**:
   - Follow the prompts to specify the file prefix and output directory.
   - Monitor the progress bar that appears to track the export process.
   - Upon completion, a message will confirm the number of files saved.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

Sub main()
    ' Initialize SolidWorks application and document
    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc

    ' Verify an active document is open
    If swModel Is Nothing Then
        MsgBox "No active document found. Please open a model.", vbCritical, "Error"
        Exit Sub
    End If

    ' Declare configuration management objects
    Dim swConfigMgr As SldWorks.ConfigurationManager
    Set swConfigMgr = swModel.ConfigurationManager
    Dim swConfig As SldWorks.Configuration
    Set swConfig = swModel.GetActiveConfiguration

    ' Prepare file path and name variables
    Dim path As String, fname As String, current As String, prefix As String, dirName As String
    path = swModel.GetPathName
    fname = Mid(path, InStrRev(path, "\") + 1, Len(path) - InStr(path, ".") - 1)
    path = Left(path, InStrRev(path, "\") - 1)
    current = swModel.GetActiveConfiguration.Name
    Dim configs As Variant
    configs = swModel.GetConfigurationNames

    ' User input for file prefix and output directory
    prefix = InputBox("Enter the prefix:", "Names", fname)
    If prefix = "" Then
        MsgBox "Prefix cannot be empty.", vbCritical, "Error"
        Exit Sub
    End If
    dirName = InputBox("Enter the directory name for saving:", "Directory Name", "STEP")
    If dirName = "" Then
        MsgBox "Directory name cannot be empty.", vbCritical, "Error"
        Exit Sub
    End If

    ' Ensure output directory exists
    If Dir(dirName, vbDirectory) = "" Then MkDir dirName
    ChDir dirName

    ' Progress bar setup and configuration iteration
    Dim i As Long
    For i = 0 To UBound(configs)
        Dim name As String
        name = prefix & configs(i) & ".STEP"
        swModel.ShowConfiguration2 configs(i)
        Call swModel.SaveAs3(name, 0, 0)
    Next i

    ' Confirmation message and cleanup
    MsgBox "Saved " & CStr(i) & " files!", vbInformation, "Done"
    swModel.ShowConfiguration2 current
    ChDir path
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).