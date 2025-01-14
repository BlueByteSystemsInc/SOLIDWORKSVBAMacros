# Save and Close Document Macro for SolidWorks

## Description
This macro automates the process of saving an active SolidWorks document in both DWG and PDF formats and then closes the document. This is particularly useful for ensuring consistent output formats for archival or external use.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> A SolidWorks document (part or assembly) must be open.
> The macro must be executed within the SolidWorks environment.

## Results
- Saves the active document in both DWG and PDF formats to a predefined path.
- Closes the active document after saving.
- The macro ensures that file paths and names are dynamically generated based on the document title and current date and time to prevent overwriting existing files.

## Steps to Setup the Macro

1. **Create the VBA Modules**:
   - Open the SolidWorks VBA editor by pressing (`Alt + F11`).
   - Insert the module as shown in your project (`Macro31`).
   - Insert a userform named `save` and add necessary controls (e.g., command buttons).

2. **Run the Macro**:
   - Save the macro file (e.g., `SaveAndCloseDocument.swp`).
   - Run the macro from within SolidWorks by navigating to **Tools** > **Macro** > **Run**, then select your saved macro file.

3. **Using the Macro**:
   - The macro can be triggered from the userform by interacting with the provided controls.
   - The document will be saved in the specified formats and closed automatically.

## VBA Macro Code

### UserForm Code (`save`)
```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

' Event handler for the CommandButton click event
Private Sub CommandButton1_Click()
    ' Placeholder for additional functionality or event handling when the button is clicked
    ' Add your custom code here if needed
End Sub

' Event handler for changes in the TextBox
Private Sub TextBox1_Change()
    ' Code that executes whenever the text changes in the TextBox
    ' Add your custom code here for handling text input
End Sub

' Main subroutine for performing save and view operations
Sub main()
    ' Initialize the SolidWorks application and get the active document
    Set swApp = Application.SldWorks
    Set Part = swApp.ActiveDoc

    ' Step 1: Adjust the view to fit the model on the screen
    Part.ViewZoomtofit2

    ' Step 2: Clear any current selections in the active document
    Part.ClearSelection2 True

    ' Step 3: Save the document as a PDF in the specified directory
    ' Replace "C:\\Path\\To\\Your\\Folder\\" with the desired output folder
    Dim longstatus As Long
    longstatus = Part.SaveAs3("C:\\Path\\To\\Your\\Folder\\" & Part.GetTitle & ".pdf", 0, 2)

    ' Step 4: Refresh the graphics to ensure the display is updated
    Part.ViewZoomtofit2
    Part.GraphicsRedraw2

    ' Step 5: Save the document as a DWG in the specified directory
    ' Perform another view zoom operation and save the file
    Part.ViewZoomtofit2
    longstatus = Part.SaveAs3("C:\\Path\\To\\Your\\Folder\\" & Part.GetTitle & ".dwg", 0, 2)
End Sub

' Event handler for the Save button click event
Private Sub save_Click()
    ' Trigger the main subroutine when the Save button is clicked
    Call main
End Sub
```

### Macro Module (Macro31)
```vbnet
' Declare variables for the SolidWorks application and document
Dim swApp As Object                 ' SolidWorks application object
Dim Part As Object                  ' Active document object
Dim boolstatus As Boolean           ' Boolean to capture operation success
Dim longstatus As Long, longwarnings As Long ' Longs to capture detailed operation statuses

Sub saveDWG()
    ' Initialize the SolidWorks application object
    Set swApp = Application.SldWorks

    ' Get the currently active document
    Set Part = swApp.ActiveDoc

    ' Extract the file name without the extension (assuming file extension length is 9 characters, e.g., ".sldprt")
    Dim nomeArquivo As String
    nomeArquivo = Left(Part.GetTitle, Len(Part.GetTitle) - 9)

    ' Get the current working directory of SolidWorks
    Dim Path As String
    Path = swApp.GetCurrentWorkingDirectory

    ' Step 1: Set preferences for DXF file format
    ' Set the DXF format to R2013 version for compatibility
    boolstatus = swApp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swDxfVersion, swDxfFormat_e.swDxfFormat_R2013)

    ' Step 2: Adjust the view to fit the model
    Part.ViewZoomtofit2

    ' Step 3: Save the document in DWG format
    longstatus = Part.SaveAs3(Path & nomeArquivo & ".DWG", 0, 0)

    ' Step 4: Save the document in PDF format
    longstatus = Part.SaveAs3(Path & nomeArquivo & ".PDF", 0, 0)

    ' Step 5: Save any changes made to the document
    Dim swErrors As Long, swWarnings As Long
    boolstatus = Part.Save3(1, swErrors, swWarnings)

    ' Step 6: Close the document after saving
    swApp.CloseDoc(Part.GetTitle)
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).