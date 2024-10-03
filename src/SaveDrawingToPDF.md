# Save Drawing as PDF in SolidWorks

## Description
This macro instantly saves the active part or assembly drawing as a PDF file. The PDF document is saved in the same folder as the drawing with the same name. This macro works best when assigned to a keyboard shortcut, making it easy to quickly export drawings to PDF format without manually navigating through the menus.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - The active document must be a saved drawing file.
> - The drawing should have at least one sheet.
> - Ensure the drawing is open and active before running the macro.

## Results
> [!NOTE]
> - All sheets of the active drawing are exported as a single PDF file.
> - The PDF is saved in the same location as the drawing file with the same name.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit
Sub main()

    ' Declare and initialize necessary SolidWorks objects
    Dim swApp As SldWorks.SldWorks             ' SolidWorks application object
    Dim swModel As SldWorks.ModelDoc2          ' Active document object
    Dim swModelDocExt As SldWorks.ModelDocExtension  ' Model document extension object
    Dim swExportData As SldWorks.ExportPdfData ' PDF export data object
    Dim boolstatus As Boolean                  ' Status of export operation
    Dim filename As String                     ' Filename of the PDF to be saved
    Dim lErrors As Long                        ' Variable to capture errors during save
    Dim lWarnings As Long                      ' Variable to capture warnings during save

    ' Initialize SolidWorks application and get the active document
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    ' Check if a document is currently open in SolidWorks
    If swModel Is Nothing Then
        MsgBox "No active document found. Please open a drawing and try again.", vbCritical, "No Active Document"
        Exit Sub
    End If

    ' Check if the active document is a drawing
    If swModel.GetType <> swDocDRAWING Then
        MsgBox "This macro only works with drawing files. Please open a drawing and try again.", vbCritical, "Invalid Document Type"
        Exit Sub
    End If

    ' Get the extension object of the active drawing document
    Set swModelDocExt = swModel.Extension

    ' Initialize the PDF export data object
    Set swExportData = swApp.GetExportFileData(swExportPDFData)

    ' Get the file path of the active drawing
    filename = swModel.GetPathName

    ' Check if the drawing has been saved
    If filename = "" Then
        MsgBox "The drawing must be saved before exporting to PDF. Please save the drawing and try again.", vbCritical, "Save Required"
        Exit Sub
    End If

    ' Modify the file path to save as PDF (replace extension with .PDF)
    filename = Strings.Left(filename, Len(filename) - 6) & "PDF"

    ' Set the export option to include all sheets in the drawing
    boolstatus = swExportData.SetSheets(swExportData_ExportAllSheets, 1)

    ' Save the drawing as a PDF using the specified filename and export data
    boolstatus = swModelDocExt.SaveAs(filename, 0, 0, swExportData, lErrors, lWarnings)

    ' Check if the export was successful and display appropriate message
    If boolstatus Then
        MsgBox "Drawing successfully saved as PDF:" & vbNewLine & filename, vbInformation, "Export Successful"
    Else
        MsgBox "Save as PDF failed. Error code: " & lErrors, vbExclamation, "Export Failed"
    End If

End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).
