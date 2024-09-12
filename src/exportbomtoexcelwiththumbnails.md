# Export BOM with Thumbnail Preview in SOLIDWORKS

## Macro Description

This VBA macro automates the process of exporting a Bill of Materials (BOM) from a SOLIDWORKS drawing into an Excel sheet, while also adding a thumbnail image preview of the parts. The macro is designed to enhance the visualization of BOMs, allowing users to include part thumbnails directly in the Excel output. This can be extremely useful for teams needing a detailed and visual breakdown of the parts for purchasing, inventory, or assembly processes.

## VBA Macro Code

```vbnet
' All rights reserved to Blue Byte Systems Inc.
' Blue Byte Systems Inc. does not provide any warranties for macros.
' Pre-conditions: BOM pre-selected.
' Results: BOM created in Excel with thumbnail preview.

' Define the width and height of the thumbnail (in pixels)
Dim Width As Long ' in pixels
Dim Height As Long ' in pixels

Dim swApp As Object
Dim swModel As Object
Dim swTableAnnotation As Object
Dim exApp As Object
Dim exWorkbook As Object
Dim exWorkSheet As Object
Dim swSelectionManager As Object

' Enums for SolidWorks document types, Excel alignment, and table header positions
Public Enum swDocumentTypes_e
    swDocDRAWING = 3
End Enum

Public Enum xlTextAlignment
    xlCenter = -4108
End Enum

Public Enum swTableHeaderPosition_e
    swTableHeader_Top = 1
    swTableHeader_Bottom = 2
    swTableHeader_None = 0
End Enum

Public Enum swSelectType_e
    swSelBOMS = 97
End Enum

Sub Main()
    ' Set the thumbnail dimensions
    Width = 21
    Height = 60

    ' Get a pointer to the SolidWorks application
    Set swApp = Application.SldWorks

    ' Get the active document
    Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then
        swApp.SendMsgToUser "There is no active document"
        End
    End If

    ' Get the selection manager
    Set swSelectionManager = swModel.SelectionManager

    ' Get the count of selected objects
    Dim Count As Long
    Count = swSelectionManager.GetSelectedObjectCount2(-1)

    ' If no BOM is selected, exit the macro
    If Count = 0 Then
        swApp.SendMsgToUser "You have not selected any bill of materials!"
        Exit Sub
    End If

    ' Traverse the selection and process all selected bill of materials
    For i = 1 To Count
        If swSelectionManager.GetSelectedObjectType3(i, -1) = SwConst.swSelectType_e.swSelANNOTATIONTABLES Then
            Set swTableAnnotation = swSelectionManager.GetSelectedObject6(i, -1)
            Dim Ret As String
            Ret = SaveBOMInExcelWithThumbNail(swTableAnnotation)
            If Ret = "" Then
                Debug.Print "Success: " & swTableAnnotation.GetAnnotation.GetName
                swApp.SendMsgToUser "The selected BOM has been exported with thumbnail preview to Excel."
            Else
                swApp.SendMsgToUser "Macro failed to export!"
            End If
        End If
    Next i

End Sub

' Function to save BOM to Excel with thumbnail preview
Public Function SaveBOMInExcelWithThumbNail(ByRef swTableAnnotation As Object) As String
    ' Initialize Excel application
    Set exApp = CreateObject("Excel.Application")
    If exApp Is Nothing Then
        SaveBOMInExcelWithThumbNail = "Unable to initialize the Excel application"
        Exit Function
    End If
    exApp.Visible = True

    ' Create a new workbook and worksheet
    Set exWorkbook = exApp.Workbooks.Add
    Set exWorkSheet = exWorkbook.ActiveSheet
    If exWorkSheet Is Nothing Then
        SaveBOMInExcelWithThumbNail = "Unable to get the active sheet"
        Exit Function
    End If

    ' If the BOM has no rows, return an error
    If swTableAnnotation.RowCount = 0 Then
        SaveBOMInExcelWithThumbNail = "BOM has no rows!"
        Exit Function
    End If

    Dim swBOMTableAnnotation As BomTableAnnotation
    Set swBOMTableAnnotation = swTableAnnotation

    ' Set the column width
    exWorkSheet.Columns(1).ColumnWidth = Width

    ' Set the header row index based on the BOM header position
    Dim HeaderRowIndex As Long
    Dim swHeaderIndex As Integer
    swHeaderTable = swTableAnnotation.GetHeaderStyle
    If swHeaderTable = swTableHeaderPosition_e.swTableHeader_Bottom Then
        swHeaderIndex = swTableAnnotation.RowCount
    Else
        swHeaderIndex = 1
    End If

    ' Traverse through each row in the BOM table
Skipper:
    For i = 0 To swTableAnnotation.RowCount - 1
        ' Skip hidden rows
        If swTableAnnotation.RowHidden(i) Then GoTo Skipper

        ' Add preview image
        Dim swComponents As Variant
        swComponents = swBOMTableAnnotation.GetComponents(i)
        If Not IsEmpty(swComponents) Then
            Dim swComponent As Object
            Set swComponent2 = swComponents(0)
            Dim swComponentModel As Object
            Set swComponentModel = swComponent2.GetModelDoc2
            If Not swComponentModel Is Nothing Then
                swComponentModel.Visible = True
                Dim imagePath As String
                imagePath = Environ("TEMP") + "\tempBitmap.jpg"
                swComponentModel.ViewZoomtofit2
                Dim saveRet As Boolean
                Dim er As Long
                Dim wr As Long
                saveRet = swComponentModel.Extension.SaveAs(imagePath, 0, 0, Nothing, er, wr)
                If er + wr > 0 Then
                    SaveBOMInExcelWithThumbNail = "An error has occurred while trying to save the thumbnail of " & swModel.GetTitle & " to the local temp folder. The macro will exit now."
                    Exit Function
                End If
                swComponentModel.Visible = False
                exWorkSheet.Rows(i + 1).RowHeight = Height
                InsertPictureInRange exWorkSheet, imagePath, exWorkSheet.Range("A" & i + 1 & ":A" & i + 1)
            End If
        End If

        ' Populate Excel sheet with BOM table data
        For j = 0 To swTableAnnotation.ColumnCount - 1
            If swTableAnnotation.ColumnHidden(j) Then GoTo Skipper
            exWorkSheet.Cells(i + 1, j + 2).Value = swTableAnnotation.DisplayedText(i, j)
        Next j
    Next i

    ' Bold the header row
    For j = 2 To swTableAnnotation.ColumnCount + 1
        exWorkSheet.Cells(swHeaderIndex, j).Font.Bold = True
    Next j

    ' Auto-fit the columns and center align the content
    Dim r As Object
    Set r = exWorkSheet.Range(exWorkSheet.Cells(1, 2), exWorkSheet.Cells(swTableAnnotation.RowCount + 1, swTableAnnotation.ColumnCount + 1))
    r.Columns.AutoFit
    r.HorizontalAlignment = xlTextAlignment.xlCenter
    r.VerticalAlignment = xlTextAlignment.xlCenter

End Function

' Subroutine to insert a picture in a specific range in Excel
Sub InsertPictureInRange(ActiveSheet As Object, PictureFileName As String, TargetCells As Object)
    ' Inserts a picture and resizes it to fit the TargetCells range
    Dim p As Object, t As Double, l As Double, w As Double, h As Double
    If TypeName(ActiveSheet) > "Worksheet" Then Exit Sub
    If Dir(PictureFileName) = "" Then Exit Sub
    ' Import picture
    Set p = ActiveSheet.Pictures.Insert(PictureFileName)
    ' Determine positions
    With TargetCells
        t = .Top
        l = .Left
        w = .Offset(0, .Columns.Count).Left - .Left
        h = .Offset(.Rows.Count, 0).Top - .Top
    End With
    ' Position picture
    With p
        .Top = t
        .Left = l
        .Width = w
        .Height = h
    End With
    Set p = Nothing
End Sub
```


## System Requirements
To run this VBA macro, ensure that your system meets the following requirements:

- SOLIDWORKS Version: SOLIDWORKS 2017 or later
- VBA Environment: Pre-installed with SOLIDWORKS (Access via Tools > Macro > New or Edit)
- Operating System: Windows 7, 8, 10, or later
- Microsoft Excel

>[!NOTE]
> Pre-conditions
> - A Bill of Materials (BOM) must be pre-selected in the SOLIDWORKS drawing.
> - Excel must be installed on the machine.

>[!NOTE]
> Post-conditions
> The BOM will be exported into an Excel file with a part thumbnail preview inserted for each row.

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).