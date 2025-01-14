# Import 3D Points into a SolidWorks Part Model

## Description
This macro imports 3D points from a text file into a 3D sketch in a SolidWorks part model. It ensures that the active document is a valid part model and creates a new part if none is open.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - A blank part model should be active before running this macro.
> - The text file containing 3D points must be accessible and formatted correctly.

## Results
> [!NOTE]
> - Imports 3D points from a text file into a new 3D sketch in the part model.
> - If no part is open, a new part model will be created using the default template.

## STEPS to Setup the Macro

1. **Create the UserForm**:
- Open the VBA editor in SolidWorks by pressing (`Alt + F11`).
- In the Project Explorer, right-click on the project (e.g., `Macro1`) and select **Insert** > **UserForm**.
    - Rename the form to `UserForm1`.
    - Design the form with the following:
        - Add a Label at the top: **Select file to open.**
        - Add a ListBox named **ListBox1** for displaying file contents.
        - Add two buttons:
            - Import: Set Name = `CmdImport` and Caption = `Import`.
            - Close: Set Name = `CmdClose` and Caption = `Close`.

2. **Add VBA Code**:
   - Copy the **Macro Code** provided below into the module.
   - Copy the **UserForm Code** into the `UserForm1` code-behind.

3. **Save and Run the Macro**:
   - Save the macro file (e.g., `3DPoints.swp`).
   - Run the macro by going to **Tools** > **Macro** > **Run** in SolidWorks, then select your saved macro.

4. **Import 3D Points**:
   - The macro will open the 3D Point Import UserForm.
   - Follow these steps:
        1. Click **Import** and select the text file containing 3D points.
        2. Ensure the text file is formatted as comma-separated values, e.g.:
            ```
            0.0,0.0,0.0  
            1.0,1.0,1.0  
            2.0,2.0,2.0  
            ```
        3. The macro will insert the points into a new 3D Sketch in the part model.
        4. Click **Close** to exit the UserForm.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

Dim swApp As Object                ' SolidWorks application object
Dim Part As Object                 ' Active document object (part model)
Dim ModelDocExt As Object          ' ModelDocExtension object for extended functionalities
Dim boolstatus As Boolean          ' Boolean status for operations
Dim longstatus As Long             ' Long status for capturing operation results
Dim longwarnings As Long           ' Long warnings for capturing warnings

' Enumeration for SolidWorks document types
Public Enum swDocumentTypes_e
    swDocNONE = 0                  ' No document type
    swDocPART = 1                  ' Part document type
    swDocASSEMBLY = 2              ' Assembly document type
    swDocDRAWING = 3               ' Drawing document type
    swDocSDM = 4                   ' Solid data manager document type
End Enum

Sub main()
    Dim FileTyp As Integer         ' Type of the active file
    Dim MassStatus As Long         ' Status of the mass properties

    ' Initialize SolidWorks application
    Set swApp = Application.SldWorks
    
    ' Get the active document
    Set Part = swApp.ActiveDoc
    
    ' Check if a document is open
    If Not Part Is Nothing Then
        FileTyp = Part.GetType     ' Get document type
        
        ' Check if the document is a part model
        If FileTyp = swDocPART Then
            Set Part = swApp.ActiveDoc
            Set ModelDocExt = Part.Extension
            
            ' Get mass properties of the part
            Dim MassValue As Variant
            MassValue = ModelDocExt.GetMassProperties(1, MassStatus)
            
            ' Check if the part is blank (no mass)
            If MassStatus = 2 Then
                PointImport.Show   ' Show user form for point import
            Else
                MsgBox "Part model has mass. Please start with a blank part model.", vbExclamation, "Invalid Part Model"
            End If
        Else
            MsgBox "Current document is not a part model. Please start with a blank part model.", vbExclamation, "Invalid Document Type"
        End If
    Else
        ' Load a new part using the default template
        Dim DefaultPart As String
        DefaultPart = swApp.GetUserPreferenceStringValue(swDefaultTemplatePart)
        Set Part = swApp.NewDocument(DefaultPart, 0, 0, 0)
        
        ' Check if the new part was created successfully
        If Not Part Is Nothing Then
            Set ModelDocExt = Part.Extension
            UserForm1.Show        ' Show user form for user input
        Else
            MsgBox "Could not automatically load part. Please start with a blank part model.", vbExclamation, "Part Creation Failed"
            MsgBox "File " & DefaultPart & " not found", vbCritical, "Template Not Found"
        End If
    End If
End Sub
```

## VBA UserForm Code
```vbnet
Option Explicit

Dim WorkDirectory As String
Dim FileName As String

'------------------------------------------------------------------------------  
' Add files to list  
'------------------------------------------------------------------------------  
Private Sub AddToFileList(Extension)
    ListBoxFiles.Clear
    FileName = Dir(WorkDirectory + Extension)   ' Retrieve file list  
    Do While FileName <> ""
        ListBoxFiles.AddItem FileName
        FileName = Dir
    Loop
End Sub

' Close Button Event  
Private Sub CommandClose_Click()
    End
End Sub

' Import Button Event  
Private Sub CommandImport_Click()
    Dim Source As String
    Dim ReadLine As String
    Dim PntCnt As Long
    Dim DimX As Double, DimY As Double, DimZ As Double
    Dim Axis1 As Double, Axis2 As Double, Axis3 As Double
    
    ' Start a 3D Sketch  
    Part.Insert3DSketch
    PntCnt = 0
    Source = WorkDirectory & ListBoxFiles.List(ListBoxFiles.ListIndex, 0)
    
    Open Source For Input As #1   ' Open the source file  
    Do While Not EOF(1)
        Input #1, ReadLine
        ' Check for lines containing "HITS"  
        If Right$(UCase(ReadLine), 4) = "HITS" Then
            Input #1, DimX, DimY, DimZ, Axis1, Axis2, Axis3
            PntCnt = PntCnt + 1
            LabelProcessing.Caption = "Processing:" & Chr(13) & "Point # " & CStr(PntCnt)
            Me.Repaint   ' Update form UI  
            Part.CreatePoint2 DimX, DimY, DimZ
        End If
    Loop
    
EndRead:
    Close #1   ' Close the file  
    LabelProcessing.Caption = "Processed:" & Chr(13) & CStr(PntCnt) & " points."
    Part.SketchManager.InsertSketch True
    Part.ClearSelection2 True
    Part.ViewZoomtofit
End Sub

' ListBox File Click Event  
Private Sub ListBoxFiles_Click()
    CommandImport.Enabled = True
End Sub

' UserForm Initialization  
Private Sub UserForm_Initialize()
    CommandImport.Enabled = False
    WorkDirectory = swApp.GetCurrentWorkingDirectory
    AddToFileList "*.txt"
    If ListBoxFiles.ListCount < 1 Then
        MsgBox "No data files found.", vbExclamation, "File Not Found"
        End
    End If
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).