# DataCharter: Chart Creator Macro for SolidWorks

## Description
This macro, known as DataCharter, allows users to create charts based on dimensional data from SolidWorks models. It facilitates the entry of parameters, chart creation, and data exportation to both CSV and Excel formats.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
- A SolidWorks part or assembly with defined dimensions must be open.
- The macro should be executed within the SolidWorks environment.

## Results
- Users can generate customizable charts based on specific dimension parameters.
- Provides functionality to export the charted data to CSV or Excel for further analysis or reporting.

## Steps to Setup the Macro

1. **Create the UserForm**:
   - Open the VBA editor in SolidWorks by pressing (`Alt + F11`).
   - In the Project Explorer, locate the `DataCharter` project.
   - Right-click on the `Forms` folder and select **Insert** > **UserForm**.
     - Rename the newly created form to `FormEnterParameters`.
     - Design the form to include:
       - Text boxes for parameter input.
       - Buttons for operations such as `Create Chart`, `Export to CSV`, and `Export to Excel`.
       - A ListBox for displaying charted data.

2. **Implement the Module**:
   - Right-click on the `Modules` folder within the `DataCharter` project.
   - Select **Insert** > **Module**.
   - Add VBA code to this module (`DataCharter1`) to handle the logic for chart creation, data retrieval, and exporting functionalities.

3. **Add Class Modules**:
   - If there are specific object-oriented functionalities or event handling required:
     - Right-click on `Class Modules` in the `DataCharter` project.
     - Select **Insert** > **Class Module**.
     - Implement class-based logic in `ClassEvents`.

4. **Save and Run the Macro**:
   - Save the macro file (e.g., `DataCharter.swp`).
   - Run the macro by navigating to **Tools** > **Macro** > **Run** in SolidWorks, then select your saved macro.

5. **Using the Macro**:
   - The macro will display the `FormEnterParameters`.
   - Input horizontal and vertical scale parameters along with desired precision.
   - Use the `Create Chart` button to generate the chart within the ListBox.
   - Export the chart data using the `Export to CSV` or `Export to Excel` buttons.

## VBA Macro Code

```vbnet
' ------------------------------------------------------------------------------
' DataCharter.swp                   
' ------------------------------------------------------------------------------
' DataCharter1:         Macro is launched from "main" subroutine in this module.
' ------------------------------------------------------------------------------

Global swApp                As Object
Global Part                 As Object
Global swSelMgr             As Object
Global swMouse              As SldWorks.mouse
Global TheMouse             As SldWorks.mouse
Global MouseObj             As New ClassEvents
Global swModelDocExt        As SldWorks.ModelDocExtension
Global swModelView          As SldWorks.ModelView
Global SelectDestination    As Object
Global MyDimName(3)         As String
Global MyDimType(3)         As Integer
Global MyWorkDim            As Integer
Global mySalmon
Global Pi

Sub main()
  ' Attach to SolidWorks and get active document
  mySalmon = RGB(255, 160, 160) ' Salmon background
  Pi = 4 * Atn(1)                                             ' The infamous pi value
  Set swApp = Application.SldWorks
  Set Part = swApp.ActiveDoc       ' Grab currently active document
  Set swModelDocExt = Part.Extension
  Set swModelView = Part.GetFirstModelView
  Set TheMouse = swModelView.GetMouse
  MouseObj.init TheMouse
' Set eventListener = New ClassEvents
  FormEnterParameters.Show
End Sub
```

## VBA UserForm Code

```vbnet
' ------------------------------------------------------------------------------
' DataCharter.swp                   
' ------------------------------------------------------------------------------
' DataCharter1:         Macro is launched from "main" subroutine in this module.
' ------------------------------------------------------------------------------

Dim BaseFileName    As String

' -----------------------------------------------------------------------------
' This function will return text based on occurance of a string "stringToFind"
' within another string "theString", and text before/after defined below.
'
' Locate the occurance of "stringToFind", based on the occur variable.
'   Occur  = 0       Get first occurrance
'            1       Get last occurrance
'
' Return the text (before/after) based on the 'Retrieve' variable.
'   StrRet = 0       Get text before
'            1       Get text after
'            2       If no occurrance found, return blank
' -----------------------------------------------------------------------------
Function NewParseString(theString As String, stringToFind As String, _
                            Occur As Boolean, StrRet As Integer)
  Location = 0
  TempLoc = 0
  PosCount = 1
  While PosCount < Len(theString)
    TempLoc = InStr(PosCount, theString, stringToFind)
    If TempLoc <> 0 Then
      If Occur = 0 Then
        If Location = 0 Then
          Location = TempLoc
        End If
      Else
        Location = TempLoc
      End If
      PosCount = TempLoc + 1
    Else
      PosCount = Len(theString)
    End If
  Wend
  If Location <> 0 Then
    If StrRet = 0 Then
      theString = Mid(theString, 1, Location - 1)
    Else
      theString = Mid(theString, Location + Len(stringToFind))
    End If
  ElseIf StrRet = 2 Then
    theString = ""
  End If
  NewParseString = theString
End Function

Function CUtoMTR(Value1 As Double, DimObj As Integer)
  If MyDimType(DimObj) = swDimensionParamTypeDoubleLinear Then
    Select Case Part.GetUserPreferenceIntegerValue(swUnitsLinear)
         Case swMM
                CUtoMTR = Value1 * 0.001
         Case swCM
                CUtoMTR = Value1 * 0.01
         Case swMETER
                CUtoMTR = Value1 * 1
         Case swINCHES
                CUtoMTR = Value1 * 0.0254
         Case swFEET
                CUtoMTR = Value1 * 0.0254 * 12
         Case Else
                MsgBox "Document units are not recognized"
    End Select
  ElseIf MyDimType(DimObj) = swDimensionParamTypeDoubleAngular Then
    Select Case Part.GetUserPreferenceIntegerValue(swUnitsAngular)
         Case swDEGREES
                CUtoMTR = Value1 * (Pi / 180)
         Case swRADIANS
                CUtoMTR = Value1 * 1
         Case Else
                MsgBox "Document units are not recognized"
    End Select
  End If
End Function

Function MTRtoCU(Value1 As Double, DimObj As Integer)
  If MyDimType(DimObj) = swDimensionParamTypeDoubleLinear Then
    Select Case Part.GetUserPreferenceIntegerValue(swUnitsLinear)
         Case swMM
                MTRtoCU = Value1 / 0.001
         Case swCM
                MTRtoCU = Value1 / 0.01
         Case swMETER
                MTRtoCU = Value1 / 1
         Case swINCHES
                MTRtoCU = Value1 / 0.0254
         Case swFEET
                MTRtoCU = Value1 / (0.0254 * 12)
         Case Else
                MsgBox "Document units are not recognized"
    End Select
  ElseIf MyDimType(DimObj) = swDimensionParamTypeDoubleAngular Then
    Select Case Part.GetUserPreferenceIntegerValue(swUnitsAngular)
         Case swDEGREES
                MTRtoCU = Value1 * (180 / Pi)
         Case swRADIANS
                MTRtoCU = Value1 * 1
         Case Else
                MsgBox "Document units are not recognized"
    End Select
  End If
End Function

Private Sub CommandChart_Click()
  Dim Values        As String
  Dim Separator     As Integer
  Dim NextCol       As Integer
  Dim NextRow       As Integer
  Dim X             As Integer
  Dim ColSpace      As String
  ' - - - - - - - - - - - - - - - - - - - -
  ' Clear Chart
  ListBoxChart.Clear
  ' - - - - - - - - - - - - - - - - - - - -
  ' Setup Top Header Row
  NextRow = 0
  NextCol = 1
On Error GoTo CheckValues
  Values = TextValHor.Value
  ListBoxChart.AddItem 0
  Separator = InStr(1, Values, " ")
  While Separator > 0
    ListBoxChart.List(NextRow, NextCol) = CDbl(Left$(Values, Separator - 1))
    Values = Mid$(Values, Separator + 1)
    Separator = InStr(1, Values, " ")
    If NextCol < 9 Then
      NextCol = NextCol + 1
    Else
      GoTo EndCols
    End If
  Wend
    If NextCol < 9 Then
      If Values <> "" Then ListBoxChart.List(NextRow, NextCol) = CDbl(Values)
    End If
  GoTo ResumeMacro
EndCols:
  MsgBox "This macro only supports 9 possible values for horizontal scale"
ResumeMacro:
  ListBoxChart.ColumnCount = NextCol + 1
  ' - - - - - - - - - - - - - - - - - - - -
  ' Setup Column Count And Column Widths
  ColSpace = ""
  For X = 0 To NextCol
    ColSpace = ColSpace + "50 pt;"
  Next X
  ColSpace = ColSpace + "50 pt"
  ListBoxChart.ColumnWidths = ColSpace
  ' - - - - - - - - - - - - - - - - - - - -
  ' Setup Left Header Column
  Values = TextValVer.Value
  Separator = InStr(1, Values, " ")
  While Separator > 0
    ListBoxChart.AddItem CDbl(Left$(Values, Separator - 1))
    Values = Mid$(Values, Separator + 1)
    Separator = InStr(1, Values, " ")
  Wend
  If Values <> "" Then ListBoxChart.AddItem CDbl(Values)
  ' - - - - - - - - - - - - - - - - - - - -
  ' At this point:
  '     All values have been read in and chart "headers" have been created.
  '     Dimensions still need to be updated and proper data can be retrieved.
  ' - - - - - - - - - - - - - - - - - - - -
  ' For each row in chart
  For Y = 1 To ListBoxChart.ListCount - 1
    ' For each column in chart
    For X = 1 To ListBoxChart.ColumnCount - 1
      ' Update dimensions for Horizontal and Vertical Scales
      Part.Parameter(TextDimHor).SystemValue _
            = CUtoMTR(CDbl(ListBoxChart.List(0, X)), 1)
      Part.Parameter(TextDimVer).SystemValue _
            = CUtoMTR(CDbl(ListBoxChart.List(Y, 0)), 2)
      ' Read Charted Dimension and apply to Dimension to Chart
      ListBoxChart.List(Y, X) _
            = Round(MTRtoCU(Part.Parameter(TextDimCht).SystemValue, 3), _
                    Len(ComboPrec) - 1)
    Next X
  Next Y
  ' - - - - - - - - - - - - - - - - - - - -
  ' Chart has been updated.
  ' Enable export buttons.
  CommandExportCSV.Enabled = True
  CommandExportExcel.Enabled = True
  ' - - - - - - - - - - - - - - - - - - - -
  ' Display charted data
  ListBoxDisplay.ListIndex = 1
  GoTo NoError
CheckValues:
  MsgBox "Invalid inputs.  Please verify inputs", vbInformation, "DataCharter"

NoError:
End Sub

Private Sub CommandCancel_Click()
  End
End Sub

Private Sub CommandExportCSV_Click()
  ' - - - - - - - - - - - - - - - - - - - -
  ' Setup Export file name
  Dim ExportName        As String
  ExportName = NewParseString(UCase(Part.GetPathName), _
                  ".SLD", 0, 0) + "-DataCht.csv"
  ' - - - - - - - - - - - - - - - - - - - -
  ' Export data
  Open ExportName For Append Access Write Lock Write As #1
  For Row = 0 To ListBoxChart.ListCount - 1             ' For each row
    PrintText = ""                                      ' Clear PrintTest
    For Col = 0 To ListBoxChart.ColumnCount - 1         ' for each column
      PrintText = PrintText + ListBoxChart.List(Row, Col)
      If Col <> ListBoxChart.ColumnCount - 1 Then
        PrintText = PrintText + ","                     ' Add delimit char
      End If
    Next Col                                            ' Next column
    Print #1, LTrim(RTrim(PrintText))                   ' Print to file
  Next Row                                              ' Next row
  Print #1, vbNewLine                                   ' Print blank line
  Close #1
  MsgBox "Charted data has been exported to the file:" & Chr(13) _
          & ExportName, vbInformation, "DataCharter"
  Me.Repaint
End Sub

' -----------------------------------------------------------------------------
' Export to Excel Spreadsheet
' -----------------------------------------------------------------------------
Private Sub CommandExportExcel_Click()
  ' - - - - - - - - - - - - - - - - - - - -
  ' Setup Export file name
  Dim ExportName        As String
  ExportName = NewParseString(UCase(Part.GetPathName), _
                  ".SLD", 0, 0) + "-DataCht.xls"
  ' - - - - - - - - - - - - - - - - - - - -
  ' Launch and attach to the Excel API
  Set XLApp = CreateObject("Excel.Application")
  XLApp.Visible = True
  ' - - - - - - - - - - - - - - - - - - - -
  ' Start new document
  Set wb = XLApp.Workbooks.Add
  ' - - - - - - - - - - - - - - - - - - - -
  ' Set column formats to "Text"
  Prec = "."
  For Col = 1 To Len(ComboPrec) - 1
    Prec = Prec + "0"
  Next Col
  ' - - - - - - - - - - - - - - - - - - - -
  ' Format precision in excel columns
  For Col = 0 To ListBoxChart.ColumnCount - 1       ' for each column
    XLApp.Columns(Chr(Col + 65) & ":" & Chr(Col + 65)).Select
    XLApp.Selection.NumberFormat = Prec             ' Set precision
  Next Col
  ' - - - - - - - - - - - - - - - - - - - -
  ' Prepare to generate BOM in Excel
  Row = 0
  ' - - - - - - - - - - - - - - - - - - - -
  ' For each row
  While (Row < ListBoxChart.ListCount)              ' For each row
    XLCol = 0
    ' - - - - - - - - - - - - - - - - - - - -
    ' Write title row to top of current page of spreadsheet
    For Col = 0 To ListBoxChart.ColumnCount - 1           ' for each column
        ' Write data to specific cell on current page of spreadsheet
        ' Convert listbox coordinates to excel coordinates
        ' NOTE: Listbox cells starts at 0,0 and Excel cells starts at A1
        XLApp.ActiveSheet.Range(Chr(XLCol + 65) & Row + 1).Value _
             = ListBoxChart.List(Row, Col)
        XLCol = XLCol + 1
    Next Col
    Row = Row + 1               ' Next row to read
    XLRow = XLRow + 1           ' Next row to write to Excel
  Wend                          ' Continue loop until all is read
  ' - - - - - - - - - - - - - - - - - - - -
  ' Set each column to AutoFit
  ' to force Autofit of Excel columns, remove apostrophes from next 3 lines
  ' For Col = 0 To ListBoxChart.ListCount - 1
  '   XLApp.Columns(Chr(Col + 65) & ":" & Chr(Col + 65)).EntireColumn.AutoFit
  ' Next Col
  ' - - - - - - - - - - - - - - - - - - - -
  ' Go to cell A1
  XLApp.Range("A1:A1").Select
  ' - - - - - - - - - - - - - - - - - - - -
  ' Save spreadsheet
  XLApp.ActiveWorkbook.SaveAs FileName:=ExportName
  ' - - - - - - - - - - - - - - - - - - - -
  ' Close Excel
  XLApp.Quit
  Set XLApp = Nothing
  MsgBox "Charted data has been exported to the file:" & Chr(13) _
          & ExportName, vbInformation, "DataCharter"
End Sub

Private Sub ListBoxDisplay_Click()
  If ListBoxDisplay.ListIndex = 0 Then
    FrameEnter.Visible = True
    FrameDataChart.Visible = False
  Else
    FrameEnter.Visible = False
    FrameDataChart.Visible = True
  End If
End Sub

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' If a field is changed, do the following:
'  * Check inputs to see if all fields contain data
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Private Sub TextDimHor_Change()
  CheckInputs
End Sub
Private Sub TextDimCht_Change()
  CheckInputs
End Sub
Private Sub TextDimVer_Change()
  CheckInputs
End Sub
Private Sub TextValHor_Change()
  CheckInputs
End Sub
Private Sub TextValVer_Change()
  CheckInputs
End Sub
Private Sub ComboPrec_Change()
  CheckInputs
End Sub
Private Sub CheckInputs()
  InputsOK = True
  ' - - - - - - - - - - - - - - - - - - - -
  ' Check for entries in input fields
  If TextDimHor.Value = "" Then InputsOK = False
  If TextValHor.Value = "" Then InputsOK = False
  If TextDimVer.Value = "" Then InputsOK = False
  If TextValVer.Value = "" Then InputsOK = False
  If TextDimCht.Value = "" Then InputsOK = False
  If ComboPrec.ListIndex < 0 Then InputsOK = False
  ' - - - - - - - - - - - - - - - - - - - -
  ' Ensure all input fields contain data.
  If InputsOK = True Then
    ' - - - - - - - - - - - - - - - - - - - -
    ' If all input fields contain data, enable "Create Chart" button
    CommandChart.Enabled = True
  Else
    ' - - - - - - - - - - - - - - - - - - - -
    ' If any input fields are blank, disable "Create Chart" button
    CommandChart.Enabled = False
  End If
  ' - - - - - - - - - - - - - - - - - - - -
  ' Input fields may have changed, disable Export buttons for now
  CommandExportCSV.Enabled = False
  CommandExportExcel.Enabled = False
End Sub

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' When these fields are entered (selected), do the following:
'   * Highlight field with salmon color
'   * Set destination of name of selected dimension
'   * Repaint form
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Private Sub TextDimHor_Enter()
  TextDimHor.BackColor = mySalmon
  TextDimHor.ForeColor = vbBlack
  MyWorkDim = 1
  Me.Repaint
End Sub
Private Sub TextDimVer_Enter()
  TextDimVer.BackColor = mySalmon
  TextDimVer.ForeColor = vbBlack
  MyWorkDim = 2
  Me.Repaint
End Sub
Private Sub TextDimCht_Enter()
  TextDimCht.BackColor = mySalmon
  TextDimCht.ForeColor = vbBlack
  MyWorkDim = 3
  Me.Repaint
End Sub

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' When these fields are exited (after update), do the following:
'   * Reset field to standard colors
'   * Clear destination of name of selected dimension
'   * Repaint form
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Private Sub TextDimHor_AfterUpdate()
  TextDimHor.BackColor = vbWindowBackground
  TextDimHor.ForeColor = vbWindowText
  Me.Repaint
End Sub

Private Sub TextDimVer_AfterUpdate()
  TextDimVer.BackColor = vbWindowBackground
  TextDimVer.ForeColor = vbWindowText
  Me.Repaint
End Sub

Private Sub TextDimCht_AfterUpdate()
  TextDimCht.BackColor = vbWindowBackground
  TextDimCht.ForeColor = vbWindowText
  Me.Repaint
End Sub

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Initialize user form
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Private Sub UserForm_Initialize()
  ' - - - - - - - - - - - - - - - - - - - -
  ' Populate precision combobox
  ComboPrec.AddItem "1"
  ComboPrec.AddItem ".1"
  ComboPrec.AddItem ".01"
  ComboPrec.AddItem ".001"
  ComboPrec.AddItem ".0001"
  ComboPrec.AddItem ".00001"
  ComboPrec.AddItem ".000001"
  ComboPrec.AddItem ".0000001"
  ComboPrec.AddItem ".00000001"
  ComboPrec.ListIndex = 3
  ' - - - - - - - - - - - - - - - - - - - -
  ' Setup frame view selection
  ListBoxDisplay.AddItem "Enter Parameters"
  ListBoxDisplay.AddItem "Charted Data"
  ListBoxDisplay.ListIndex = 0
  ' - - - - - - - - - - - - - - - - - - - -
  ' Reposition form elements and resize form
  FrameEnter.Visible = True
  FrameDataChart.Visible = False
  FrameDataChart.Top = 6
  FormEnterParameters.Height = 168
  ' - - - - - - - - - - - - - - - - - - - -
  ' Handle Pre-Selections.  Selections made before macro was launched
  ' - - - - - - - - - - - - - - - - - - - -
  ' Get selection manager
  Set swSelMgr = Part.SelectionManager
  ' - - - - - - - - - - - - - - - - - - - -
  ' Get number of selections
  Count = swSelMgr.GetSelectedObjectCount2(-1)
  If Count > 0 Then                             ' Is anything selected
    For MyCount = 1 To 3
      ' - - - - - - - - - - - - - - - - - - - -
      ' What type of object was selected
      SelType = swSelMgr.GetSelectedObjectType3(MyCount, -1)
      ' - - - - - - - - - - - - - - - - - - - -
      ' Is it a dimension?
      If swSelMgr.GetSelectedObjectType3(MyCount, -1) = swSelDIMENSIONS Then
        ' - - - - - - - - - - - - - - - - - - - -
        ' For dimensions only
        Set DimSelect = swSelMgr.GetSelectedObject6(MyCount, -1)
        Set swDimen = DimSelect.GetDimension
        If Not DimSelect Is Nothing Then
          Select Case MyCount
                 Case 1
                      TextDimHor.Value = swDimen.FullName
                      MyDimType(1) = swDimen.GetType
                 Case 2
                      TextDimVer.Value = swDimen.FullName
                      MyDimType(2) = swDimen.GetType
                 Case 3
                      TextDimCht.Value = swDimen.FullName
                      MyDimType(3) = swDimen.GetType
          End Select
        End If
      End If
    Next MyCount
  End If
  CheckInputs
End Sub
```

### ClassEvents Class Code

```vbnet
' ------------------------------------------------------------------------------
' DataCharter.swp                 
' ------------------------------------------------------------------------------
' Mouse handler Class for dimension selection in FormEnterParameters
' ------------------------------------------------------------------------------
Option Explicit
Dim WithEvents ms As SldWorks.mouse

Private Sub Class_Initialize()
End Sub

Public Sub init(mouse As Object)
Set ms = mouse
End Sub

' Private Function ms_MouseNotify(ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' Debug.Print "Event Message= " & Message & " wParam= " & wParam & " lParam= "; lParam
' End Function

Private Function ms_MouseSelectNotify(ByVal ix As Long, ByVal iy As Long, ByVal X As Double, ByVal Y As Double, ByVal Z As Double) As Long
  Dim UserSelect        As Boolean
  Dim Count             As Integer
  Dim SelType           As String
  Dim DimSelect         As Object
  Dim swDimen           As Object
  If Not MyWorkDim = 0 Then
    UserSelect = swSelMgr.EnableSelection
    If UserSelect = True Then
      Count = swSelMgr.GetSelectedObjectCount2(-1)
      If Count > 0 Then                               ' Is anything selected
        ' What type of object was selected
        SelType = swSelMgr.GetSelectedObjectType3(1, -1)
        ' Is it a dimension?
        If swSelMgr.GetSelectedObjectType3(1, -1) = swSelDIMENSIONS Then
          ' For dimensions only
          Set DimSelect = swSelMgr.GetSelectedObject6(1, -1)
          Set swDimen = DimSelect.GetDimension
          If Not DimSelect Is Nothing Then
            MyDimName(MyWorkDim) = swDimen.FullName
            Select Case MyWorkDim
                   Case 1
                        FormEnterParameters.TextDimHor.Value = MyDimName(MyWorkDim)
                   Case 2
                        FormEnterParameters.TextDimVer.Value = MyDimName(MyWorkDim)
                   Case 3
                        FormEnterParameters.TextDimCht.Value = MyDimName(MyWorkDim)
            End Select
            MyDimType(MyWorkDim) = swDimen.GetType
          End If
        End If
      End If
    End If
  End If
End Function
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).