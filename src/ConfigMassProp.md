# ConfigMassProp: SELECT CONFIGURATION Macro for SolidWorks

## Description
This macro provides a user interface to manage and export mass properties of various configurations within a SolidWorks document. It allows users to select configurations, retrieve their mass properties, and then export these details to a text file for further analysis or reporting.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
- A SolidWorks document (assembly or part) must be open.
- The macro should be executed within the SolidWorks environment with access rights to read and modify document properties.

## Results
- Users can view and export the mass properties of selected configurations in the assembly.
- The exported data includes both mass in pounds and kilograms, which can be useful for compliance and design documentation.

## Steps to Setup the Macro

1. **Create the UserForm**:
   - Open the VBA editor in SolidWorks by pressing (`Alt + F11`).
   - In the Project Explorer, right-click on the project (e.g., `CfgMassProp`) and select **Insert** > **UserForm**.
     - Rename the form to `FormConfigMassProp`.
     - Design the form with the following components:
       - A `ListBox` named `ListBoxConfigurations` to display the names of the configurations.
       - Buttons: `ALL`, `None`, `About`, `Retrieve`, `Export`, and `Close`.

2. **Add VBA Code**:
   - Implement the code to populate the list box with configurations from the active SolidWorks document.
   - Handle button clicks to select configurations, retrieve mass properties, and export the data.

3. **Save and Run the Macro**:
   - Save the macro file (e.g., `CfgMassProp.swp`).
   - Run the macro by navigating to **Tools** > **Macro** > **Run** in SolidWorks, then select your saved macro.

4. **Using the Macro**:
   - The macro will open the `ConfigMassProp` UserForm.
   - Use the form controls to manage configurations and perform actions like retrieving and exporting mass properties.
   - Data is saved in a text file named after the document with a suffix "_MASS.TXT".

## VBA Macro Code

```vbnet
'-------------------------------------------------------------------------------
' CfgMassProp.swp                  
'-------------------------------------------------------------------------------
Global swApp As Object
Global ModelDoc2 As Object
Global swSelMgr As Object
Global swSelObj As Object
Global SelModDoc2 As Object
Global Configuration As Object
Global FileTyp As String
Global numConfigs As Integer
Global ConfigNames As Variant
Global Retval As Integer
Global DelCount As Integer
Global Mode As Integer
' SolidWorks constants
Global Const swDocNONE = 0
Global Const swDocPART = 1
Global Const swDocASSEMBLY = 2
Global Const swDocDRAWING = 3
Global Const swDocSDM = 4
Global Const swSelDRAWINGVIEWS = 12
Global Const swSelCOMPONENTS = 20
Global Const version = "1.00"

'------------------------------------------------------------------------------
Sub Main()
  Set swApp = CreateObject("SldWorks.Application")              ' Attach to SWX
  Set ModelDoc2 = swApp.ActiveDoc                               ' Grab active doc
  If ModelDoc2 Is Nothing Then                                  ' Is doc loaded
    MsgBox "No active document found in SolidWorks." & Chr(13) & Chr(13) _
          & "Please load/activate a SolidWorks " & Chr(13) _
          & "document and try again.", vbExclamation
  Else                                                          ' Doc loaded?
    FileTyp = ModelDoc2.GetType                               ' Get doc type
    If FileTyp = swDocASSEMBLY Then                           ' Assy model?
      Set swSelMgr = ModelDoc2.SelectionManager
      Set swSelObj = swSelMgr.GetSelectedObject4(1)           ' Get selected
      If (Not swSelObj Is Nothing) And _
         (swSelMgr.GetSelectedObjectType2(1) = swSelCOMPONENTS) Then
        Mode = 1                                              ' Manage Comp
      Else                                                    ' pre-selection
        Mode = 0                                              ' Manage Assy
      End If
      FormCfgMassProp.Show
    ElseIf FileTyp = swDocPART Then                           ' Part model?
      Mode = 0                                                ' Manage Part
      FormCfgMassProp.Show
    ElseIf FileTyp = swDocDRAWING Then                        ' Drawing?
      MsgBox "Sorry, this macro does not work with drawings", vbExclamation
    Else                                                      ' Else doc type
      MsgBox "Sorry, could not determine document type.", vbExclamation
    End If                                                    ' End doc type
  End If                                                        ' End doc load
End Sub
```

## VBA UserForm Code

```vbnet
'-------------------------------------------------------------------------------
' CfgMassProp.swp                   
'-------------------------------------------------------------------------------
Dim ModDoc2         As Object
Dim Update          As Boolean
Dim LoadErr         As Long
Dim LoadWarn        As Long
Dim ExpFileName     As String
Dim ModelName       As String
Dim WorkModelDoc    As Object
Const swOpenDocOptions_AutoMissingConfig = &H20

'------------------------------------------------------------------------------
' This function will locate the occurance of "stringToFind" inside "theString".
'   Occur  = 0       Get first occurrance
'            1       Get last occurrance
' From that position, it will return text based on the 'Retrieve' variable.
'   StrRet = 0       Get text before
'            1       Get text after
'            2       If no occurrance found, return blank
'------------------------------------------------------------------------------
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

Private Sub CommandAll_Click()
  For x = 0 To ListBoxConfigurations.ListCount - 1
    ListBoxConfigurations.Selected(x) = True
  Next x
End Sub

Private Sub CommandNone_Click()
  For x = 0 To ListBoxConfigurations.ListCount - 1
    ListBoxConfigurations.Selected(x) = False
  Next x
  CommandExport.Enabled = False
End Sub

Private Sub CommandaBOUT_Click()
  FormAbout.Show
End Sub

Private Sub CommandRetrieve_Click()
  Dim ModelDocExt       As Object
  Dim MassStatus        As Long
  Dim MassValue         As Variant
  Set ModelDocExt = WorkModelDoc.Extension
  For x = 0 To ListBoxConfigurations.ListCount - 1
    If ListBoxConfigurations.Selected(x) = True Then
      WorkModelDoc.ShowConfiguration2 ListBoxConfigurations.List(x, 0)
      MassValue = ModelDocExt.GetMassProperties(1, MassStatus)
      ListBoxConfigurations.List(x, 1) = Format(MassValue(5) * 2.204622, "0.000")
      ListBoxConfigurations.List(x, 2) = Format(MassValue(5), "0.00")
      If CommandExport.Enabled = False Then CommandExport.Enabled = True
    End If
  Next x
End Sub

Private Sub CommandExport_Click()
  Dim Config        As String
  Dim MassLb
  Dim MassKG
  Open ExpFileName For Append Access Write Lock Write As #1
  Print #1, ModelName & ", Lbs, Kg"
  For x = 0 To ListBoxConfigurations.ListCount - 1
    If ListBoxConfigurations.Selected(x) = True Then
      Config = ListBoxConfigurations.List(x, 0)
      MassLb = ListBoxConfigurations.List(x, 1)
      MassKG = ListBoxConfigurations.List(x, 2)
      Print #1, Config & ", " & MassLb & ", " & MassKG
    End If
  Next x
  Print #1, vbNewLine
  Close #1
  MsgBox "File Created: " & Chr(13) & ExpFileName
End Sub

Private Sub CommandClose_Click()
  End
End Sub

Private Sub UserForm_Initialize()
  Update = False
  ' Clear form
  ListBoxConfigurations.Clear
  Select Case Mode
         Case 0         ' Top level assembly or part
                Set WorkModelDoc = ModelDoc2
                
         Case 1         ' Sub-component within assembly
                Set WorkModelDoc = swSelObj.GetModelDoc
  End Select
  ' Get configuration list and populate list in form
  ConfigNames = WorkModelDoc.GetConfigurationNames()
  numConfigs = WorkModelDoc.GetConfigurationCount()
  Set Configuration = WorkModelDoc.GetActiveConfiguration
  ExpFileName = NewParseString(UCase(WorkModelDoc.GetPathName), _
              ".SLD", 0, 0) & "_MASS.TXT"
  ModelName = NewParseString(UCase(WorkModelDoc.GetPathName), _
              "\", 1, 1)
  For x = 0 To (numConfigs - 1)                      ' For each config
    ListBoxConfigurations.AddItem ConfigNames(x)
  Next                                               ' Get next config
  ' Sort entries in ListBoxConfigurations via simple bubble sort technique.
  For x = 0 To ListBoxConfigurations.ListCount - 2   ' Number of passes
    For y = 0 To ListBoxConfigurations.ListCount - 2 ' Pass thru each row
      If ListBoxConfigurations.List(y, 0) _
         > ListBoxConfigurations.List(y + 1, 0) Then ' If greater, re-order
        Temp = ListBoxConfigurations.List(y, 0)      ' Re-order entries
        ListBoxConfigurations.List(y, 0) = ListBoxConfigurations.List(y + 1, 0)
        ListBoxConfigurations.List(y + 1, 0) = Temp
      End If                                         ' Re-order check
    Next y                                           ' Pass thru each row
  Next x                                             ' Number of passes
  ' Identify current configuration
  For x = 0 To (ListBoxConfigurations.ListCount - 1) ' For each config
    If ListBoxConfigurations.List(x, 0) = ConfigName Then
      ListBoxConfigurations.Selected(x) = True       ' Found active conf
    End If                                           ' active config
  Next                                               ' Get next config
  CommandExport.Enabled = False
  LabelDocument = "Document: " & ModelName
  Update = True
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).