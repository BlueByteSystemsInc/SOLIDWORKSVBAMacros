# ConfigManager Macro for SolidWorks

## Description
The ConfigManager macro allows users to efficiently manage configurations in SolidWorks models. It enables selecting and switching between configurations in parts, assemblies, and drawings, with additional options to display descriptions and zoom to fit the selected configuration.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
- A SolidWorks model (part, assembly, or drawing) must be open.
- The macro should be executed within the SolidWorks environment.

## Results
- Displays a list of configurations available in the active document or referenced components.
- Allows users to switch between configurations and updates properties such as descriptions.
- Offers a zoom-to-fit functionality for improved visualization.

## Steps to Setup the Macro

1. **Create the UserForm**:
   - Open the VBA editor in SolidWorks by pressing (`Alt + F11`).
   - In the Project Explorer, right-click on the project (e.g., `ConfigManager`) and select **Insert** > **UserForm**.
     - Rename the form to `FormConfigMgr`.
     - Design the form with the following components:
       - A `ListBox` named `ListBoxConfigurations` to display the configuration names.
       - A `Label` named `LabelDescription` to display the description of the selected configuration.
       - A `CheckBox` named `CheckFitScreen` for enabling or disabling the zoom-to-fit functionality.
       - Buttons: `About` (to display macro information) and `Close`.

2. **Add VBA Code**:
   - Copy the **Macro Code** provided below into the main module.
   - Copy the **UserForm Code** into the `FormConfigMgr` code-behind.

3. **Save and Run the Macro**:
   - Save the macro file (e.g., `ConfigManager.swp`).
   - Run the macro by navigating to **Tools** > **Macro** > **Run** in SolidWorks, then select your saved macro.

4. **Using the Macro**:
   - The macro displays a list of configurations for the active document or its referenced components.
   - Select a configuration to switch to it and view its description.
   - Use the `CheckFitScreen` option to zoom to the selected configuration.

## VBA Macro Code

```vbnet
'-------------------------------------------------------------------------------
' ConfigManager.swp                 
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
Global Const version = "1.10"

'------------------------------------------------------------------------------
Sub Main()
  Set swApp = CreateObject("SldWorks.Application")            ' Attach to SWX
  Set ModelDoc2 = swApp.ActiveDoc                             ' Grab active doc
  If ModelDoc2 Is Nothing Then                                ' Is doc loaded
    MsgBox "No active document found in SolidWorks." & Chr(13) & Chr(13) _
          & "Please load/activate a SolidWorks " & Chr(13) _
          & "document and try again.", vbExclamation
  Else                                                        ' Doc loaded?
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
      FormConfigMgr.Show
    ElseIf FileTyp = swDocPART Then                           ' Part model?
      Mode = 0                                                ' Manage Part
      FormConfigMgr.Show
    ElseIf FileTyp = swDocDRAWING Then                        ' Drawing?
      Set swSelMgr = ModelDoc2.SelectionManager
      Set swSelObj = swSelMgr.GetSelectedObject4(1)           ' Get selected
      If (Not swSelObj Is Nothing) And _
         (swSelMgr.GetSelectedObjectType2(1) = swSelDRAWINGVIEWS) Then
        Mode = 2                                              ' Preselect views
      Else                                                    ' Valid selection
        Mode = 3                                              ' All views
      End If
      FormConfigMgr.Show
    Else                                                      ' Else doc type
      MsgBox "Sorry, could not determine document type.", vbExclamation
    End If                                                    ' End doc type
  End If                                                      ' End doc load
End Sub
```

## VBA UserForm Code

```vbnet
'-------------------------------------------------------------------------------
' ConfigManager.swp                 
'-------------------------------------------------------------------------------
Dim ModDoc2 As Object
Dim Update As Boolean
Dim LoadErr As Long
Dim LoadWarn As Long
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

Private Sub CommandaBOUT_Click()
  FormAbout.Show
End Sub

Private Sub ListBoxConfigurations_Click()
  If Update = True Then
    Select Case Mode
         Case 0         ' Top level assembly or part
                ' Show configuration
                ModelDoc2.ShowConfiguration2 ListBoxConfigurations.Value
                CheckFitScreen_Click
                ' Retrieve DESCRIPTION property
                LabelDescription = _
                  ModelDoc2.GetCustomInfoValue(ListBoxConfigurations.Value, _
                  "DESCRIPTION")
         Case 1         ' Sub-component within assembly
                ' Show configuration
                swSelObj.ReferencedConfiguration = ListBoxConfigurations.Value
                ' Retrieve DESCRIPTION property
                LabelDescription = _
                  SelModDoc2.GetCustomInfoValue(ListBoxConfigurations.Value, _
                  "DESCRIPTION")
                ModelDoc2.EditRebuild
         Case 2         ' Drawing view
                Count = swSelMgr.GetSelectedObjectCount
                For x = 1 To Count
                  LabelDescription = "View " & x & "/" & Count & " - " _
                      & SelModDoc2.GetCustomInfoValue(ListBoxConfigurations.Value, _
                      "DESCRIPTION")
                  FormConfigMgr.Repaint
                  ' Show configuration
                  Set swSelObj = swSelMgr.GetSelectedObject4(x)           ' Get selected
                  swSelObj.ReferencedConfiguration = ListBoxConfigurations.Value
                Next x
                ' Retrieve DESCRIPTION property
                ModelDoc2.EditRebuild
         Case 3         ' All drawing views
                SheetNames = ModelDoc2.GetSheetNames            ' Get names of sheets
                Count = ModelDoc2.GetSheetCount - 1
                For x = 0 To Count                              ' For each sheet
                  LabelDescription = "Sheet " & x + 1 & "/" & Count + 1 & " - " _
                      & ModDoc2.GetCustomInfoValue(ListBoxConfigurations.Value, _
                      "DESCRIPTION")
                  FormConfigMgr.Repaint
                  ModelDoc2.ActivateSheet (SheetNames(x))       ' Activate sheet
                  Set view = ModelDoc2.GetFirstView             ' Get first view
                  While Not view Is Nothing                     ' View is valid
                    view.ReferencedConfiguration = ListBoxConfigurations.Value
                    Set view = view.GetNextView                 ' Get next view
                  Wend                                          ' Repeat for all views
                  ModelDoc2.EditRebuild
                Next x                                          ' Next sheet
                If x > 0 Then                                   ' More than 1 sheet
                  ModelDoc2.ActivateSheet (SheetNames(0))      ' Back to 1st sheet
                  ModelDoc2.EditRebuild
                End If
    End Select
  End If
End Sub

Private Sub CommandClose_Click()
  End
End Sub

Private Sub CheckFitScreen_Click()
  If (CheckFitScreen.Value = True) And (CheckFitScreen.enabled = True) _
        Then ModelDoc2.ViewZoomtofit
End Sub

Private Sub UserForm_Initialize()
  Update = False
  ' Clear form
  LabelDescription.Caption = ""
  ListBoxConfigurations.Clear
  Select Case Mode
         Case 0         ' Top level assembly or part
                ' Get configuration list and populate list in form
                ConfigNames = ModelDoc2.GetConfigurationNames()
                numConfigs = ModelDoc2.GetConfigurationCount()
                Set Configuration = ModelDoc2.GetActiveConfiguration
         Case 1         ' Sub-component within assembly
                Set SelModDoc2 = swSelObj.GetModelDoc
                ConfigName = swSelObj.ReferencedConfiguration ' Get configuration
                ConfigNames = SelModDoc2.GetConfigurationNames()
                numConfigs = SelModDoc2.GetConfigurationCount()
                Set Configuration = SelModDoc2.GetConfigurationByName(Name)
         Case 2         ' Preselected views
                ModelName = swSelObj.GetReferencedModelName
                NameCheck = NewParseString(UCase(ModelName), ".SLD", 1, 1)
                If UCase(Left(NameCheck, 3)) = "PRT" Then
                  DocType = swDocPART
                Else
                  DocType = swDocASSEMBLY
                End If
                ShortCheck = NewParseString(NewParseString(UCase(ModelName), _
                             "\", 1, 1), ".SLD", 0, 0)
                ConfigName = swSelObj.ReferencedConfiguration ' Get configuration
                Set SelModDoc2 = swApp.OpenDoc6(ShortCheck, DocType, _
                             swOpenDocOptions_AutoMissingConfig, _
                             ConfigName, LoadErr, LoadWarn)
                ConfigNames = SelModDoc2.GetConfigurationNames()
                numConfigs = SelModDoc2.GetConfigurationCount()
                Set Configuration = SelModDoc2.GetConfigurationByName(ConfigName)
                CheckFitScreen.enabled = False
         Case 3         ' All drawing views
                ModelName = ModelDoc2.GetDependencies(False, False)
                NumDepend = ModelDoc2.GetNumDependencies(False, False)
                If NumDepend > 1 Then
                  NameCheck = NewParseString(UCase(ModelName(1)), _
                                ".SLD", 1, 1)
                  If UCase(Left(NameCheck, 3)) = "PRT" Then
                    DocType = swDocPART
                  Else
                    DocType = swDocASSEMBLY
                  End If
                  ShortCheck = NewParseString(NewParseString( _
                               UCase(ModelName(1)), "\", 1, 1), ".SLD", 0, 0)
                  SheetNames = ModelDoc2.GetSheetNames       ' List sheet names
                  ModelDoc2.ActivateSheet (SheetNames(0))    ' Activate sheet
                  Set view = ModelDoc2.GetFirstView          ' Get first view
                  Set view = view.GetNextView
                  If view.IsModelLoaded = True Then
                    ModelName = view.GetReferencedModelName
                    ShortCheck = NewParseString(NewParseString( _
                                 UCase(ModelName), "\", 1, 1), ".SLD", 0, 0)
                    ConfigName = view.ReferencedConfiguration ' Get configuration
                    Set ModDoc2 = swApp.OpenDoc6(ShortCheck, DocType, _
                               swOpenDocOptions_AutoMissingConfig, ConfigName, _
                               LoadErr, LoadWarn)
                    ConfigNames = ModDoc2.GetConfigurationNames()
                    numConfigs = ModDoc2.GetConfigurationCount()
                    Set ModConfig = ModDoc2.GetConfigurationByName(ConfigName)
                    ModelData = True
                  Else
                    ModelData = False
                  End If
                End If
                CheckFitScreen.enabled = False
  End Select
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
  Update = True
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).