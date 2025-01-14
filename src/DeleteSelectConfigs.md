# DeleteSelectConfigs Macro for SolidWorks

## Description
This macro provides functionalities for deleting unwanted configurations from SolidWorks models. It allows users to selectively clear configurations or reset them to a default state, improving management of configuration clutter within SolidWorks documents.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
- A part or assembly must be open in SolidWorks with multiple configurations available.
- The macro should be executed within the SolidWorks environment.

## Results
- Selectively deletes configurations based on user input.
- Provides options to reset the current configuration to a default configuration.
- Updates the number of configurations dynamically and displays the deletion status.

## Steps to Setup the Macro

1. **Create the UserForm**:
   - Open the VBA editor in SolidWorks by pressing (`Alt + F11`).
   - In the Project Explorer, right-click on the project (e.g., `DeleteSelectConfigs`) and select **Insert** > **UserForm**.
     - Rename the form to `FormDeleteConfig`.
     - Design the form with the following components:
       - A `ListBox` named `ListBoxConfigurations` to display configuration names.
       - Two `CheckBoxes` for 'Select all configurations' and 'Clear all configurations'.
       - A `TextBox` for filtering configurations by name.
       - Buttons: `Reset`, `About`, `Delete`, and `Cancel`.

2. **Add VBA Code**:
   - Insert the macro code provided into the appropriate modules.
   - Implement event handlers for form controls such as buttons, checkboxes, and list boxes.

3. **Save and Run the Macro**:
   - Save the macro file (e.g., `DeleteSelectConfigs.swp`).
   - Run the macro by navigating to **Tools** > **Macro** > **Run** in SolidWorks, then select your saved macro.

4. **Using the Macro**:
   - The macro will display a form that lists all configurations of the current document.
   - Use the form controls to select configurations for deletion, set configurations to default, or filter configurations.
   - Confirm deletions or resets as needed by clicking the respective buttons.

## VBA Macro Code

```vbnet
'-------------------------------------------------------------------------------
' DeleteSelectConfigs.swp           
'-------------------------------------------------------------------------------
Global swApp As Object
Global ModelDoc2 As Object
Global Configuration As Object
Global FileTyp As String
Global numConfigs As Integer
Global ConfigNames As Variant
Global Retval As Integer
Global DelCount As Integer
' SolidWorks constants
Global Const swDocPART = 1
Global Const swDocASSEMBLY = 2
Global Const swMbWarning = 1
Global Const swMbInformation = 2
Global Const swMbQuestion = 3
Global Const swMbStop = 4
Global Const swMbAbortRetryIgnore = 1
Global Const swMbOk = 2
Global Const swMbOkCancel = 3
Global Const swMbRetryCancel = 4
Global Const swMbYesNo = 5
Global Const swMbYesNoCancel = 6
Global Const swMbHitAbort = 1
Global Const swMbHitIgnore = 2
Global Const swMbHitNo = 3
Global Const swMbHitOk = 4
Global Const swMbHitRetry = 5
Global Const swMbHitYes = 6
Global Const swMbHitCancel = 7
Global Const version = "1.50"

'------------------------------------------------------------------------------
' Delete extra configurations from model and set current config to Default
'------------------------------------------------------------------------------
Sub Main()
  Set swApp = CreateObject("SldWorks.Application")            ' Attach to SWX
  Set ModelDoc2 = swApp.ActiveDoc                             ' Grab active doc
  If ModelDoc2 Is Nothing Then                                ' Is doc loaded
    MsgBox "No active part or assembly model found in SolidWorks." & Chr(13) & Chr(13) _
          & "Please load/activate a SolidWorks part or assembly " & Chr(13) _
          & "model and try again.", vbExclamation
  Else                                                        ' Doc loaded?
    FileTyp = ModelDoc2.GetType                               ' Get doc type
    If FileTyp = swDocASSEMBLY Or FileTyp = swDocPART Then    ' Doc model?
      numConfigs = ModelDoc2.GetConfigurationCount()          ' Get # configs
      If numConfigs > 1 Then                                  ' Check # configs
        FormDeleteConfig.Show
      Else                                                    ' Else # config
        MsgBox "Only one configuration exists in this model.", vbExclamation
      End If                                                  ' End # config
    Else                                                      ' Else doc type
      MsgBox "Active document is not a SolidWorks part or assembly model." _
              & Chr(13) & Chr(13) _
              & "Please load/activate a SolidWorks model and try again.", _
              vbExclamation
    End If                                                    ' End doc type
  End If                                                      ' End doc load

End Sub
```

## VBA UserForm Code

```vbnet
'-------------------------------------------------------------------------------
' DeleteSelectConfigs.swp           
'-------------------------------------------------------------------------------
Dim DelCount As Integer

Private Sub CheckClearAll_Click()
  For x = 0 To ListBoxConfigurations.ListCount - 1             ' Each config
    ListBoxConfigurations.Selected(x) = False                  ' Clear select
  Next x
  CheckSelectAll = False                                       ' Clear check
  CheckClearAll = False                                        ' Clear check
End Sub

Private Sub CheckSelectAll_Click()
  For x = 0 To ListBoxConfigurations.ListCount - 1             ' Each config
    ListBoxConfigurations.Selected(x) = True                   ' Set select
  Next x
  CheckSelectAll = False                                       ' Clear check
  CheckClearAll = False                                        ' Clear check
End Sub

Private Sub CheckBoxSelectFilter_Click()
  If CheckBoxSelectFilter = True Then
    TextBoxFilter.Enabled = True
    TextBoxFilter.BackColor = vbWindowBackground
    CommandSelectFilter.Enabled = True
  Else
    TextBoxFilter.Enabled = False
    TextBoxFilter.BackColor = vbButtonFace
    CommandSelectFilter.Enabled = False
  End If
End Sub

Private Sub CommandAbout_Click()
  FormAbout.Show
End Sub

Private Sub CommandSelectFilter_Click()
  For x = 0 To ListBoxConfigurations.ListCount - 1             ' Each config
    For y = 0 To Len(ListBoxConfigurations.List(x, 0)) - Len(TextBoxFilter)
      If UCase(Mid$(ListBoxConfigurations.List(x, 0), y + 1, Len(TextBoxFilter))) _
             = UCase(TextBoxFilter) Then
        ListBoxConfigurations.Selected(x) = True               ' Set select
      End If
    Next y
  Next x
End Sub

Private Sub CheckDeleteConfigs_Click()
  ProcessCheck
End Sub

Private Sub CheckKeepConfig_Click()
  ProcessCheck
End Sub

Private Sub CheckBoxReset_Click()
  CommandReset.Enabled = True
End Sub

Private Sub ProcessCheck()
  If CheckDeleteConfigs = True And CheckKeepConfig = True Then
    CommandDelete.Enabled = True
  Else
    CommandDelete.Enabled = False
  End If
End Sub

Private Sub CommandCancel_Click()
  End
End Sub

Private Sub CommandDelete_Click()
  ConfigNames = ModelDoc2.GetConfigurationNames()
  For i = 0 To ListBoxConfigurations.ListCount - 1    ' For each config in list
    If ListBoxConfigurations.Selected(i) = True Then  ' If selected
      ModelDoc2.DeleteConfiguration2 ListBoxConfigurations.List(i) ' delete config
      DelCount = DelCount + 1                         ' Inc del counter
      LabelConfigsDeleted = LTrim(Str(DelCount)) + " configurations processed."
      FormDeleteConfig.Repaint
    End If                                          ' active config
  Next                                              ' Get next config
  numConfigs = numConfigs - ModelDoc2.GetConfigurationCount()
  LabelConfigsDeleted = "Done: " + LTrim(Str(numConfigs)) _
                      + " configurations deleted."
  UserForm_ReInitialize
End Sub

Private Sub CommandReset_Click()
  If CheckBoxReset = True Then
    Configuration.Name = "Default"                    ' Set config to
    Configuration.AlternateName = "Default"           ' "Default" &
    Configuration.UseAlternateNameInBOM = 0           ' AlternateName
  End If
  CheckBoxReset.Value = False
  CheckBoxReset.Enabled = False
  CommandReset.Enabled = False
  LabelConfigsDeleted = "Configuration reset to 'Default'."
End Sub

Private Sub UserForm_ReInitialize()
  ListBoxConfigurations.Clear
  Set Configuration = ModelDoc2.GetActiveConfiguration
  CurrentConfigName = Configuration.Name
  ConfigNames = ModelDoc2.GetConfigurationNames()
  numConfigs = ModelDoc2.GetConfigurationCount()    ' Get # configs
  For i = 0 To (numConfigs - 1)                     ' For each config
    If ConfigNames(i) <> CurrentConfigName Then     ' Not active conf
      ListBoxConfigurations.AddItem ConfigNames(i)
    End If                                          ' active config
  Next                                              ' Get next config
  Set Configuration = ModelDoc2.GetActiveConfiguration
  If DelCount = 0 Then
    LabelConfigsDeleted = "Current configuration: '" + Configuration.Name _
                        + "' cannot be deleted."
  End If
  DelCount = 0
  CommandReset.top = CommandDelete.top
  CheckBoxReset.top = CheckBoxSelectFilter.top
  If ModelDoc2.GetConfigurationCount() = 1 Then
    CheckBoxReset.Visible = True
    ListBoxConfigurations.Enabled = False
    ListBoxConfigurations.BackColor = vbButtonFace
    CommandReset.Visible = True
    CommandDelete.Visible = False
    CheckBoxSelectFilter.Visible = False
    TextBoxFilter.Visible = False
    CommandSelectFilter.Visible = False
    CheckSelectAll.Visible = False
    CheckClearAll.Visible = False
  Else
    CheckBoxReset.Visible = False
    CommandReset.Enabled = False
    CommandReset.Visible = False
  End If
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
End Sub

Private Sub UserForm_Initialize()
  FormDeleteConfig.Caption = FormDeleteConfig.Caption & " " & version
  FormDeleteConfig.Height = 298
  UserForm_ReInitialize
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).