# Component State Switcher Macro for SolidWorks

## Description
This macro provides an interface for managing component states in a SolidWorks assembly. It allows users to switch between suppressed, resolved, hidden, and shown states of components. The interface displays components' current states and facilitates batch operations on these states.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
- An assembly must be open in SolidWorks.
- The macro should be executed within the SolidWorks environment with access rights to modify component properties.

## Results
- Users can change the suppression or visibility state of selected components from the assembly.
- Provides feedback on the number of components modified and allows undoing changes.
- Enhances efficiency in managing large assemblies with complex component hierarchies.

## Steps to Setup the Macro

1. **Create the UserForm**:
   - Open the VBA editor in SolidWorks by pressing (`Alt + F11`).
   - In the Project Explorer, right-click on the project (e.g., `ComponentStateSwitcher`) and select **Insert** > **UserForm**.
     - Rename the form to `FormComponentState`.
     - Design the form with the following components:
       - A `ListBox` named `lstPartName` to display the component names and their states.
       - Radio buttons for filtering: `All Components`, `Suppressed Only`, `Hidden Only`.
       - Buttons: `Suppress`, `Visibility`, `Delete`, `Undo Delete`, and `Close`.

2. **Add VBA Code**:
   - Implement the code to populate the list box with components and their states.
   - Handle button clicks to update component states based on the user's selections and inputs.

3. **Save and Run the Macro**:
   - Save the macro file (e.g., `ComponentStateSwitcher.swp`).
   - Run the macro by navigating to **Tools** > **Macro** > **Run** in SolidWorks, then select your saved macro.

4. **Using the Macro**:
   - The macro will display a form that lists all components of the current document along with their suppression and visibility states.
   - Select components and use the form controls to change states as needed.
   - Confirm changes or undo previous actions using the respective buttons.

## VBA Macro Code

```vbnet
Option Explicit

Sub Main()
    frmPreview.Show
End Sub
```

## VBA UserForm Code

```vbnet
' -----------------------------------------------------------------------------03/09/2005
' ComponentStateSwitcher by Leonard Kikstra
' ---------------------------------------------------------------------------------------
' This macro will display the state (suppress/resolve/hidden/shown) of each component in
' the top level of the current assembly.  When user selects a component, the user will
' see a preview of the component.  The user can then change the state of the selected
' component.  See top right to see the date of the last update of this macro.
'
' Format of columns in lstPartName Listbox
'       1) Component number.
'       2) Full path to component. (hidden from user)
'       3) Suppression/Resolved state.
'       4) Shown/Hidden visibility state.
'
' ---------------------------------------------------------------------------------------
' This macro is highly based on Suppressed Parts Helper by Alastair Cardwell (See below)
'****************************************************************************************
'Suppressed Parts Helper by Alastair Cardwell
'Macro to preview suppressed parts in an assembly.
'This macro will list all suppressed parts in the assembly. Clicking on
'a part name in the list will display a preview of the selected part. When a
'part is selected it can be unsuppressed or deleted by clicking the appropriate button.
'
'Last Updated 22 Feb 2005
'*****************************************************************************************

Option Explicit

Dim swApp       As SldWorks.SldWorks
Dim swAssem     As SldWorks.ModelDoc2
Dim swRootComp  As SldWorks.Component2
Dim swChildComp As SldWorks.Component2
Dim vChildComp  As Variant
Dim SelMgr      As SelectionMgr
Dim SelPart     As SldWorks.Component2
Dim Retval      As String
Dim intSelIndex As Integer
Dim intUndoCnt  As Integer

Const swDocPart = 1
Const swDocASSEMBLY = 2
Const swOpenDocOptions_Silent = 0
Const swComponentHidden = 0
Const swComponentVisible = 1
Const swComponentSuppressed = 0
Const swComponentFullyResolved = 2
Const swThisConfiguration = 1
Const ClickUnsup = 0
Const ClickSupp = 1
Const ClickShow = 2
Const ClickHide = 3
Const ClickDel = 4
Const Revision = "V1.00"

Function GetSelPart() As String
    Dim swSelect        As String
    swSelect = swAssem.GetTitle
    swSelect = Replace(swSelect, ".sldasm", "", , , vbTextCompare)
    swSelect = lstPartName.Column(0) & "@" & swSelect
    Debug.Print (swSelect)
    GetSelPart = swSelect
End Function

Sub TraverseComponent(swComp As SldWorks.Component2, nLevel As Long)
    Dim swCompConfing   As SldWorks.Configuration
    Dim sPadStr         As String
    Dim i               As Long
    Dim swCompState     As String
    Dim swVisibState    As String
    Dim intCount        As Integer
    Dim Retval          As String
    intCount = 0
    vChildComp = swComp.GetChildren
    
' ---------------------------------------------------------------------------------------
' Portions of this code are custom by Leonard Kikstra
' ---------------------------------------------------------------------------------------
    For i = 0 To UBound(vChildComp)
        ' Where are we in the list?
        intCount = Me.lstPartName.ListCount
        Set swChildComp = vChildComp(i)
        'TraverseComponent swChildComp, nLevel + 1
        Debug.Print sPadStr & swChildComp.Name2 & " <" & _
                                swChildComp.ReferencedConfiguration & "> " & swCompState
        ' Add component name and component's path to listbox to listbox
        Me.lstPartName.AddItem (swChildComp.Name2)
        Me.lstPartName.List(intCount, 1) = (swChildComp.GetPathName)
        ' Get and show suppression status
        
        swCompState = swChildComp.GetSuppression
        If swCompState = 0 Then
          Me.lstPartName.List(intCount, 2) = "Sup"
        Else
          Me.lstPartName.List(intCount, 2) = "Res"
        End If
        ' Get and show visibility status
        
        swVisibState = swChildComp.Visible
        If swVisibState = 0 Then
          Me.lstPartName.List(intCount, 3) = "Hid"
        Else
          Me.lstPartName.List(intCount, 3) = "Vis"
        End If
    
        If (swCompState <> 0 And OptionSuppressedOnly.Value = True) _
            Or (swVisibState <> 0 And OptionHiddenOnly.Value = True) _
            Then Me.lstPartName.RemoveItem intCount
    
    Next i
    Me.lstPartName.BoundColumn = 2
End Sub

Sub RefreshList()
'Refresh Parts List or end if no more parts.
    cmdState.Enabled = False
    cmdVisibility.Enabled = False
    Me.Image1.Picture = Nothing
    Me.Image1.Visible = False
    If intSelIndex = -1 Then
        End
    Else
        lstPartName.Clear
        TraverseComponent swRootComp, 1
        If lstPartName.ListCount <> 0 Then
            lstPartName.Selected(intSelIndex) = True
        End If
    End If
End Sub

Sub ChangePart(intBtn As Integer)
' ---------------------------------------------------------------------------------------
' Portions of this code are custom by Leonard Kikstra
' ---------------------------------------------------------------------------------------
'   If lstPartName.ListIndex = lstPartName.ListCount - 1 Then
'       intSelIndex = lstPartName.ListIndex - 1
'   Else
'       intSelIndex = lstPartName.ListIndex
'   End If
    Debug.Print lstPartName.ListCount
    Debug.Print intSelIndex
    Select Case intBtn
        Case 0
            'set part to resolved
            SelPart.SetSuppression (swComponentFullyResolved)
        Case 1
            'set part to suppressed
            SelPart.SetSuppression (swComponentSuppressed)
        Case 2
            'set part to shown
            SelPart.SetVisibility swComponentVisible, swThisConfiguration, ""
        Case 3
            'set part to hidden
            SelPart.SetVisibility swComponentHidden, swThisConfiguration, ""
        Case 4
           'Delete Part
            Retval = MsgBox("Are you sure you want to delete this part?", _
                     vbYesNo + vbDefaultButton2, "Delete Part")
            If Retval = vbYes Then
                swAssem.DeleteSelection (False)
                intUndoCnt = intUndoCnt + 1
                cmdUndo.Enabled = True
            End If
        End Select
'Call sub to refresh parts list
    RefreshList
End Sub

Private Sub cmdClose_Click()
   End
End Sub

Private Sub cmdState_Click()
' ---------------------------------------------------------------------------------------
' Portions of this code are custom by Leonard Kikstra
' ---------------------------------------------------------------------------------------
    If Me.lstPartName.List(Me.lstPartName.ListIndex, 2) = "Sup" Then
        ChangePart ClickUnsup
    Else
        ChangePart ClickSupp
    End If
'   lstPartName_Click
End Sub

Private Sub cmdVisibility_Click()
' ---------------------------------------------------------------------------------------
' Portions of this code are custom by Leonard Kikstra
' ---------------------------------------------------------------------------------------
    If Me.lstPartName.List(Me.lstPartName.ListIndex, 3) = "Hid" Then
        ChangePart ClickShow
    Else
        ChangePart ClickHide
    End If
'   lstPartName_Click
End Sub

Private Sub cmdDelete_Click()
    ChangePart ClickDel
End Sub

Private Sub cmdUndo_Click()
    swAssem.EditUndo2 (1)
    intUndoCnt = intUndoCnt - 1
    If intUndoCnt = 0 Then cmdUndo.Enabled = False
'Call sub to refresh parts list
    RefreshList
End Sub

Private Sub lstPartName_Click()
    Dim PreViewImg      As stdole.StdPicture
    Dim fileerror       As Long
    Dim filewarning     As Long
    Dim strSelPath      As String
    Dim strRefConfig    As String
    'Get full path of selected part from list box
    strSelPath = lstPartName.Column(1)
    ' Get Used Config
    Retval = swAssem.SelectByID(GetSelPart, "COMPONENT", 0#, 0#, 0#)
    Set SelPart = SelMgr.GetSelectedObject5(1)
    strRefConfig = SelPart.ReferencedConfiguration
    'Display preview image
    Set PreViewImg = swApp.GetPreviewBitmap(strSelPath, strRefConfig)
    Me.Image1.Picture = PreViewImg
    Me.Image1.Visible = True
    cmdState.Enabled = True
    cmdVisibility.Enabled = True
    cmdDelete.Enabled = False
    cmdUndo.Enabled = False
' ---------------------------------------------------------------------------------------
' Portions of this code are custom by Leonard Kikstra
' ---------------------------------------------------------------------------------------
    If Me.lstPartName.List(Me.lstPartName.ListIndex, 2) = "Sup" Then
      cmdState.Caption = "Resolve"
    Else
      cmdState.Caption = "Suppress"
    End If
    If Me.lstPartName.List(Me.lstPartName.ListIndex, 3) = "Hid" Then
      cmdVisibility.Caption = "Show"
    Else
      cmdVisibility.Caption = "Hide"
    End If
End Sub

Private Sub OptionHiddenOnly_Click()
'Call sub to refresh parts list
    RefreshList
End Sub

Private Sub OptionShowAll_Click()
'Call sub to refresh parts list
    RefreshList
End Sub

Private Sub OptionSuppressedOnly_Click()
'Call sub to refresh parts list
    RefreshList
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = Me.Caption & Revision
    Dim swConf      As SldWorks.Configuration
    Set swApp = CreateObject("SldWorks.Application")
    Set swAssem = swApp.ActiveDoc
    intUndoCnt = 0
'Check for open assembly document
    If (swAssem Is Nothing) Then
        Retval = MsgBox("Please open an assembly first!", , "Error")
        End
    End If
'Check document is an assembly
    If (swAssem.GetType <> swDocASSEMBLY) Then
        Retval = MsgBox("This macro only works on Assemblies!", , "Error")
        End
    End If
    Set SelMgr = swAssem.SelectionManager
    Set swConf = swAssem.GetActiveConfiguration
    Set swRootComp = swConf.GetRootComponent
    Debug.Print "File = " & swAssem.GetPathName
'Call sub to transverse feature tree and populate list box
    TraverseComponent swRootComp, 1
'Check that suppressed parts exist in assembly & if so select first in list
    If lstPartName.ListCount > 0 Then
        lstPartName.Selected(0) = True
    Else
        Retval = MsgBox("No suppressed parts found!", , "Error")
        End
    End If
' ---------------------------------------------------------------------------------------
' Portions of this code are custom by Leonard Kikstra
' ---------------------------------------------------------------------------------------
' Set status of buttons
    cmdState.Enabled = False
    cmdVisibility.Enabled = False
    cmdDelete.Enabled = False
    cmdUndo.Enabled = False
    OptionShowAll.Value = True
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).