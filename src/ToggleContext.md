# Toggle Context Macro

## Description
This macro toggles the editing context in SolidWorks between assembly and part environments, depending on the current context. It reselects the previously selected geometry after the context switch.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer  
- **Operating System**: Windows 7 or later  

## Pre-Conditions
> [!NOTE]
> - An assembly document must be active in SolidWorks.
> - At least one entity or component should be selected for optimal behavior.

## Results
> [!NOTE]
> - If the user is in the assembly editing context, the macro switches to the part editing context.
> - If the user is in the part editing context, the macro switches back to the assembly context.
> - The previously selected entity is reselected after the context change.

### Steps to Use the Macro
- Open an assembly document in SolidWorks.
- Select an entity or component in the assembly.
- Run the macro to toggle between assembly and part editing contexts.
- The previously selected entity will remain selected after the context switch.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

Sub main()
    On Error Resume Next ' Enable error handling to prevent runtime errors

    ' Declare SolidWorks application and document variables
    Dim swApp       As SldWorks.SldWorks ' SolidWorks application instance
    Dim swAssy      As SldWorks.AssemblyDoc ' Active assembly document
    Dim swDoc       As SldWorks.ModelDoc2 ' Generic model document
    Dim swFM        As SldWorks.FeatureManager ' Feature manager for the active document
    Dim swSelMgr    As SldWorks.SelectionMgr ' Selection manager for handling selections
    Dim swSelData   As SldWorks.SelectData ' Selection data
    Dim swComp      As SldWorks.Component2 ' Selected component
    Dim swEnt       As SldWorks.Entity ' Selected entity (generic object)
    Dim swSafeEnt   As SldWorks.Entity ' Safe entity reference for re-selection
    Dim status      As Long ' Status of operations

    ' Initialize SolidWorks application and active document
    Set swApp = Application.SldWorks
    Set swDoc = swApp.ActiveDoc

    ' Check if the active document is an assembly
    If swDoc.GetType <> swDocASSEMBLY Then Exit Sub ' Exit if no assembly document is open

    ' Initialize assembly, feature manager, and selection manager
    Set swAssy = swDoc
    Set swFM = swDoc.FeatureManager
    Set swSelMgr = swDoc.SelectionManager

    ' Retrieve the first selected object and its component
    Set swEnt = swSelMgr.GetSelectedObject6(1, -1) ' Get the first selected entity
    Set swComp = swSelMgr.GetSelectedObjectsComponent3(1, -1) ' Get the component containing the selection

    ' Get a safe reference to the selected entity for re-selection after context change
    If Not swEnt Is Nothing Then Set swSafeEnt = swEnt.GetSafeEntity

    ' Toggle context between assembly and part
    If swDoc.IsEditingSelf Then
        ' If currently in assembly context, enter part context
        If Not (swEnt Is Nothing And swComp Is Nothing) Then 
            swAssy.EditPart2 True, True, status ' Switch to part editing context
        End If
    Else
        ' If currently in part context, enter assembly context
        swAssy.EditAssembly ' Switch back to assembly context
    End If

    ' Re-select the previously selected entity in the new context (if applicable)
    If Not swEnt Is Nothing Then
        Set swSelData = swSelMgr.CreateSelectData ' Create a new selection data instance
        swSafeEnt.Select4 True, swSelData ' Re-select the entity
    End If

    ' Clean up all object references
    Set swApp = Nothing
    Set swDoc = Nothing
    Set swAssy = Nothing
    Set swSelMgr = Nothing
    Set swSelData = Nothing
    Set swEnt = Nothing
    Set swSafeEnt = Nothing
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).