# Exploded View Toggle Macro for SolidWorks

## Description
This macro automates the process of toggling the exploded view state of an active assembly in SolidWorks. It checks if an exploded view exists and then either explodes or collapses it based on its current state.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - An assembly document must be open in SolidWorks.
> - The assembly should have at least one exploded view configured.

## Results
> [!NOTE]
> - Toggles the state of the first exploded view found in the active configuration of the assembly.
> - Either explodes or collapses the assembly view depending on its initial state.

## Steps to Setup the Macro

1. **Create the VBA Module**:
   - Open the SolidWorks VBA editor by pressing (`Alt + F11`).
   - Insert a new module into your project.

2. **Run the Macro**:
   - Ensure that an assembly with an exploded view is open.
   - Execute the macro from within SolidWorks by navigating to **Tools** > **Macro** > **Run**, then select your macro file.

3. **Using the Macro**:
   - The macro checks for the presence of an exploded view and changes its state.
   - No user interaction is required apart from initiating the macro.

## VBA Macro Code

### Main Subroutine
```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

' Declare variables for SolidWorks objects and operations
Dim swApp As SldWorks.SldWorks                   ' SolidWorks application object
Dim swModel As SldWorks.ModelDoc2                ' Active document object
Dim swModelDocExt As SldWorks.ModelDocExtension  ' ModelDocExtension object
Dim swAssembly As SldWorks.AssemblyDoc           ' Assembly document object
Dim swConfigMgr As SldWorks.ConfigurationManager ' Configuration manager object
Dim swConfig As SldWorks.Configuration           ' Active configuration object
Dim activeConfigName As String                   ' Name of the active configuration
Dim viewNames As Variant                         ' Array of exploded view names
Dim viewName As String                           ' Individual exploded view name
Dim i As Long                                    ' Loop counter
Dim xViewCount As Long                           ' Number of exploded views in the configuration
Dim boolstatus As Boolean                        ' Boolean for operation success

' Main subroutine
Sub main()
    ' Initialize the SolidWorks application and active document objects
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    
    ' Ensure the active document is an assembly
    If (swModel.GetType <> swDocASSEMBLY) Then Exit Sub
    
    ' Cast the active document as an assembly
    Set swAssembly = swModel
    
    ' Get the active configuration name
    Set swConfigMgr = swModel.ConfigurationManager
    Set swConfig = swConfigMgr.ActiveConfiguration
    activeConfigName = swConfig.Name
    
    ' Get the number of exploded views in the active configuration
    xViewCount = swAssembly.GetExplodedViewCount2(activeConfigName)
    If xViewCount < 1 Then End ' Exit if no exploded views exist
    
    ' Retrieve the names of the exploded views in the active configuration
    viewNames = swAssembly.GetExplodedViewNames2(activeConfigName)
    
    ' Select the first exploded view in the list
    boolstatus = swAssembly.Extension.SelectByID2(viewNames(0), "EXPLODEDVIEWS", 0, 0, 0, False, 0, Nothing, 0)
    If boolstatus = False Then End ' Exit if the selection fails
    
    ' Check if the assembly is currently exploded or collapsed
    boolstatus = swAssembly.IsExploded
    
    ' Toggle the explode/collapse state based on the current state
    If boolstatus = False Then
        swAssembly.ViewExplodeAssembly ' Explode the assembly
    Else
        swAssembly.ViewCollapseAssembly ' Collapse the assembly
    End If
    
    ' Clear the selection to finish the operation
    swAssembly.ClearSelection2 True
    
    ' Clean up references and end the macro
    Set swModel = Nothing
    Set swAssembly = Nothing
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).