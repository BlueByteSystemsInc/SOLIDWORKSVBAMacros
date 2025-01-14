# Derived Configuration and Tolerance Adjustment Macro for SolidWorks

## Description
This macro facilitates the creation of derived configurations in SolidWorks parts or assemblies, where each configuration is adjusted based on the minimum and maximum tolerances of dimensions. It's designed to automate the process of setting dimensions to their extreme tolerances for testing or analysis purposes.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - A part or assembly must be open.
> - Optionally, the part or assembly may have dimensions with tolerances set.

## Results
> [!NOTE]
> - Creates derived configurations based on the current configuration.
> - Sets all dimensions within the derived configurations to their minimum and maximum tolerances.
> - Removes all tolerances from dimensions after setting.
> - Activates the derived configuration as the current configuration.

## Steps to Setup the Macro

1. **Create the VBA Module**:
   - Open the SolidWorks VBA editor by pressing (`Alt + F11`).
   - Insert a new module into your project.

2. **Run the Macro**:
   - Ensure a part or assembly is open and optionally has dimensions with tolerances.
   - Execute the macro from within SolidWorks by navigating to **Tools** > **Macro** > **Run**, then select your macro file.

3. **Using the Macro**:
   - The macro processes the currently active configuration.
   - It creates two new configurations: one with all dimensions set to their minimum tolerance and one set to maximum.
   - The new configurations are then cleaned of any dimensional tolerances.

## VBA Macro Code

### Main Subroutine
```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).
Option Explicit

' Declare the SolidWorks application object
Dim swApp As SldWorks.SldWorks

' Main subroutine that manages the workflow of the macro
Sub main()
    Dim swModel As SldWorks.ModelDoc2            ' Active document object
    Dim swComp As SldWorks.Component2            ' Assembly component object
    Dim i As Integer                             ' Counter for selected components
    
    ' Initialize SolidWorks application
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    
    ' Determine if the active document is a part or an assembly
    If swModel.GetType = swDocPART Then
        ' Process the active part document
        ProcessPart swModel
    ElseIf swModel.GetType = swDocASSEMBLY Then
        ' Handle assemblies by processing selected components
        Dim swSelMgr As SldWorks.SelectionMgr
        Set swSelMgr = swModel.SelectionManager
        
        ' Loop through all selected objects in the assembly
        For i = 1 To swSelMgr.GetSelectedObjectCount2(-1)
            Set swComp = swSelMgr.GetSelectedObjectsComponent4(i, -1)
            If Not swComp Is Nothing Then
                ' Process each selected component as a part
                ProcessPart swComp.GetModelDoc2
            End If
        Next
    Else
        ' Prompt the user to open a part or assembly if none are active
        MsgBox "Please open a part or assembly first."
    End If
End Sub

' Subroutine to process part documents
Sub ProcessPart(ByRef swModel As SldWorks.ModelDoc2)
    Dim swConf As SldWorks.Configuration          ' Active configuration object
    Dim swConfMgr As SldWorks.ConfigurationManager ' Configuration manager object
    
    ' Get the active configuration and configuration manager
    Set swConf = swModel.GetActiveConfiguration
    Set swConfMgr = swModel.ConfigurationManager
    
    ' Create derived configurations for minimum and maximum tolerances
    CreateDerivedConfig swModel, 0  ' Min tolerance
    swModel.ShowConfiguration2 swConf.Name        ' Switch back to the active configuration
    CreateDerivedConfig swModel, 1  ' Max tolerance
End Sub

' Subroutine to create derived configurations for tolerances
Sub CreateDerivedConfig(ByRef swModel As SldWorks.ModelDoc2, pass As Integer)
    Dim swFeat As SldWorks.Feature                 ' Feature object
    Dim swDispDim As SldWorks.DisplayDimension     ' Display dimension object
    Dim swDim As SldWorks.Dimension               ' Dimension object
    Dim bRet As Boolean                           ' Boolean for operation success
    Dim swMidConf As SldWorks.Configuration       ' New configuration object
    Dim swConf As SldWorks.Configuration          ' Current configuration object
    Dim swConfMgr As SldWorks.ConfigurationManager ' Configuration manager object
    
    ' Get the current active configuration
    Set swConf = swModel.GetActiveConfiguration
    Set swConfMgr = swModel.ConfigurationManager
    
    ' Avoid creating derived configurations if already in tolerance-specific configurations
    If InStr(1, swConf.Name, " - max tolerance") Or InStr(1, swConf.Name, " - min tolerance") Then
        Exit Sub
    End If
    
    ' Create a new derived configuration for the specified tolerance (min/max)
    Set swMidConf = swConfMgr.AddConfiguration( _
        swConf.Name & " - " & IIf(pass = 0, "min", "max") & " tolerance", _
        IIf(pass = 0, "min", "max") & " tolerance", _
        IIf(pass = 0, "min", "max") & " tolerance", _
        1, _
        swConf.Name, _
        IIf(pass = 0, "min", "max") & " tolerance")
    If swMidConf Is Nothing Then Exit Sub
    
    ' Process features and dimensions in the model for tolerance adjustments
    Set swFeat = swModel.FirstFeature
    Debug.Print "File = " & swModel.GetPathName
    Debug.Print "  Nominal Tolerance:"
    ProcessMassProps swModel   ' Log mass properties for nominal tolerance
    Debug.Print "  -----------------------------"
    
    ' Loop through each feature in the model
    Do While Not swFeat Is Nothing
        Debug.Print "  " & swFeat.Name
        Set swDispDim = swFeat.GetFirstDisplayDimension
        ' Loop through dimensions in the feature
        Do While Not swDispDim Is Nothing
            Set swDim = swDispDim.GetDimension
            ProcessDimension swModel, swDim
            SetDimensionToMidTolerance swModel, swDim, pass
            Set swDispDim = swFeat.GetNextDisplayDimension(swDispDim)
        Loop
        Set swFeat = swFeat.GetNextFeature
    Loop
    
    ' Rebuild the model to apply changes
    bRet = swModel.ForceRebuild3(False): Debug.Assert bRet
    Debug.Print "  Middle Tolerance:"
    ProcessMassProps swModel   ' Log mass properties for middle tolerance
    Debug.Print "  -----------------------------"
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).