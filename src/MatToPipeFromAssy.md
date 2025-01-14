# Material Assignment Macro for Pipes and Tubes in SolidWorks Assemblies

## Description
This macro automates the process of assigning materials to pipe and tube components within a SolidWorks assembly. It iterates through all components, identifies pipes and tubes, and applies a predefined material to each.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - An assembly document must be open in SolidWorks.
> - The assembly should contain pipe and tube components that are compatible with the Routing component manager.

## Results
> [!NOTE]
> - Assigns the material 'AISI 304' to all pipe and tube components found in the assembly.
> - Updates each configuration of the parts with the selected material.

## Steps to Setup the Macro

1. **Create the VBA Module**:
   - Open the SolidWorks VBA editor by pressing (`Alt + F11`).
   - Insert a new module into your project and copy the provided macro code into this module.

2. **Run the Macro**:
   - Make sure an assembly with pipe and tube components is open.
   - Execute the macro from within SolidWorks by navigating to **Tools** > **Macro** > **Run**, then select your macro file.

3. **Using the Macro**:
   - Upon execution, the macro automatically processes the active assembly.
   - Each pipe and tube component will have the material 'AISI 304' assigned across all configurations.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

' Declaration of SolidWorks application and document variables
Dim swApp As SldWorks.SldWorks                          ' SolidWorks application object
Dim swAssyDoc As SldWorks.AssemblyDoc                   ' Assembly document object
Dim swComponent As SldWorks.Component2                  ' Individual assembly component object
Dim swRtCompMgr As SWRoutingLib.RoutingComponentManager ' Routing component manager for determining component type

Dim swModelDoc As SldWorks.ModelDoc2                    ' Active SolidWorks document
Dim swModelDocComp As SldWorks.ModelDoc2                ' Component document within the assembly
Dim swModelDocExt As SldWorks.ModelDocExtension         ' Model extension object for accessing extended functionalities

Dim swPipePart As SldWorks.PartDoc                      ' Part document for pipe components
Dim swTubePart As SldWorks.PartDoc                      ' Part document for tube components

Dim vComponents As Variant                              ' Array of components in the assembly
Dim vComponent As Variant                               ' Individual component in the array
Dim sComponentName0 As String                           ' Current component name
Dim sComponentName1 As String                           ' Previous component name (used for comparison)

Dim sArrayPipeNames() As String                         ' Array to store pipe component names
Dim sArrayTubeNames() As String                         ' Array to store tube component names

Dim vConfigNames As Variant                             ' Array of configuration names in a part
Dim sConfigName As String                               ' Individual configuration name
Dim swConfig As SldWorks.Configuration                  ' Configuration object
Dim sMaterialDataBase As String                         ' Path to the material database
Dim sMaterialName As String                             ' Name of the material to assign

Dim i As Integer                                        ' Loop counter for components
Dim p As Integer                                        ' Counter for pipe components
Dim t As Integer                                        ' Counter for tube components

' Main subroutine to process assembly components
Sub main()
    ' Initialize SolidWorks application and active document
    Set swApp = Application.SldWorks
    Set swModelDoc = swApp.ActiveDoc

    ' Check if a document is open
    If swModelDoc Is Nothing Then
        MsgBox "No document loaded."
        End
    End If

    p = 1  ' Initialize pipe counter
    t = 1  ' Initialize tube counter

    ' Check if the current document is an assembly
    If swModelDoc.GetType = 2 Then
        ' Cast the document as an assembly
        Set swAssyDoc = swModelDoc

        ' Get all components in the assembly
        vComponents = swAssyDoc.GetComponents(False)

        ' Initialize arrays to store component names
        ReDim sArrayPipeNames(1)
        ReDim sArrayTubeNames(1)

        ' Process pipe components
        For i = 0 To UBound(vComponents)
            Set swComponent = vComponents(i)
            Set swModelDocComp = swComponent.GetModelDoc2

            If Not swModelDocComp Is Nothing Then
                If swModelDocComp.GetType = 1 Then ' Check if the component is a part
                    Set swModelDocExt = swModelDocComp.Extension
                    Set swRtCompMgr = swModelDocExt.GetRoutingComponentManager

                    If swRtCompMgr.GetComponentType = 1 Then ' Check if the part is a pipe
                        Set swPipePart = swComponent.GetModelDoc2
                        sComponentName0 = swComponent.GetPathName
                        sArrayPipeNames(p) = sComponentName0

                        If sComponentName0 <> sComponentName1 Then
                            p = p + 1
                            ReDim Preserve sArrayPipeNames(UBound(sArrayPipeNames) + 1)
                            Call AssigMaterial(swPipePart) ' Assign material to the pipe
                        End If

                        sComponentName1 = sComponentName0
                    End If
                End If
            End If
        Next i

        ' Process tube components
        For i = 0 To UBound(vComponents)
            Set swComponent = vComponents(i)
            Set swModelDocComp = swComponent.GetModelDoc2

            If Not swModelDocComp Is Nothing Then
                If swModelDocComp.GetType = 1 Then ' Check if the component is a part
                    Set swModelDocExt = swModelDocComp.Extension
                    Set swRtCompMgr = swModelDocExt.GetRoutingComponentManager

                    If swRtCompMgr.GetComponentType = 25 Then ' Check if the part is a tube
                        Set swTubePart = swComponent.GetModelDoc2
                        sComponentName0 = swComponent.GetPathName
                        sArrayTubeNames(t) = sComponentName0

                        If sComponentName0 <> sComponentName1 Then
                            t = t + 1
                            ReDim Preserve sArrayTubeNames(UBound(sArrayTubeNames) + 1)
                            Call AssigMaterial(swTubePart) ' Assign material to the tube
                        End If

                        sComponentName1 = sComponentName0
                    End If
                End If
            End If
        Next i
    ElseIf swModelDoc.GetType <> 2 Then
        MsgBox "Current Document must be Assembly."
        End
    End If
End Sub

' Subroutine to assign material to parts
Sub AssigMaterial(swPipePart As SldWorks.PartDoc)
    Dim j As Integer                                    ' Loop counter for configurations

    ' Get all configuration names in the part
    vConfigNames = swPipePart.GetConfigurationNames

    ' Define material database path and material name
    sMaterialDataBase = "C:/Program Files/SOLIDWORKS Corp/SolidWorks (2)/lang/english/sldmaterials/SolidWorks Materials.sldmat"
    sMaterialName = "AISI 304"

    ' Assign the material to all configurations
    For j = 0 To UBound(vConfigNames)
        sConfigName = vConfigNames(j)
        Call swPipePart.SetMaterialPropertyName2(sConfigName, sMaterialDataBase, sMaterialName)
    Next j
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).