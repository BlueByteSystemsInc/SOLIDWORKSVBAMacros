# Color Strip Macro for SolidWorks

## Description
This macro facilitates the removal of colors applied at various levels within SolidWorks documents, including assemblies and parts. It targets colors set at the assembly, part, feature, body, and face levels, ensuring that all components exhibit their default appearance.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - A part or assembly document must be open in SolidWorks.
> - The user can optionally select one or more assembly components if specific component color stripping is required.

## Results
> [!NOTE]
> - The macro removes colors from all selected or active components across multiple levels (assembly, part, feature, body, and face).
> - It ensures a clean, uniform appearance for all components by stripping away any custom color modifications.

## Steps to Setup the Macro

1. **Open SolidWorks**:
   - Ensure that a part or assembly document is open and, optionally, select specific components if targeted color removal is desired.

2. **Load and Run the Macro**:
   - Open the SolidWorks VBA editor by pressing (`Alt + F11`).
   - Insert a new module and copy the provided macro code into this module.
   - Run the macro from the VBA editor or by using **Tools** > **Macro** > **Run** within SolidWorks.

3. **Using the Macro**:
   - Upon running, the macro will automatically detect the type of the active document and begin the color removal process.
   - If specific components are selected, only those will be processed; otherwise, all components in the document will be targeted.
   - A confirmation message will appear once the color stripping is completed.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

' Constants to control which elements have their colors removed
Const CLEAR_ASSY As Boolean = True ' Clear colors from assemblies
Const CLEAR_PART As Boolean = True ' Clear colors from parts
Const CLEAR_BODY As Boolean = True ' Clear colors from bodies
Const CLEAR_FEAT As Boolean = True ' Clear colors from features
Const CLEAR_FACE As Boolean = True ' Clear colors from faces
Const SAVE_FLAG As Boolean = False ' Save the document after clearing colors
Const OK_FLAG As Boolean = True    ' Show completion message

' SolidWorks application and component list variables
Dim swApp As SldWorks.SldWorks
Dim swCompList As Object

' Main subroutine
Sub main()
    Dim swDoc As SldWorks.ModelDoc2

    ' Initialize SolidWorks application and active document
    Set swApp = Application.SldWorks
    Set swDoc = swApp.ActiveDoc
    Set swCompList = CreateObject("Scripting.Dictionary") ' Initialize component list

    ' Check the document type and clear colors accordingly
    If swDoc.GetType = swDocPART Then
        ClearPartColors swDoc ' Clear part colors
    ElseIf swDoc.GetType = swDocASSEMBLY Then
        ClearAssyColors swDoc ' Clear assembly colors
    End If

    ' Show completion message if OK_FLAG is set
    If OK_FLAG Then MsgBox "Color Stripping Completed", vbOKOnly
End Sub

' Subroutine to clear colors from assemblies and their components
Private Sub ClearAssyColors(swDoc As SldWorks.ModelDoc2)
    On Error Resume Next
    Dim swAssy As SldWorks.AssemblyDoc
    Dim swComp As SldWorks.Component2
    Dim swComponents As Variant
    Dim swSelMgr As SldWorks.SelectionMgr
    Dim SelCount As Integer
    Dim i As Integer
    Dim Errors As Long, Warnings As Long

    ' Cast the document to an assembly
    Set swAssy = swDoc
    Set swSelMgr = swDoc.SelectionManager
    SelCount = swSelMgr.GetSelectedObjectCount

    ' Get components either from selection or the entire assembly
    If SelCount < 1 Then
        swComponents = swAssy.GetComponents(True)
        If IsEmpty(swComponents) Then Exit Sub
    Else
        ReDim swComponents(SelCount - 1)
        For i = 0 To SelCount - 1
            Set swComp = swSelMgr.GetSelectedObjectsComponent3(i + 1, -1)
            Set swComponents(i) = swComp
        Next i
    End If

    ' Iterate over components to remove colors
    For i = 0 To UBound(swComponents)
        Set swComp = swComponents(i)
        If CLEAR_ASSY Then swComp.RemoveMaterialProperty2 swAllConfiguration, Empty
        Dim swDoc2 As SldWorks.ModelDoc2
        Set swDoc2 = swComp.GetModelDoc2
        ' Recursively clear colors for subassemblies and parts
        If swDoc2.GetType = swDocASSEMBLY Then
            ClearAssyColors swDoc2
        ElseIf swDoc2.GetType = swDocPART Then
            ClearPartColors swDoc2
        End If
        ' Save the document if SAVE_FLAG is enabled
        If SAVE_FLAG Then swDoc2.Save3 swSaveAsOptions_Silent, Errors, Warnings
    Next i

    ' Save the main assembly document if SAVE_FLAG is enabled
    If SAVE_FLAG Then swDoc.Save3 swSaveAsOptions_Silent, Errors, Warnings
End Sub

' Subroutine to clear colors from parts, bodies, features, and faces
Private Sub ClearPartColors(swDoc As SldWorks.ModelDoc2)
    Dim swPart As SldWorks.PartDoc
    Set swPart = swDoc

    ' Remove colors from the part level
    If CLEAR_PART Then swDoc.Extension.RemoveMaterialProperty swAllConfiguration, Empty

    ' Remove colors from bodies
    Dim swBodies As Variant
    swBodies = swPart.GetBodies2(swAllBodies, False)
    If IsEmpty(swBodies) Then Exit Sub
    Dim swBody As SldWorks.Body2
    For Each swBody In swBodies
        If CLEAR_BODY Then swBody.RemoveMaterialProperty swAllConfiguration, Empty
        ' Remove colors from faces
        If CLEAR_FACE Then
            Dim swFace As SldWorks.Face2
            For Each swFace In swBody.GetFaces
                swFace.RemoveMaterialProperty2 swAllConfiguration, Empty
            Next
        End If
    Next

    ' Remove colors from features
    If CLEAR_FEAT Then
        Dim swFeat As SldWorks.Feature
        Set swFeat = swPart.FirstFeature
        Do While Not swFeat Is Nothing
            swFeat.RemoveMaterialProperty2 swAllConfiguration, Empty
            Set swFeat = swFeat.GetNextFeature
        Loop
    End If
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).