# Cycle References Macro

## Description
This macro cycles through the selection of reference geometry (planes, axes, origin) of a component within an assembly in SolidWorks. It allows for easy selection of features to use for mating components within assemblies. Running the macro again will cycle through all the reference features of the last selected component.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer  
- **Operating System**: Windows 7 or later  

## Pre-Conditions
> [!NOTE]
> - An assembly document must be currently open in SolidWorks.
> - One or more items must be selected either in the graphics area or the feature tree.

## Results
> [!NOTE]
> - Selects one of the reference features (plane, axis, or origin) of the last selected component.
> - Cycles through all available reference features with each subsequent macro execution.

## Steps to Setup the Macro

### 1. **Check Document Type**:
   - Ensure that an assembly document is active. Exit the macro if another type of document is active.

### 2. **Get Selected Component**:
   - Retrieve the last selected component from the selection manager. Exit the macro if no components are selected.

### 3. **Initialize Feature Cycling**:
   - Start with the first feature of the selected component and check each feature to see if it is a reference feature (plane or axis) or the origin.
   - Add valid reference features to a collection for cycling.

### 4. **Determine Current Selection**:
   - Check if the current selection matches any features in the collection. If it does, prepare to select the next feature in the cycle.

### 5. **Cycle Through Features**:
   - If the current feature is the last in the collection, wrap around to the first feature.
   - Select the determined next feature in the SolidWorks assembly.

### 6. **Handle Origin Selection**:
   - If the origin is allowed and selected, and if stopping at the origin is enabled, exit the macro after selecting the origin.
   - Otherwise, continue to the next reference feature.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

' Constants for controlling feature selection behavior
Const STOPATORIGIN As Boolean = False       ' Stop cycling at the origin feature
Const FIRSTREF As Long = 1                  ' Start cycling with the first reference feature
Const SELECTAXIS As Boolean = True          ' Allow selection of reference axes
Const SELECTORIGIN As Boolean = True        ' Allow selection of the origin

Sub main()
    ' SolidWorks application and document objects
    Dim swApp As SldWorks.SldWorks          ' SolidWorks application object
    Dim swModel As SldWorks.ModelDoc2       ' Active document object
    Dim swSelMgr As SldWorks.SelectionMgr   ' Selection manager object
    Dim swSelComp As SldWorks.Component2    ' Selected component object
    Dim swFeat As SldWorks.Feature          ' Feature object for traversing component features
    Dim GeneralSelObj As Object             ' General object for the last selected entity
    Dim myFeatureCollection As New Collection ' Collection to store reference features
    Dim i As Integer                        ' Loop counter
    Dim CurSelCount As Long                 ' Current selection count

    ' Initialize SolidWorks application and active document
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    ' Ensure an assembly document is open
    If swModel Is Nothing Or swModel.GetType <> swDocASSEMBLY Then
        MsgBox "This macro only works on assembly documents.", vbExclamation, "Invalid Document Type"
        Exit Sub
    End If

    ' Initialize selection manager and get count of selected items
    Set swSelMgr = swModel.SelectionManager
    CurSelCount = swSelMgr.GetSelectedObjectCount
    If CurSelCount = 0 Then
        MsgBox "No items selected.", vbInformation, "No Selection"
        Exit Sub
    End If

    ' Cycle through selected components and gather their reference features
    For i = 1 To CurSelCount
        Set swSelComp = swSelMgr.GetSelectedObjectsComponent(i) ' Get the component for each selection
        If Not swSelComp Is Nothing Then
            Set swFeat = swSelComp.FirstFeature ' Access the first feature of the component
            Do While Not swFeat Is Nothing
                ' Add reference features (planes, axes, origin) to the collection
                Select Case swFeat.GetTypeName
                    Case "RefPlane" ' Add reference planes to the collection
                        myFeatureCollection.Add swFeat
                    Case "RefAxis"  ' Add reference axes if enabled
                        If SELECTAXIS Then myFeatureCollection.Add swFeat
                    Case "OriginProfileFeature" ' Add origin if enabled
                        If SELECTORIGIN Then myFeatureCollection.Add swFeat
                End Select
                Set swFeat = swFeat.GetNextFeature ' Move to the next feature
            Loop
        End If
    Next i

    ' Determine the next feature to select
    Set GeneralSelObj = swSelMgr.GetSelectedObject6(CurSelCount, -1) ' Get the last selected object
    For i = 1 To myFeatureCollection.Count
        If GeneralSelObj Is myFeatureCollection.Item(i) Then
            ' Cycle to the next feature in the collection
            Set GeneralSelObj = myFeatureCollection.Item((i Mod myFeatureCollection.Count) + 1)
            Exit For
        End If
    Next

    ' Select the next feature in the collection
    If Not GeneralSelObj Is Nothing Then
        GeneralSelObj.Select4 True, Nothing, False
    End If
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).