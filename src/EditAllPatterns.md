# Edit Linear Patterns in SolidWorks Macro

## Description
This macro allows users to edit linear patterns in parts and assemblies directly within SolidWorks. The macro searches for linear patterns and provides a PropertyManager Page interface to modify instance counts and spacing for each direction. It supports both part-level and assembly-level patterns.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer  
- **Operating System**: Windows 7 or later  

## Pre-Conditions
> [!NOTE]
> - A part or assembly document must be active in SolidWorks.
> - Linear patterns must exist in the feature tree of the active document.

## Results
> [!NOTE]
> - Identifies linear patterns in the active document.
> - Displays a PropertyManager Page for editing pattern parameters (instances and spacing).
> - Updates the document based on user modifications.

## Steps to Setup the Macro

### 1. **Prepare the Document**:
   - Open the part or assembly containing linear patterns.

### 2. **Run the Macro**:
   - Execute the macro. It will detect linear patterns in the feature tree and provide an interface to edit them.

### 3. **Edit Patterns**:
   - Use the PropertyManager Page to adjust spacing and instance counts for the linear patterns.

## VBA Macro Code

### Main Macro

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

' Declare global variables for SolidWorks application and pattern features
Public swApp As SldWorks.SldWorks                  ' SolidWorks application object
Dim swPatternFeats() As SldWorks.Feature           ' Array for linear pattern features in a part
Dim swComponentPatterns() As SldWorks.Feature      ' Array for component patterns in an assembly
Dim pm_page As EditPatternPropertyPage             ' Property manager page instance
Dim swPatStep As Integer                           ' Counter for patterns in part
Dim swCompStep As Integer                          ' Counter for patterns in assembly
Dim isPart As Boolean                              ' Boolean to check if the document is a part
Dim mainAssy As String                             ' Path of the main assembly file

' Main subroutine
Sub main()
    Dim swPart As SldWorks.ModelDoc2
    Dim docType As Integer

    ' Initialize SolidWorks application and get the active document
    Set swApp = Application.SldWorks
    Set swPart = swApp.ActiveDoc

    ' Check if a document is open
    If swPart Is Nothing Then
        MsgBox "No active file found. Please open a part or assembly document and try again.", vbCritical, "Error"
        Exit Sub
    End If

    ' Check document type (part, assembly, or drawing)
    docType = swPart.GetType
    If docType = swDocDRAWING Then
        MsgBox "This macro does not support drawings. Open a part or assembly and try again.", vbCritical, "Error"
        Exit Sub
    End If

    ' Determine if the active document is a part or assembly
    isPart = (docType = swDocPART)
    If Not isPart Then
        mainAssy = swPart.GetPathName
    End If

    ' Rebuild the active document
    swPart.ForceRebuild3 True

    ' Create an instance of the property manager page
    Set pm_page = New EditPatternPropertyPage

    ' Traverse the document for pattern features
    Call Change_Patterns(swPart)
End Sub

' Subroutine to find and process patterns in parts or assemblies
Sub Change_Patterns(swPart As SldWorks.ModelDoc2)
    If swPart.GetType = swDocPART Then
        ' Handle patterns in a part document
        swPatStep = 0
        swPatternFeats = TraverseFeatureFeatures(swPart) ' Get linear patterns
        If UBound(swPatternFeats) >= 0 Then              ' Check if patterns were found
            Call ShowPropPage("PART")                    ' Display PMP for parts
        End If
    Else
        ' Handle patterns in an assembly document
        mainAssy = swPart.GetPathName
        swComponentPatterns = TraverseAssemblyFeatures(swPart) ' Get component patterns
        If UBound(swComponentPatterns) >= 0 Then               ' Check if patterns were found
            Call ShowPropPage("ASSEMBLY")                      ' Display PMP for assemblies
        End If
    End If
End Sub

' Function to traverse features in a part and find linear patterns
Function TraverseFeatureFeatures(swPart As SldWorks.ModelDoc2) As Variant
    Dim PatternFeats() As SldWorks.Feature           ' Array to store pattern features
    Dim tFeat As SldWorks.Feature                    ' Temporary feature object
    Dim isPattern As Boolean                         ' Flag to check if patterns are found

    ' Start with the first feature in the part
    Set tFeat = swPart.FirstFeature
    Do While Not tFeat Is Nothing
        ' Check if the feature is a linear pattern
        If tFeat.GetTypeName = "LPattern" Then
            ' Add the pattern to the array
            If Not isPattern Then
                ReDim PatternFeats(0)
            Else
                ReDim Preserve PatternFeats(UBound(PatternFeats) + 1)
            End If
            Set PatternFeats(UBound(PatternFeats)) = tFeat
            isPattern = True
        End If
        ' Move to the next feature
        Set tFeat = tFeat.GetNextFeature
    Loop

    ' Return the array of patterns or an empty array if none are found
    If isPattern Then
        TraverseFeatureFeatures = PatternFeats
    Else
        TraverseFeatureFeatures = Array()
    End If
End Function
```

### PropertyManager Page Class

```vbnet
Option Explicit

Implements PropertyManagerPage2Handler5

' PropertyManager Page and group variables
Dim BoundingBoxPropPage As SldWorks.PropertyManagerPage2  ' Main PMP
Dim grpOldPatternBox As SldWorks.PropertyManagerPageGroup ' Group for old pattern
Dim grpNewPatternBox As SldWorks.PropertyManagerPageGroup ' Group for new pattern

' Controls for old pattern properties
Dim lbld1A As SldWorks.PropertyManagerPageLabel           ' Label for Dir 1 Spacing
Dim numOldDir1Amount As SldWorks.PropertyManagerPageNumberbox ' NumberBox for Dir 1 Spacing
Dim lbld1N As SldWorks.PropertyManagerPageLabel           ' Label for Dir 1 Instances
Dim numOldDir1Number As SldWorks.PropertyManagerPageNumberbox ' NumberBox for Dir 1 Instances
Dim lbld2A As SldWorks.PropertyManagerPageLabel           ' Label for Dir 2 Spacing
Dim numOldDir2Amount As SldWorks.PropertyManagerPageNumberbox ' NumberBox for Dir 2 Spacing
Dim lbld2N As SldWorks.PropertyManagerPageLabel           ' Label for Dir 2 Instances
Dim numOldDir2Number As SldWorks.PropertyManagerPageNumberbox ' NumberBox for Dir 2 Instances

' Public variables for storing pattern properties
Public dir1OldAmount As Double, dir1OldNumber As Double   ' Old Dir 1 properties
Public dir2OldAmount As Double, dir2OldNumber As Double   ' Old Dir 2 properties
Public dir1NewAmount As Double, dir1NewNumber As Double   ' New Dir 1 properties
Public dir2NewAmount As Double, dir2NewNumber As Double   ' New Dir 2 properties
Public PatternName As String                              ' Name of the pattern

' PMP Initialization: Creates the UI
Private Sub Class_Initialize()
    Dim options As Long
    Dim longerrors As Long

    ' Options for PMP (e.g., OK, Cancel buttons)
    options = swPropertyManager_OkayButton + swPropertyManager_CancelButton + SwConst.swPropertyManagerOptions_CanEscapeCancel

    ' Create the PMP
    Set BoundingBoxPropPage = swApp.CreatePropertyManagerPage("Edit Linear Patterns", options, Me, longerrors)

    ' Add the Old Pattern Group
    Set grpOldPatternBox = BoundingBoxPropPage.AddGroupBox(1200, "Current Pattern Amounts", swGroupBoxOptions_Visible + swGroupBoxOptions_Expanded)
    
    ' Initialize controls for old pattern properties
    InitOldPatternControls
End Sub

' Initialize controls for the Old Pattern group
Private Sub InitOldPatternControls()
    Dim ju As Integer, lblType As Integer, conAlign As Integer, numType As Integer

    ju = SwConst.swControlOptions_Visible + SwConst.swControlOptions_Enabled
    lblType = SwConst.swPropertyManagerPageControlType_e.swControlType_Label
    conAlign = SwConst.swControlAlign_LeftEdge
    numType = SwConst.swPropertyManagerPageControlType_e.swControlType_Numberbox

    ' Direction 1 Spacing
    Set lbld1A = grpOldPatternBox.AddControl(1210, lblType, "Dir 1 Spacing", conAlign, ju, "")
    Set numOldDir1Amount = grpOldPatternBox.AddControl(1220, numType, "Distance", conAlign, ju, "Distance")
    numOldDir1Amount.Style = swPropMgrPageNumberBoxStyle_e.swPropMgrPageNumberBoxStyle_NoScrollArrows

    ' Direction 1 Instances
    Set lbld1N = grpOldPatternBox.AddControl(1230, lblType, "Dir 1 Instances", conAlign, ju, "")
    Set numOldDir1Number = grpOldPatternBox.AddControl(1240, numType, "Distance", conAlign, ju, "Distance")
    numOldDir1Number.SetRange2 swNumberBox_UnitlessInteger, 1, 100, True, 1, 1, 1
    numOldDir1Number.Style = swPropMgrPageNumberBoxStyle_e.swPropMgrPageNumberBoxStyle_AvoidSelectionText
    numOldDir1Number.Value = 1

    ' Direction 2 Spacing
    Set lbld2A = grpOldPatternBox.AddControl(1250, lblType, "Dir 2 Spacing", conAlign, ju, "")
    Set numOldDir2Amount = grpOldPatternBox.AddControl(1260, numType, "Distance", conAlign, ju, "Distance")
    numOldDir2Amount.Style = swPropMgrPageNumberBoxStyle_e.swPropMgrPageNumberBoxStyle_NoScrollArrows

    ' Direction 2 Instances
    Set lbld2N = grpOldPatternBox.AddControl(1270, lblType, "Dir 2 Instances", conAlign, ju, "")
    Set numOldDir2Number = grpOldPatternBox.AddControl(1280, numType, "Distance", conAlign, ju, "Distance")
    numOldDir2Number.SetRange2 swNumberBox_UnitlessInteger, 1, 100, True, 1, 1, 1
    numOldDir2Number.Value = 1
End Sub

' Populate values in the PMP
Private Sub GetMateValues()
    BoundingBoxPropPage.Title = "Edit [" & Me.PatternName & "]"
    numOldDir1Amount.Value = dir1OldAmount
    numOldDir2Amount.Value = dir2OldAmount
    numOldDir1Number.Value = dir1OldNumber
    numOldDir2Number.Value = dir2OldNumber
End Sub

' Event: Called when PMP closes
Private Sub PropertyManagerPage2Handler5_OnClose(ByVal Reason As Long)
    dir1OldAmount = numOldDir1Amount.Value
    dir2OldAmount = numOldDir2Amount.Value
    dir1OldNumber = numOldDir1Number.Value
    dir2OldNumber = numOldDir2Number.Value
End Sub

' Display the PropertyManager Page
Sub Show()
    Call GetMateValues
    BoundingBoxPropPage.Show
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).