# Match Text Properties in SolidWorks Drawing

## Description
This macro matches text properties such as height and font of the selected text to a parent text object in a SolidWorks drawing. The macro enables users to ensure consistency in text properties across multiple notes and dimensions by using a single function call to apply the formatting of the selected parent text.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - The active document must be a drawing file.
> - The user must first select the parent text object (either a note or dimension) whose properties will be matched.
> - Subsequent selections must include the text objects to be modified.

## Results
> [!NOTE]
> - The selected text objects will have their properties (font, height) updated to match the parent text.
> - A confirmation message will be shown in the Immediate window for each updated text object.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

' Global variables for the SolidWorks application, selected text objects, and text object counters
Public swApp As SldWorks.SldWorks
Public vNoteObjects() As Object      ' Array to hold selected note objects
Public vdisDimObjects() As Object    ' Array to hold selected dimension objects
Dim iDisDim As Integer               ' Counter for dimension objects
Dim iNote As Integer                 ' Counter for note objects
Dim parentNote As Note               ' Parent note object for text matching
Dim parentDisDim As DisplayDimension ' Parent dimension object for text matching
Dim isNote As Boolean                ' Flag indicating if a note is selected
Dim isDisDim As Boolean              ' Flag indicating if a dimension is selected

' Constants for SolidWorks selection types
Const vSelNote As Integer = SwConst.swSelectType_e.swSelNOTES       ' Selection type for notes
Const vSelDims As Integer = SwConst.swSelectType_e.swSelDIMENSIONS  ' Selection type for dimensions

' Subroutine to get the selected text objects and populate arrays
Public Sub GetTheTextObjects()
    Dim swSelMgr As SelectionMgr    ' Selection manager object
    Dim swPart As ModelDoc2         ' Active document object
    Dim i As Integer                ' Number of selected objects
    Dim t As Integer                ' Loop counter

    ' Initialize objects
    Set swPart = swApp.ActiveDoc
    Set swSelMgr = swPart.SelectionManager
    i = swSelMgr.GetSelectedObjectCount2(-1)

    ' Check if no text objects are selected
    If i = 0 Then
        MsgBox "No text objects selected."
        Exit Sub
    End If

    On Error GoTo Errhandler
    ' Load the arrays with the selected text objects
    For t = 0 To i
        ' Check if selected object is a note
        If swSelMgr.GetSelectedObjectType3(t, -1) = vSelNote Then
            isNote = True
            If iNote = 0 Then
                ReDim vNoteObjects(0)
                iNote = 1
            Else
                ReDim Preserve vNoteObjects(UBound(vNoteObjects) + 1)
            End If
            Set vNoteObjects(UBound(vNoteObjects)) = swSelMgr.GetSelectedObject6(t, -1)
        End If
        ' Check if selected object is a dimension
        If swSelMgr.GetSelectedObjectType3(t, -1) = vSelDims Then
            isDisDim = True
            If iDisDim = 0 Then
                ReDim vdisDimObjects(0)
                iDisDim = 1
            Else
                ReDim Preserve vdisDimObjects(UBound(vdisDimObjects) + 1)
            End If
            Set vdisDimObjects(UBound(vdisDimObjects)) = swSelMgr.GetSelectedObject6(t, -1)
        End If
    Next
    On Error GoTo 0

    ' Call subroutine to apply text properties
    Call changeText
    Exit Sub

Errhandler:
    MsgBox "Error occurred while getting text objects: " & Err.Description
    Resume Next
End Sub

' Function to grab the parent text properties for matching
Public Function GrabParentText() As Boolean
    Dim swSelMgr As SelectionMgr    ' Selection manager object
    Dim swPart As ModelDoc2         ' Active document object
    Dim retVal As Boolean           ' Return value indicating if a parent text is found
    Dim i As Integer                ' Number of selected objects
    Dim t As Integer                ' Loop counter

    retVal = False  ' Initialize return value to False

    ' Initialize objects
    Set parentNote = Nothing
    Set parentDisDim = Nothing
    Set swPart = swApp.ActiveDoc
    Set swSelMgr = swPart.SelectionManager
    i = swSelMgr.GetSelectedObjectCount2(-1)

    ' Check if no objects are selected
    If i = 0 Then
        MsgBox "No objects selected."
        GrabParentText = retVal
        Exit Function
    End If

    On Error GoTo Errhandler
    ' Loop through selected objects to find a parent text object
    For t = 0 To i
        ' Check if selected object is a note
        If swSelMgr.GetSelectedObjectType3(t, -1) = vSelNote Then
            Set parentNote = swSelMgr.GetSelectedObject6(t, -1)
            retVal = True
            Exit For
        End If
        ' Check if selected object is a dimension
        If swSelMgr.GetSelectedObjectType3(t, -1) = vSelDims Then
            Set parentDisDim = swSelMgr.GetSelectedObject6(t, -1)
            retVal = True
            Exit For
        End If
    Next

Errhandler:
    ' Show message if no valid parent text object is found
    If parentNote Is Nothing And parentDisDim Is Nothing Then
        MsgBox "No text objects selected."
    End If
    GrabParentText = retVal
End Function

' Subroutine to apply parent text properties to the selected text objects
Sub changeText()
    Dim sdModel As ModelDoc2             ' Active document object
    Dim swSelMgr As SelectionMgr         ' Selection manager object
    Dim pFont As String                  ' Parent text font
    Dim pCharHt As Double                ' Parent text character height
    Dim pAnn As Annotation               ' Parent annotation object
    Dim pTextFor As TextFormat           ' Parent text format object
    Dim pIsUseDocFormat As Boolean       ' Flag for using document format
    Dim swAnn As Annotation              ' Annotation object for selected text
    Dim swTxtFormat As TextFormat        ' Text format object for selected text
    Dim swNote As Note                   ' Note object
    Dim swdisdim As DisplayDimension     ' Display dimension object
    Dim nAngle As Double                 ' Angle for the note

    ' Retrieve parent text properties
    If Not parentNote Is Nothing Then
        Set pAnn = parentNote.GetAnnotation
    ElseIf Not parentDisDim Is Nothing Then
        Set pAnn = parentDisDim.GetAnnotation
    End If

    ' Get text format properties from parent annotation
    pIsUseDocFormat = pAnn.GetUseDocTextFormat(0)
    Set pTextFor = pAnn.GetTextFormat(0)
    pFont = pTextFor.TypeFaceName
    pCharHt = pTextFor.CharHeight

    Set sdModel = swApp.ActiveDoc
    Set swSelMgr = sdModel.SelectionManager

    ' Apply parent properties to selected note objects
    If isNote = True Then
        On Error GoTo Errhandler
        For i = 0 To UBound(vNoteObjects)
            Set swNote = vNoteObjects(i)
            nAngle = swNote.Angle
            Set swAnn = swNote.GetAnnotation
            Set swTxtFormat = swAnn.GetTextFormat(0)
            swTxtFormat.CharHeight = pCharHt
            swTxtFormat.TypeFaceName = pFont
            swAnn.SetTextFormat 0, pIsUseDocFormat, swTxtFormat
            swNote.Angle = nAngle
        Next
    End If

    ' Apply parent properties to selected dimension objects
    If isDisDim = True Then
        On Error GoTo Errhandler1
        For i = 0 To UBound(vdisDimObjects)
            Set swdisdim = vdisDimObjects(i)
            Set swAnn = swdisdim.GetAnnotation
            Set swTxtFormat = swAnn.GetTextFormat(0)
            swTxtFormat.CharHeight = pCharHt
            swTxtFormat.TypeFaceName = pFont
            swAnn.SetTextFormat 0, pIsUseDocFormat, swTxtFormat
        Next
    End If

    ' Redraw graphics to apply changes
    sdModel.GraphicsRedraw2
    Exit Sub

Errhandler:
    MsgBox "Error occurred while applying note properties: " & Err.Description
    Resume Next

Errhandler1:
    MsgBox "Error occurred while applying dimension properties: " & Err.Description
    Resume Next
End Sub

' Main subroutine to initialize the macro
Sub main()
    Dim swPart As ModelDoc2  ' Active document object
    Set swApp = Application.SldWorks
    Set swPart = swApp.ActiveDoc

    ' Check if the active document is a drawing
    If swPart Is Nothing Then
        MsgBox "No Active File.", vbCritical, "Wrong File Type"
        End
    End If

    ' Check if the active document type is a drawing
    If swPart.GetType <> 3 Then
        MsgBox "Can only run in a drawing." & vbNewLine & "Active document must be a drawing file."
        End
    End If

    ' Initialize counters
    iNote = 0
    iDisDim = 0

    ' Rebuild the drawing to ensure proper updates
    swPart.ForceRebuild3 True

    ' Show form for user interaction (assumes a form named frmSelect exists)
    frmSelect.Show vbModeless
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).
