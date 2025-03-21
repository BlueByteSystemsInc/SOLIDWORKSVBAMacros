# Annotations Pro SolidWorks Macro

## Description
This macro automates the adjustment of annotations within SolidWorks drawings and models to match specified document settings.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later


## Pre-Conditions
- A SolidWorks document should be active before running this macro.
- Ensure all required libraries and constants (swconst.bas) are included.

## Results
- Sets annotations and dimensions to use the document font.
- Adjusts view titles, BOM balloon sizes, and arrowhead styles to match document settings.
- Ensures consistent sizing for revision triangles and weld symbol notes.

## Steps to Setup the Macro

1. **Create the UserForm**:
   - Open the VBA editor in SolidWorks with `Alt + F11`.
   - In Project Explorer, right-click the project and select **Insert > UserForm**.
   - Rename to `Form_AnnotationsPro`.
   - In Project Explorer, right-click the project and select **Insert > Module**.
   - Rename to `swConst`.
   - Design the form with a ListBox and two buttons (Process and Close).

2. **Add VBA Code**:
   - Insert the macro code into the main module.
   - Insert the UserForm code into the Form_AnnotationsPro code-behind.
   - Insert the swConst code into the swConst module.

3. **Save and Run the Macro**:
   - Save the macro (e.g., `AnnotationsPro.swp`).
   - Run by navigating to **Tools > Macro > Run** and selecting your macro.

4. **Using the Macro**:
   - The UserForm `Annotations Pro` will appear for interaction.
   - Select settings from the ListBox and press Process to apply.
   - Close the form when done.

## VBA Macro Code

```vbnet
'-----------------------------------------------------------------------------
' Annotations_Pro.swp
' This macro requires SolidWorks constants (swconst.bas) module to be inserted.
'
' This macro completes the folowing tasks when run on a model or drawing:
'   * Set annotations and dimensions to 'Use Document Font'.              (All)
'   * Set view titles to match appropriate text format and underline      (Dwg)
'   * Set BOM balloon size/fit and bent leader to match document setting  (Dwg)
'   * Set arrowhead for annotations to match document setting             (All)
'   * Set arrowhead for dimensions to match document setting              (All)
'   * Set annotations with triangle border to a consistent size (revision)(Dwg)
'   * Trims/Pads spaces in weld symbol notes for consistent weld symbols  (Dwg)
'
' Start this macro with 'Main' subroutine below
'------------------------------------------------------------------------------
' Predefine common objects and variables used in this module
'------------------------------------------------------------------------------
Global swApp As Object             ' SolidWorks application
Global Document As Object          ' Active document

'------------------------------------------------------------------------------
' Start this macro here
'------------------------------------------------------------------------------
Sub Main()
  Set swApp = CreateObject("SldWorks.Application")  ' Attach to SWX
  Set Document = swApp.ActiveDoc                    ' Grab active doc
  If Document Is Nothing Then                       ' Is doc loaded
    MsgBox "No model loaded."                       ' Nothing - Warn
  Else                                              ' Check doc loaded
    Form_AnnotationsPro.Show
  End If                                            ' Check doc loaded
  If Not Document Is Nothing Then                   ' If document is loaded
    Document.GraphicsRedraw2                        ' Redraw gfx window
  End If
End Sub
```

## VBA UserForm Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

Dim FileType As Integer         ' Active document type
Dim Annotation As Object        ' Current annotation
Dim Note As Object              ' Specific annotation
Dim TextFormat As Object        ' Annotation text format
Dim DisplayDim As Object        ' Dimension object
Dim View As Object              ' Current drawing view
Dim Feature As Object           ' Current model feature
Dim TriangleSize As Integer     ' Triangle size preset
Dim FirstView As Boolean        ' First drawing view indicator
Const Version = "1.21"

'------------------------------------------------------------------------------
' Set annotation text to correct text format, and set 'Use Document Font'.
'------------------------------------------------------------------------------
Private Sub FixAnnotationFont()
  ' Detect type of annotation, then reset all annotations to match current
  ' document based settings.
  Select Case Annotation.GetType              ' Detect type of annotation
      Case swCThread                          ' Cosmetic Thread
      Case swDatumTag                         ' Datum Feature
      Case swDatumTargetSym                   ' Datum Target
      Case swDisplayDimension                 ' Component Dimensions
           ' Updated via FixDisplayDims routine
      Case swGTol                             ' Geometric Tolerance
           ' Set text format to swDetailingNoteTextFormat
           If ListBoxAnnotations.Selected(5) = True Then
             Annotation.SetTextFormat 0, True, _
             Document.GetUserPreferenceTextFormat(swDetailingNoteTextFormat)
           End If
      Case swNote                             ' Note
           FixNoteText                        ' Fix text settings
           FixLeaders                         ' Fix leader settings
      Case swSFSymbol                         ' Surface Finish Symbol
           ' Set text format to swDetailingSurfaceFinishTextFormat
           If ListBoxAnnotations.Selected(8) = True Then
             Annotation.SetTextFormat 0, True, _
             Document.GetUserPreferenceTextFormat(swDetailingSurfaceFinishTextFormat)
             
           End If
      Case swWeldSymbol                       ' Weld Symbol
           ' Set text format to swDetailingWeldSymbolTextFormat
           If ListBoxAnnotations.Selected(6) = True Then
             Annotation.SetTextFormat 0, True, _
             Document.GetUserPreferenceTextFormat(swDetailingWeldSymbolTextFormat)
           End If
      Case swCustomSymbol                     ' Custom Symbol
      Case swDowelSym                         ' Dowel Symbol
      Case swLeader                           ' Leader Note
           FixNoteText                        ' Fix text settings
           FixLeaders                         ' Fix leader settings
      Case Else
  End Select
End Sub

'------------------------------------------------------------------------------
' Set view titles (SECTION/DETAIL/VIEW) to appropriate text format & underline
' Set annotations with triangle border to a consistent size (revision)
'------------------------------------------------------------------------------
Private Sub FixNoteText()
  Dim NoteVal As String
  Set Note = Annotation.GetSpecificAnnotation      ' Get note text
  ' Get contents of the note text
  NoteVal = Note.GetText
  ' Detect content of note text
  ' Update text format to current note text format
  If ListBoxAnnotations.Selected(0) = True Then
  Annotation.SetTextFormat 0, True, _
      Document.GetUserPreferenceTextFormat(swDetailingNoteTextFormat)
                                                 ' True = 'Use Doc Format'
  End If
  ' If annotation balloon style matches BOM balloon style set use of
  ' bent leader to match document based settings
  If ListBoxAnnotations.Selected(3) = True Then
    If Note.GetBalloonStyle = _
          Document.GetUserPreferenceIntegerValue(swDetailingBOMBalloonStyle) _
          Then
      ' Update balloon leader settings
      Annotation.SetLeader2 True, swLS_SMART, True, _
          Document.GetUserPreferenceToggle(swDetailingBalloonsDisplayWithBentLeader), _
          False, False
      ' Update balloon size
      Note.SetBalloon _
          Document.GetUserPreferenceIntegerValue(swDetailingBOMBalloonStyle), _
          Document.GetUserPreferenceIntegerValue(swDetailingBOMBalloonFit)
    End If
  End If
    ' If annotation balloon style is a triangle, treat it as a
    ' revision triangle and update its size.
  If Note.GetBalloonStyle = swBS_Triangle Then ' Triangle Balloon = Rev
    If ListBoxAnnotations.Selected(4) = True Then
      Annotation.SetLeader2 False, swLS_SMART, True, False, False, False
      Note.SetBalloon swBS_Triangle, TriangleSize    ' Update triangle size
    End If
  End If
  If ListBoxAnnotations.Selected(1) = True Then
    If Mid$(NoteVal, 1, 7) = "SECTION" Then
      ' If note text begins with "SECTION", then set to section font & underline
      Set TextFormat = _
          Document.GetUserPreferenceTextFormat(swDetailingSectionTextFormat)
      TextFormat.Underline = True                    ' Underline note
      Note.SetTextFormat False, TextFormat           ' False = use my settings
    ElseIf Mid$(NoteVal, 1, 6) = "DETAIL" Then
      ' If note text begins with "DETAIL", then set to detail font & underline
      Set TextFormat = _
          Document.GetUserPreferenceTextFormat(swDetailingDetailTextFormat)
      TextFormat.Underline = True                    ' Underline note
      Note.SetTextFormat False, TextFormat           ' False = use my settings
    ElseIf Mid$(NoteVal, 1, 4) = "VIEW" Then
      ' If note text begins with "VIEW", then set to view arrow font & underline
      Set TextFormat = _
          Document.GetUserPreferenceTextFormat(swDetailingViewArrowTextFormat)
      TextFormat.Underline = True                    ' Underline note
      Note.SetTextFormat False, TextFormat           ' False = use my settings
    ElseIf NoteVal = "NOTES:" Or NoteVal = "NOTE:" Then
      ' If note text equals "NOTES:" or "NOTE:", then set text to note font,
      ' underline, and increase text size
      Set TextFormat = _
          Document.GetUserPreferenceTextFormat(swDetailingNoteTextFormat)
      TextFormat.Underline = True                    ' Underline note
      TextFormat.CharHeight = TextFormat.CharHeight * 1.5 ' Set char height
      Note.SetTextFormat False, TextFormat           ' False = use my settings
    ElseIf NoteVal = "<MOD-CL>" Then
      ' If note equals "<MOD-CL>" (Centerline) then set text to note font
      ' and increase text size
      Set TextFormat = _
          Document.GetUserPreferenceTextFormat(swDetailingNoteTextFormat)
      TextFormat.Underline = True                    ' Remove underline
      TextFormat.CharHeight = TextFormat.CharHeight * 1.5 ' Set char height
      Note.SetTextFormat False, TextFormat           ' False = use my settings
    End If
  End If
End Sub

'------------------------------------------------------------------------------
' Sets arrowhead for annotations to match document setting
'------------------------------------------------------------------------------
Private Sub FixLeaders()
  If Annotation.GetLeader = True Then              ' If a leader exists
    For i = 0 To Annotation.GetLeaderCount         ' Get number of leaders
      Annotation.GetArrowHeadStyleAtIndex (i)      ' Fix arrow for each leader
      Annotation.SetLeader2 Annotation.GetLeader, _
                Annotation.GetLeaderSide, _
                True, Annotation.GetBentLeader, _
                Annotation.GetLeaderPerpendicular, _
                Annotation.GetLeaderAllAround      ' True = smartArrowHeadStyle
    Next i                                         ' Next leader
  End If                                           ' Leader displayed
End Sub

'------------------------------------------------------------------------------
' Sets arrowhead for dimensions to match document setting
'------------------------------------------------------------------------------
Private Sub FixDisplayDims(MyDim)
  If ListBoxAnnotations.Selected(2) = True Then
    ' Reset arrow selections
    MyDim.SetArrowHeadStyle False, _
      Document.GetUserPreferenceIntegerValue(swDetailingArrowStyleForDimensions)
                                                      ' False = use my setting
    ' Set arrow selections to use document settings.
    MyDim.SetArrowHeadStyle True, _
      Document.GetUserPreferenceIntegerValue(swDetailingArrowStyleForDimensions)
                                                      ' True = use doc setting
    ' Reset dimension font to use document font.
    Set Annotation = DisplayDim.GetAnnotation         ' Get next annotation
    Annotation.SetTextFormat 0, True, _
      Document.GetUserPreferenceTextFormat(swDetailingDimensionTextFormat)
    ' Reset tolerance font to use document font.
    Set DimObj = DisplayDim.GetDimension              ' Get dimension
    FontInfo = DimObj.GetToleranceFontInfo            ' Get tolerance font info
    DimHt = Document.GetUserPreferenceDoubleValue(swDetailingToleranceTextHeight)
    DimObj.SetToleranceFontInfo True, 1, DimHt       ' Use doc setting
  End If
End Sub

'------------------------------------------------------------------------------
' Trims/Pads spaces in weld symbol notes
'------------------------------------------------------------------------------
Private Sub FixWeldSym(MyDoc)
  Dim NewText(9) As String
  ' Trim spaces from each weld annotation
  For z = 1 To 9
    text = MyDoc.GetText(z)         ' Get individual text
    text = RTrim(LTrim(text))       ' Trim left and right spaces
    If text = " " Then text = ""    ' Trim single space
    NewText(z) = UCase(text)        ' Force annotation to upper case
  Next z
  ' Pad spaces for weld length above weld line
  If NewText(1) <> "" And Len(NewText(3)) < 4 Then
    NewText(3) = Left$(NewText(3) + "    ", 4)
  End If
  ' Pad spaces for weld length below weld line
  If NewText(5) <> "" And Len(NewText(7)) < 4 Then
    NewText(7) = Left$(NewText(7) + "    ", 4)
  End If
  ' Update annotations above weld line
  MyDoc.SetText True, NewText(1), NewText(2), _
      NewText(3), NewText(4), MyDoc.GetContour(True)
  ' Update annotations below weld line
  MyDoc.SetText False, NewText(5), NewText(6), _
      NewText(7), NewText(8), MyDoc.GetContour(False)
  ' Update weld process annotations. Use existing settings
  MyDoc.SetProcess MyDoc.GetProcess, _
      NewText(9), MyDoc.GetProcessReference
End Sub

'------------------------------------------------------------------------------
' Traverse each sheet of the drawing, then each drawing view and update/repair
' all annotations and dimensions
'------------------------------------------------------------------------------
Sub RepairDrawing()
  SheetNames = Document.GetSheetNames               ' Get names of sheets
  For i = 0 To Document.GetSheetCount - 1           ' For each sheet
    Document.ActivateSheet (SheetNames(i))          ' Activate sheet
    Set View = Document.GetFirstView                ' Get first view
    FirstView = True
    While Not View Is Nothing                       ' View is valid
      ' Do the following only if current view is not FirstView (Sheet Format)
      If FirstView = False Then                     ' If not FirstView
        Set Annotation = View.GetFirstAnnotation2   ' Get first annotation
        While Not Annotation Is Nothing             ' Annotation is valid
          Set Note = Annotation.GetSpecificAnnotation ' Get note text
          FixAnnotationFont                         ' Set correct text format
          Set Annotation = Annotation.GetNext2      ' Get next annotation
        Wend                                        ' Repeat each annotation
      End If                                        ' FirstView check
      FirstView = False
      ' Traverse thru all DisplayDim's in drawing view
      Set DisplayDim = View.GetFirstDisplayDimension3 ' Get first dimension
      While Not DisplayDim Is Nothing               ' dimension valid
        FixDisplayDims DisplayDim                   ' Fix dimension settings
        Set DisplayDim = DisplayDim.GetNext3        ' Get next dimension
      Wend                                          ' Repeat for all dimension
      ' Traverse thru all WeldSymbols in drawing view
      If ListBoxAnnotations.Selected(7) = True Then
        Set WeldSymbol = View.GetFirstWeldSymbol      ' Get first weld symb
        While Not WeldSymbol Is Nothing               ' Weld symb valid
          FixWeldSym WeldSymbol                       ' Fix weld symb notes
          Set WeldSymbol = WeldSymbol.GetNext         ' Next weld symb
        Wend                                          ' Repeat for all weld symbs
      End If
      Set View = View.GetNextView                   ' Get next view
    Wend                                            ' Repeat for all views
  Next i                                            ' Next sheet
  If i > 0 Then                                     ' More than 1 sheet
    Document.ActivateSheet (SheetNames(0))          ' Back to 1st sheet
  End If
End Sub

'------------------------------------------------------------------------------
' Traverse each annotation and update/repair text format
' Traverse each model feature and update/repair all dimensions
'------------------------------------------------------------------------------
Sub RepairModel()
  Set Annotation = Document.GetFirstAnnotation2     ' Get first annot
  While Not Annotation Is Nothing                   ' Annotation is valid
    Set Note = Annotation.GetSpecificAnnotation     ' Get note text
    FixAnnotationFont                               ' Fix annotation font
    Set Annotation = Annotation.GetNext2            ' Get next annotation
  Wend                                              ' Repeat for each annotation
  ' Update dimensions in a SolidWorks model by traversing thru each feature
  ' in the model, then traversing thru each dimension in the feature.
  Set Feature = Document.FirstFeature               ' Get the 1st feature
  While Not Feature Is Nothing                      ' Feature is valid
    ' Traverse thru all DisplayDim's in model feature
    Set DisplayDim = Feature.GetFirstDisplayDimension ' Get first dimension
    While Not DisplayDim Is Nothing                 ' dimension valid
      FixDisplayDims DisplayDim                     ' Fix dimension settings
      Set DisplayDim = _
        Feature.GetNextDisplayDimension(DisplayDim) ' Get next dimension
    Wend                                            ' Repeat for all dimension
    Set SubFeat = Feature.GetFirstSubFeature        ' Get first feature
    ' Traverse thru all sub features in model feature
    While Not SubFeat Is Nothing                    ' Feature is valid
      ' Traverse thru all DisplayDim's in model sub feature
      Set DisplayDim = SubFeat.GetFirstDisplayDimension ' Get first dim
      While Not DisplayDim Is Nothing               ' dimension valid
        FixDisplayDims DisplayDim                   ' Fix dimension settings
        Set DisplayDim = _
          Feature.GetNextDisplayDimension(DisplayDim) ' Get next dimension
      Wend                                          ' Repeat for all dimension
      Set SubFeat = SubFeat.GetNextSubFeature       ' Get next feature
    Wend
    Set Feature = Feature.GetNextFeature()          ' Get next feature
  Wend                                              ' Repeat for all features
End Sub

Private Sub CommandClose_Click()
  End
End Sub

Private Sub CommandProcess_Click()
  ' Settings for revision triangle size
  BalloonSize = Document.GetUserPreferenceIntegerValue(swDetailingBOMBalloonFit)
  If BalloonSize > 0 Then
    TriangleSize = BalloonSize - 1                ' 1 less than BOM balloon
  Else
    TriangleSize = BalloonSize                    ' Size equals BOM balloon
  End If
  FirstView = False                               ' Clear FirstView indicator
  FileType = Document.GetType                     ' Get doc type
  If FileType = swDocDRAWING Then                 ' Is doc type drawing
    RepairDrawing
  ElseIf FileType = swDocPART Or _
         FileType = swDocASSEMBLY Then            ' Is doc type model
    RepairModel
  End If                                          ' Check doc type
  For i = 0 To ListBoxAnnotations.ListCount - 1
    ListBoxAnnotations.Selected(i) = False
  Next i
  Document.WindowRedraw                           ' Redraw screen
End Sub

Private Sub UserForm_Initialize()
  Form_AnnotationsPro.Caption = Form_AnnotationsPro.Caption + " " + Version
  ListBoxAnnotations.Clear
  ListBoxAnnotations.AddItem "Correct note fonts."                  ' 0
  ListBoxAnnotations.AddItem "Highlight view and note titles."      ' 1
  ListBoxAnnotations.AddItem "Correct Dimension fonts."             ' 2
  ListBoxAnnotations.AddItem "Correct BOM ballons sizes."           ' 3
  ListBoxAnnotations.AddItem "Correct Revision triangle size."      ' 4
  ListBoxAnnotations.AddItem "Correct Geometric Tolerance fonts."   ' 5
  ListBoxAnnotations.AddItem "Correct weld symbol note fonts."      ' 6
  ListBoxAnnotations.AddItem "Correct weld symbol note padding."    ' 7
  ListBoxAnnotations.AddItem "Correct surface finish symbol fonts." ' 8
End Sub
```
## swConst Code

```vbnet

' Defines

Public Const swTnChamfer As String = "Chamfer"
Public Const swTnFillet As String = "Fillet"
Public Const swTnCavity As String = "Cavity"
Public Const swTnDraft As String = "Draft"
Public Const swTnMirrorSolid As String = "MirrorSolid"
Public Const swTnCirPattern As String = "CirPattern"
Public Const swTnLPattern As String = "LPattern"
Public Const swTnMirrorPattern As String = "MirrorPattern"
Public Const swTnShell As String = "Shell"
Public Const swTnBlend As String = "Blend"
Public Const swTnBlendCut As String = "BlendCut"
Public Const swTnExtrusion As String = "Extrusion"
Public Const swTnBoss As String = "Boss"
Public Const swTnCut As String = "Cut"
Public Const swTnRefCurve As String = "RefCurve"
Public Const swTnRevolution As String = "Revolution"
Public Const swTnRevCut As String = "RevCut"
Public Const swTnSweep As String = "Sweep"
Public Const swTnSweepCut As String = "SweepCut"
Public Const swTnStock As String = "Stock"
Public Const swTnSurfCut As String = "SurfCut"
Public Const swTnThicken As String = "Thicken"
Public Const swTnThickenCut As String = "ThickenCut"
Public Const swTnVarFillet As String = "VarFillet"
Public Const swTnSketchHole As String = "SketchHole"
Public Const swTnHoleWzd As String = "HoleWzd"
Public Const swTnImported As String = "Imported"
Public Const swTnBaseBody As String = "BaseBody"
Public Const swTnDerivedLPattern As String = "DerivedLPattern"
Public Const swTnCosmeticThread As String = "CosmeticThread"
Public Const swTnSheetMetal As String = "SheetMetal"
Public Const swTnFlattenBends As String = "FlattenBends"
Public Const swTnProcessBends As String = "ProcessBends"
Public Const swTnOneBend As String = "OneBend"
Public Const swTnBaseFlange As String = "SMBaseFlange"
Public Const swTnSketchBend As String = "SketchBend"
Public Const swTnSM3dBend As String = "SM3dBend"
Public Const swTnEdgeFlange As String = "EdgeFlange"
Public Const swTnFlatPattern As String = "FlatPattern"
Public Const swTnCenterMark As String = "CenterMark"
Public Const swTnDrSheet As String = "DrSheet"
Public Const swTnAbsoluteView As String = "AbsoluteView"
Public Const swTnDetailView As String = "DetailView"
Public Const swTnRelativeView As String = "RelativeView"
Public Const swTnSectionPartView As String = "SectionPartView"
Public Const swTnSectionAssemView As String = "SectionAssemView"
Public Const swTnUnfoldedView As String = "UnfoldedView"
Public Const swTnAuxiliaryView As String = "AuxiliaryView"
Public Const swTnDetailCircle As String = "DetailCircle"
Public Const swTnDrSectionLine As String = "DrSectionLine"
Public Const swTnMateCoincident As String = "MateCoincident"
Public Const swTnMateConcentric As String = "MateConcentric"
Public Const swTnMateDistanceDim As String = "MateDistanceDim"
Public Const swTnMateParallel As String = "MateParallel"
Public Const swTnMateTangent As String = "MateTangent"
Public Const swTnReference As String = "Reference"
Public Const swTnRefPlane As String = "RefPlane"
Public Const swTnRefAxis As String = "RefAxis"
Public Const swTnReferenceCurve As String = "ReferenceCurve"
Public Const swTnRefSurface As String = "RefSurface"
Public Const swTnCoordinateSystem As String = "CoordSys"
Public Const swTnAttribute As String = "Attribute"
Public Const swTnProfileFeature As String = "ProfileFeature"
Public Const SYMBOL_MARKER_START As String = "<"
Public Const SYMBOL_MARKER_END As String = ">"
Public Const SYMBOL_MARKER_SPACE As String = "-"
Public Const PLANE_TYPE As Integer = 4001
Public Const CYLINDER_TYPE As Integer = 4002
Public Const CONE_TYPE As Integer = 4003
Public Const SPHERE_TYPE As Integer = 4004
Public Const TORUS_TYPE As Integer = 4005
Public Const BSURF_TYPE As Integer = 4006
Public Const BLEND_TYPE As Integer = 4007
Public Const OFFSET_TYPE As Integer = 4008
Public Const EXTRU_TYPE As Integer = 4009
Public Const SREV_TYPE As Integer = 4010
Public Const LINE_TYPE As Integer = 3001
Public Const CIRCLE_TYPE As Integer = 3002
Public Const ELLIPSE_TYPE As Integer = 3003
Public Const INTERSECTION_TYPE As Integer = 3004
Public Const BCURVE_TYPE As Integer = 3005
Public Const SPCURVE_TYPE As Integer = 3006
Public Const CONSTPARAM_TYPE As Integer = 3008
Public Const TRIMMED_TYPE As Integer = 3009
Public Const TIME_ORIGIN As String = "1990, 1, 1, 0, 0, 0"
Public Const SWBODYINTERSECT As Integer = 15901
Public Const SWBODYCUT As Integer = 15902
Public Const SWBODYADD As Integer = 15903
Public Const NUM_HOLE_GENERIC_TYPES As Integer = 6
Public Const NUM_HOLE_TYPES As Integer = 57
Public Const NUM_HOLE_STANDARD_TYPES As Integer = 249


' Enumerations

' ----
'  Document Types
' ----
Public Enum swDocumentTypes_e
        swDocNONE = 0   '  Used to be TYPE_NONE
        swDocPART = 1   '  Used to be TYPE_PART
        swDocASSEMBLY = 2       '  Used to be TYPE_ASSEMBLY
        swDocDRAWING = 3        '  Used to be TYPE_DRAWING
        swDocSDM = 4    '  Solid data manager.
End Enum

' ----
'  Selection Types
' ----
'  The following are the possible type ids returned by the function
'      ISelectionMgr::GetSelectedObjectType.
'  The string names to the right of the type id definition is the "type name"
'      used by the methods:  IModelDoc::SelectByID && AndSelectByID
Public Enum swSelectType_e
        swSelNOTHING = 0
        swSelEDGES = 1  '  "EDGE"
        swSelFACES = 2  '  "FACE"
        swSelVERTICES = 3       '  "VERTEX"
        swSelDATUMPLANES = 4    '  "PLANE"
        swSelDATUMAXES = 5      '  "AXIS"
        swSelDATUMPOINTS = 6    '  "DATUMPOINT"
        swSelOLEITEMS = 7       '  "OLEITEM"
        swSelATTRIBUTES = 8     '  "ATTRIBUTE"
        swSelSKETCHES = 9       '  "SKETCH"
        swSelSKETCHSEGS = 10    '  "SKETCHSEGMENT"
        swSelSKETCHPOINTS = 11  '  "SKETCHPOINT"
        swSelDRAWINGVIEWS = 12  '  "DRAWINGVIEW"
        swSelGTOLS = 13 '  "GTOL"
        swSelDIMENSIONS = 14    '  "DIMENSION"
        swSelNOTES = 15 '  "NOTE"
        swSelSECTIONLINES = 16  '  "SECTIONLINE"
        swSelDETAILCIRCLES = 17 '  "DETAILCIRCLE"
        swSelSECTIONTEXT = 18   '  "SECTIONTEXT"
        swSelSHEETS = 19        '  "SHEET"
        swSelCOMPONENTS = 20    '  "COMPONENT"
        swSelMATES = 21 '  "MATE"
        swSelBODYFEATURES = 22  '  "BODYFEATURE"
        swSelREFCURVES = 23     '  "REFCURVE"
        swSelEXTSKETCHSEGS = 24 '  "EXTSKETCHSEGMENT"
        swSelEXTSKETCHPOINTS = 25       '  "EXTSKETCHPOINT"
        swSelHELIX = 26 '  "HELIX" (is this wrong?)
        swSelREFERENCECURVES = 26       '  "REFERENCECURVES"
        swSelREFSURFACES = 27   '  "REFSURFACE"
        swSelCENTERMARKS = 28   '  "CENTERMARKS"
        swSelINCONTEXTFEAT = 29 '  "INCONTEXTFEAT"
        swSelMATEGROUP = 30     '  "MATEGROUP"
        swSelBREAKLINES = 31    '  "BREAKLINE"
        swSelINCONTEXTFEATS = 32        '  "INCONTEXTFEATS"
        swSelMATEGROUPS = 33    '  "MATEGROUPS"
        swSelSKETCHTEXT = 34    '  "SKETCHTEXT"
        swSelSFSYMBOLS = 35     '  "SFSYMBOL"
        swSelDATUMTAGS = 36     '  "DATUMTAG"
        swSelCOMPPATTERN = 37   '  "COMPPATTERN"
        swSelWELDS = 38 '  "WELD"
        swSelCTHREADS = 39      '  "CTHREAD"
        swSelDTMTARGS = 40      '  "DTMTARG"
        swSelPOINTREFS = 41     '  "POINTREF"
        swSelDCABINETS = 42     '  "DCABINET"
        swSelEXPLVIEWS = 43     '  "EXPLODEDVIEWS"
        swSelEXPLSTEPS = 44     '  "EXPLODESTEPS"
        swSelEXPLLINES = 45     '  "EXPLODELINES"
        swSelSILHOUETTES = 46   '  "SILHOUETTE"
        swSelCONFIGURATIONS = 47        '  "CONFIGURATIONS"
        swSelOBJHANDLES = 48
        swSelARROWS = 49        '  "VIEWARROW"
        swSelZONES = 50 '  "ZONES"
        swSelREFEDGES = 51      '  "REFERENCE-EDGE"
        swSelREFFACES = 52
        swSelREFSILHOUETTE = 53
        swSelBOMS = 54  '  "BOM"
        swSelEQNFOLDER = 55     '  "EQNFOLDER"
        swSelSKETCHHATCH = 56   '  "SKETCHHATCH"
        swSelIMPORTFOLDER = 57  '  "IMPORTFOLDER"
        swSelVIEWERHYPERLINK = 58       '  "HYPERLINK"
        swSelMIDPOINTS = 59
        swSelCUSTOMSYMBOLS = 60 '  "CUSTOMSYMBOL"
        swSelCOORDSYS = 61      '  "COORDSYS"
        swSelDATUMLINES = 62    '  "REFLINE"
        swSelROUTECURVES = 63
        swSelBOMTEMPS = 64      '  "BOMTEMP"
        swSelROUTEPOINTS = 65   '  "ROUTEPOINT"
        swSelCONNECTIONPOINTS = 66      '  "CONNECTIONPOINT"
        swSelROUTESWEEPS = 67
        swSelPOSGROUP = 68      '  "POSGROUP"
        swSelBROWSERITEM = 69   '  "BROWSERITEM"
        swSelFABRICATEDROUTE = 70       '  "ROUTEFABRICATED"
        swSelSKETCHPOINTFEAT = 71       '  "SKETCHPOINTFEAT"
        swSelEMPTYSPACE = 72    '  (is this wrong?)
        swSelCOMPSDONTOVERRIDE = 72
        swSelLIGHTS = 73        '  "LIGHTS"
        swSelWIREBODIES = 74
        swSelSURFACEBODIES = 75 '  "SURFACEBODY"
        swSelSOLIDBODIES = 76   '  "SOLIDBODY"
        swSelFRAMEPOINT = 77    '  "FRAMEPOINT"
        swSelSURFBODIESFIRST = 78
        swSelMANIPULATORS = 79  '  "MANIPULATOR"
        swSelPICTUREBODIES = 80 '  "PICTURE BODY"
        swSelSOLIDBODIESFIRST = 81
        swSelDOWELSYMS = 86     '  "DOWELSYM"
        swSelEXTSKETCHTEXT = 88 '  "EXTSKETCHTEXT"
        swSelBLOCKINST = 93     '  "BLOCKINST"
        swSelSKETCHREGION = 95  '  "SKETCHREGION"
        swSelSKETCHCONTOUR = 96 '  "SKETCHCONTOUR"
        swSelBLOCKDEF = 99      '  "BLOCKDEF"
        swSelCENTERMARKSYMS = 100       '  "CENTERMARKSYMS"
        swSelCENTERLINES = 103  '  "CENTERLINE"
'       swSelEVERYTHING = 4294967293
'       swSelLOCATIONS = 4294967294
'       swSelUNSUPPORTED = 4294967295
End Enum

' ----
'  Events Notifications
' ----
Public Enum swViewNotify_e      '  For IModelView ( DIID_DSldWorksEvents
        swViewRepaintNotify = 1
        swViewChangeNotify = 2
        swViewDestroyNotify = 3
        swViewRepaintPostNotify = 4
        swViewBufferSwapNotify = 5
        swViewDestroyNotify2 = 6
End Enum

Public Enum swFMViewNotify_e    '  For IFeatMgrView ( DIID_DSldWorksEvents
        swFMViewActivateNotify = 1
        swFMViewDeactivateNotify = 2
        swFMViewDestroyNotify = 3
End Enum

Public Enum swPartNotify_e      '  For IPartDoc ( DIID_DPartDocEvents )
        swPartRegenNotify = 1
        swPartDestroyNotify = 2
        swPartRegenPostNotify = 3
        swPartViewNewNotify = 4
        swPartNewSelectionNotify = 5
        swPartFileSaveNotify = 6
        swPartFileSaveAsNotify = 7
        swPartLoadFromStorageNotify = 8
        swPartSaveToStorageNotify = 9
        swPartConfigChangeNotify = 10
        swPartConfigChangePostNotify = 11
        swPartAutoSaveNotify = 12
        swPartAutoSaveToStorageNotify = 13
        swPartViewNewNotify2 = 14
        swPartLightingDialogCreateNotify = 15
        swPartAddItemNotify = 16
        swPartRenameItemNotify = 17
        swPartDeleteItemNotify = 18
        swPartModifyNotify = 19
        swPartFileReloadNotify = 20
        swPartAddCustomPropertyNotify = 21
        swPartChangeCustomPropertyNotify = 22
        swPartDeleteCustomPropertyNotify = 23
        swPartFeatureEditPreNotify = 24
        swPartFeatureSketchEditPreNotify = 25
        swPartFileSaveAsNotify2 = 26
        swPartDeleteSelectionPreNotify = 27
        swPartFileReloadPreNotify = 28
        swPartBodyVisibleChangeNotify = 29
        swPartRegenPostNotify2 = 30
        swPartFileSavePostNotify = 31
        swPartLoadFromStorageStoreNotify = 32
        swPartSaveToStorageStoreNotify = 33
End Enum

Public Enum swDrawingNotify_e   '  For IDrawingDoc ( DIID_DDrawingDocEvents )
        swDrawingRegenNotify = 1
        swDrawingDestroyNotify = 2
        swDrawingRegenPostNotify = 3
        swDrawingViewNewNotify = 4
        swDrawingNewSelectionNotify = 5
        swDrawingFileSaveNotify = 6
        swDrawingFileSaveAsNotify = 7
        swDrawingLoadFromStorageNotify = 8
        swDrawingSaveToStorageNotify = 9
        swDrawingAutoSaveNotify = 10
        swDrawingAutoSaveToStorageNotify = 11
        swDrawingConfigChangeNotify = 12
        swDrawingConfigChangePostNotify = 13
        swDrawingViewNewNotify2 = 14
        swDrawingAddItemNotify = 15
        swDrawingRenameItemNotify = 16
        swDrawingDeleteItemNotify = 17
        swDrawingModifyNotify = 18
        swDrawingFileReloadNotify = 19
        swDrawingAddCustomPropertyNotify = 20
        swDrawingChangeCustomPropertyNotify = 21
        swDrawingDeleteCustomPropertyNotify = 22
        swDrawingFileSaveAsNotify2 = 23
        swDrawingDeleteSelectionPreNotify = 24
        swDrawingFileReloadPreNotify = 25
        swDrawingFileSavePostNotify = 26
        swDrawingLoadFromStorageStoreNotify = 27
        swDrawingSaveToStorageStoreNotify = 28
End Enum

Public Enum swAssemblyNotify_e  '  For IAssemblyDoc ( DIID_DAssemblyDocEvents )
        swAssemblyRegenNotify = 1
        swAssemblyDestroyNotify = 2
        swAssemblyRegenPostNotify = 3
        swAssemblyViewNewNotify = 4
        swAssemblyNewSelectionNotify = 5
        swAssemblyFileSaveNotify = 6
        swAssemblyFileSaveAsNotify = 7
        swAssemblyLoadFromStorageNotify = 8
        swAssemblySaveToStorageNotify = 9
        swAssemblyConfigChangeNotify = 10
        swAssemblyConfigChangePostNotify = 11
        swAssemblyAutoSaveNotify = 12
        swAssemblyAutoSaveToStorageNotify = 13
        swAssemblyBeginInContextEditNotify = 14
        swAssemblyEndInContextEditNotify = 15
        swAssemblyViewNewNotify2 = 16
        swAssemblyLightingDialogCreateNotify = 17
        swAssemblyAddItemNotify = 18
        swAssemblyRenameItemNotify = 19
        swAssemblyDeleteItemNotify = 20
        swAssemblyModifyNotify = 21
        swAssemblyComponentStateChangeNotify = 22
        swAssemblyFileDropNotify = 23
        swAssemblyFileReloadNotify = 24
        swAssemblyComponentStateChangeNotify2 = 25
        swAssemblyAddCustomPropertyNotify = 26
        swAssemblyChangeCustomPropertyNotify = 27
        swAssemblyDeleteCustomPropertyNotify = 28
        swAssemblyFeatureEditPreNotify = 29
        swAssemblyFeatureSketchEditPreNotify = 30
        swAssemblyFileSaveAsNotify2 = 31
        swAssemblyInterferenceNotify = 32
        swAssemblyDeleteSelectionPreNotify = 33
        swAssemblyFileReloadPreNotify = 34
        swAssemblyComponentMoveNotify = 35
        swAssemblyComponentVisibleChangeNotify = 36
        swAssemblyBodyVisibleChangeNotify = 37
        swAssemblyFileDropPreNotify = 38
        swAssemblyFileSavePostNotify = 39
        swAssemblyLoadFromStorageStoreNotify = 40
        swAssemblySaveToStorageStoreNotify = 41
End Enum

Public Enum swAppNotify_e       '  For ISldWorks ( DIID_DSldWorksEvents )
        swAppFileOpenNotify = 1
        swAppFileNewNotify = 2
        swAppDestroyNotify = 3
        swAppActiveDocChangeNotify = 4
        swAppActiveModelDocChangeNotify = 5
        swAppPropertySheetCreateNotify = 6
        swAppNonNativeFileOpenNotify = 7
        swAppLightSheetCreateNotify = 8
        swAppDocumentConversionNotify = 9
        swAppLightweightComponentOpenNotify = 10
        swAppDocumentLoadNotify = 11
        swAppFileNewNotify2 = 12
        swAppFileOpenNotify2 = 13
        swAppReferenceNotFoundNotify = 14
        swAppPromptForFilenameNotify = 15
        swAppBeginTranslationNotify = 16
        swAppEndTranslationNotify = 17
End Enum

Public Enum swPropertySheetNotify_e
        swPropertySheetDestroyNotify = 1
        swPropertySheetHelpNotify = 2
End Enum

' ----
'  Parameter Types
' ----
Public Enum swParamType_e       '  For use with IAttributeDef::AddParameter (for example)
        swParamTypeDouble = 0
        swParamTypeString = 1
        swParamTypeInteger = 2
        swParamTypeDVector = 3
End Enum

' ----
'  The following is for angular dimension info returned GetDimensionInfo()
' ----
Public Enum swQuadant_e
        swQuadUnknown = 0
        swQuadPosQ1 = 1
        swQuadNegQ1 = 2
        swQuadPosQ2 = 3
        swQuadNegQ2 = 4
End Enum

' ----
'  The following enum is for interpreting ellipse data
' ----
Public Enum swEllipsePts_e
        swEllipseStartPt = 0
        swEllipseEndPt = 1
        swEllipseCenterPt = 2
        swEllipseMajorPt = 3
        swEllipseMinorPt = 4
End Enum

Public Enum swParabolaPts_e
        swParabolaStartPt = 0
        swParabolaEndPt = 1
        swParabolaFocusPt = 2
        swParabolaApexPt = 3
End Enum

' ----
'  The following define gtol symbol indices
' ----
Public Enum swGtolMatCondition_e
        swMcNONE = 0
        swMcMMC = 1
        swMcRFS = 2
        swMcLMC = 3
        swMsNONE = 4
        swMsPROJTOLZONE = 5
        swMsDIA = 6
        swMsSPHDIA = 7
        swMsRAD = 8
        swMsSPHRAD = 9
        swMsREF = 10
        swMsARCLEN = 11
End Enum

Public Enum swGtolGeomCharSymbol_e
        swGcsNONE = 12
        swGcsSYMMETRY = 13
        swGcsSTRAIGHT = 14
        swGcsFLAT = 15
        swGcsROUND = 16
        swGcsCYL = 17
        swGcsLINEPROF = 18
        swGcsSURFPROF = 19
        swGcsANG = 20
        swGcsPERP = 21
        swGcsPARALLEL = 22
        swGcsPOSITION = 23
        swGcsCONC = 24
        swGcsCIRCRUNOUT = 25
        swGcsTOTALRUNOUT = 26
        swGcsCIRCOPENRUNOUT = 27
        swGcsTOTALOPENRUNOUT = 28
End Enum

Public Enum swMateType_e
        swMateCOINCIDENT = 0
        swMateCONCENTRIC = 1
        swMatePERPENDICULAR = 2
        swMatePARALLEL = 3
        swMateTANGENT = 4
        swMateDISTANCE = 5
        swMateANGLE = 6
        swMateUNKNOWN = 7
        swMateSYMMETRIC = 8
        swMateCAMFOLLOWER = 9
End Enum

'  Enumerations for Detail View Creation
Public Enum swDetCircleShowType_e
        swDetCirclePROFILE = 0
        swDetCircleCIRCLE = 1
        swDetCircleDONTSHOW = 2
End Enum

'  Enumerations for Detail View Style
Public Enum swDetViewStyle_e
        swDetViewSTANDARD = 0
        swDetViewBROKEN = 1
        swDetViewLEADER = 2
        swDetViewNOLEADER = 3
        swDetViewCONNECTED = 4
End Enum

'  This enum has been changed to correct improper mate alignment mapping
Public Enum swMateAlign_e
        swMateAlignALIGNED = 0
        swMateAlignANTI_ALIGNED = 1
        swMateAlignCLOSEST = 2
        swAlignNONE = 0
        swAlignSAME = 1
        swAlignAGAINST = 2
End Enum

Public Enum swDisplayMode_e
        swWIREFRAME = 0
        swHIDDEN_GREYED = 1
        swHIDDEN = 2
        swSHADED = 3
        swFACETED_WIREFRAME = 4
        swFACETED_HIDDEN_GREYED = 5
        swFACETED_HIDDEN = 6
End Enum

Public Enum swArrowStyle_e
        swOPEN_ARROWHEAD = 0
        swCLOSED_ARROWHEAD = 1
        swSLASH_ARROWHEAD = 2
        swDOT_ARROWHEAD = 3
        swORIGIN_ARROWHEAD = 4
        swWIDE_ARROWHEAD = 5
        swISOWIDE_ARROWHEAD = 6
        swRUS_ARROWHEAD = 7
        swCLOSETOP_ARROWHEAD = 8
        swCLOSEBOT_ARROWHEAD = 9
        swNO_ARROWHEAD = 10
End Enum

Public Enum swLeaderSide_e
        swLS_SMART = 0
        swLS_LEFT = 1
        swLS_RIGHT = 2
End Enum

' ----
'  The following define Surface Finish Symbol types and options
'  Used by InsertSurfaceFinishSymbol, ModifySurfaceFinishSymbol
' ----
Public Enum swSFSymType_e
        swSFBasic = 0
        swSFMachining_Req = 1
        swSFDont_Machine = 2
        swSFJIS_Surface_Texture_1 = 3   '  Add next 5 JIS types, 08/26/99
        swSFJIS_Surface_Texture_2 = 4
        swSFJIS_Surface_Texture_3 = 5
        swSFJIS_Surface_Texture_4 = 6
        swSFJIS_No_Machining = 7
End Enum

Public Enum swSFLaySym_e
        swSFNone = 0
        swSFCircular = 1
        swSFCross = 2
        swSFMultiDir = 3
        swSFParallel = 4
        swSFPerp = 5
        swSFRadial = 6
        swSFParticulate = 7
End Enum

'  The different possibilities for types of texts in a Surface Finish symbol. (SFSymbol::Get/SetText)
Public Enum swSurfaceFinishSymbolText_e
        swSFSymbolMaterialRemovalAllowance = 1
        swSFSymbolProductionMethod = 2
        swSFSymbolSamplingLength = 3
        swSFSymbolOtherRoughnessValue = 4
        swSFSymbolMaximumRoughness = 5
        swSFSymbolMinimumRoughness = 6
        swSFSymbolRoughnessSpacing = 7
End Enum

Public Enum swLeaderStyle_e
        swNO_LEADER = 0
        swSTRAIGHT = 1
        swBENT = 2
End Enum

' ----
'  Balloon Information.  swBS_SplitCirc is not valid for Notes only for Balloons
' ----
Public Enum swBalloonStyle_e
        swBS_None = 0
        swBS_Circular = 1
        swBS_Triangle = 2
        swBS_Hexagon = 3
        swBS_Box = 4
        swBS_Diamond = 5
        swBS_SplitCirc = 6
        swBS_Pentagon = 7
        swBS_FlagPentagon = 8
        swBS_FlagTriangle = 9
        swBS_Underline = 10
End Enum

Public Enum swBalloonFit_e
        swBF_Tightest = 0
        swBF_1Char = 1
        swBF_2Chars = 2
        swBF_3Chars = 3
        swBF_4Chars = 4
        swBF_5Chars = 5
End Enum

'  Possible values for the Balloon upper and lower text content.
Public Enum swBalloonTextContent_e
        swBalloonTextCustom = 0
        swBalloonTextItemNumber = 1
        swBalloonTextQuantity = 2
End Enum

' ----
'  The following define length and angle unit types
' ----
Public Enum swLengthUnit_e
        swMM = 0
        swCM = 1
        swMETER = 2
        swINCHES = 3
        swFEET = 4
        swFEETINCHES = 5
        swANGSTROM = 6
        swNANOMETER = 7
        swMICRON = 8
        swMIL = 9
        swUIN = 10
End Enum

Public Enum swAngleUnit_e
        swDEGREES = 0
        swDEG_MIN = 1
        swDEG_MIN_SEC = 2
        swRADIANS = 3
End Enum

Public Enum swFractionDisplay_e
        swNONE = 0
        swDECIMAL = 1
        swFRACTION = 2
End Enum

' ----
'  Drawing Paper Sizes
' ----
Public Enum swDwgPaperSizes_e
        swDwgPaperAsize = 0
        swDwgPaperAsizeVertical = 1
        swDwgPaperBsize = 2
        swDwgPaperCsize = 3
        swDwgPaperDsize = 4
        swDwgPaperEsize = 5
        swDwgPaperA4size = 6
        swDwgPaperA4sizeVertical = 7
        swDwgPaperA3size = 8
        swDwgPaperA2size = 9
        swDwgPaperA1size = 10
        swDwgPaperA0size = 11
        swDwgPapersUserDefined = 12
End Enum

' ----
'  Drawing Templates
' ----
Public Enum swDwgTemplates_e
        swDwgTemplateAsize = 0
        swDwgTemplateAsizeVertical = 1
        swDwgTemplateBsize = 2
        swDwgTemplateCsize = 3
        swDwgTemplateDsize = 4
        swDwgTemplateEsize = 5
        swDwgTemplateA4size = 6
        swDwgTemplateA4sizeVertical = 7
        swDwgTemplateA3size = 8
        swDwgTemplateA2size = 9
        swDwgTemplateA1size = 10
        swDwgTemplateA0size = 11
        swDwgTemplateCustom = 12
        swDwgTemplateNone = 13
End Enum

' ----
'  Drawing Templates
' ----
Public Enum swStandardViews_e
        swFrontView = 1
        swBackView = 2
        swLeftView = 3
        swRightView = 4
        swTopView = 5
        swBottomView = 6
        swIsometricView = 7
        swTrimetricView = 8
        swDimetricView = 9
End Enum

' ----
'  Repaint Notification types
' ----
Public Enum swRepaintTypes_e
        swStandardUpdate = 0
        swLightUpdate = 1
        swMaterialUpdate = 2
        swSectionedUpdate = 3
        swExplodedUpdate = 4
        swInsertSketchUpdate = 5
        swViewDisplayUpdate = 6
        swDamageRepairUpdate = 7
        swSelectionUpdate = 8
        swSectionedExitUpdate = 9
        swScrollViewUpdate = 10
End Enum

' ----
'  User Interface State
' ----
Public Enum swUIStates_e
        swIsHiddenInFeatureMgr = 1
End Enum

' ----
'  Type names
' ----
'  Body Features
'  Sheet Metal features
'  Drawing Related
'  Assembly Related
'  Reference Geometry
'  Misc
'  Symbol markers
' ----
'  Surface Types.  For use with Surface::Identity method.
' ----
' ----
'  Curve Types.  For use with Curve::Identity method.
' ----
' ----
'  This is the beginning of time. Used to initialize su_CTime.
' ----
'  Items that can be configured to have a line style in drawings.
Public Enum swLineTypes_e
        swLF_VISIBLE = 0
        swLF_HIDDEN = 1
        swLF_SKETCH = 2
        swLF_DETAIL = 3
        swLF_SECTION = 4
        swLF_DIMENSION = 5
        swLF_CENTER = 6
        swLF_HATCH = 7
        swLF_TANGENT = 8
End Enum

'  Dimension tolerance types
Public Enum swTolType_e
        swTolNONE = 0
        swTolBASIC = 1
        swTolBILAT = 2
        swTolLIMIT = 3
        swTolSYMMETRIC = 4
        swTolMIN = 5
        swTolMAX = 6
        swTolMETRIC = 7 ' same as swTolFIT as of 2001Plus
        swTolFIT = 7
        swTolFITWITHTOL = 8
        swTolFITTOLONLY = 9
End Enum

Public Enum swFitType_e
        swFitUSER = 0
        swFitCLEARANCE = 1
        swFitTRANSITIONAL = 2
        swFitPRESS = 3
End Enum

'  Tolerances which the user can set using Modeler::SetTolerances
Public Enum swTolerances_e
        swBSCurveOutputTol = 0  ' 3D bspline curve output tolerance (meters)
        swBSCurveNonRationalOutputTol = 1       ' 3D non-rational bspline curve output tolerance (meters)
        swUVCurveOutputTol = 2  ' 2D trim curve output tolerance (fraction of characteristic min. face dimension)
        swSurfChordTessellationTol = 3  ' chord tolerance or deviation for tessellation for surfaces
        swSurfAngularTessellationTol = 4        ' angular tolerance or deviation for tessellation for surfaces
        swCurveChordTessellationTol = 5 ' chord tolerance or deviation for tessellation for curves
End Enum

' ----
'  Mate Entity Types
'
'   The following are the possible mate entity type ids returned by the function
'   IMateEntity::GetEntityType.
' ----
Public Enum swMateEntityTypes_e
        swMateUnsupported = 0
        swMatePoint = 1
        swMateLine = 2
        swMatePlane = 3
        swMateCylinder = 4
        swMateCone = 5
End Enum

' ----
'  Attribute Callback Support
'
'   The following are the possible callback types for IAttributeDefs
' ----
Public Enum swAttributeCallbackTypes_e
        swACBDelete = 0
End Enum

Public Enum swAttributeCallbackOptions_e
        swACBRequiresCallback = 1
End Enum

Public Enum swAttributeCallbackReturnValues_e
        swACBDeleteIt = 1
End Enum

'  Text reference point position
Public Enum swTextPosition_e
        swUPPER_LEFT = 0
        swLOWER_LEFT = 1
        swCENTER = 2
        swUPPER_RIGHT = 3
        swLOWER_RIGHT = 4
        swUPPER_CENTER = 5
End Enum

' ----
'  The following are the different types of topology resulting from a call to GetTrimCurves
' ----
Public Enum swTopologyTypes_e
        swTopologyNull = 0
        swTopologyCoEdge = 1
        swTopologyVertex = 2
End Enum

' ----
'  Attributes associated entity state
' ----
Public Enum swAssociatedEntityStates_e
        swIsEntityInvalid = 0
        swIsEntitySuppressed = 1
        swIsEntityAmbiguous = 2
        swIsEntityDeleted = 3
End Enum

' ---
'  Search Folder Types
' ---
Public Enum swSearchFolderTypes_e
        swDocumentType = 0
End Enum

' ---
'  User Preference Toggles.
'  The different User Preference Toggles for GetUserPreferenceToggle & SetUserPreferenceToggle
' ---
Public Enum swUserPreferenceToggle_e
        swUseFolderSearchRules = 0
        swDisplayArcCenterPoints = 1
        swDisplayEntityPoints = 2
        swIgnoreFeatureColors = 3
        swDisplayAxes = 4
        swDisplayPlanes = 5
        swDisplayOrigins = 6
        swDisplayTemporaryAxes = 7
        swDxfMapping = 8
        swSketchAutomaticRelations = 9
        swInputDimValOnCreate = 10
        swFullyConstrainedSketchMode = 11
        swXTAssemSaveFormat = 12
        swDisplayCoordSystems = 13
        swExtRefOpenReadOnly = 14
        swExtRefNoPromptOrSave = 15
        swExtRefMultipleContexts = 16
        swExtRefAutoGenNames = 17
        swExtRefUpdateCompNames = 18
        swDisplayReferencePoints = 19
        swUseShadedFaceHighlight = 20
        swDXFDontShowMap = 21
        swThumbnailGraphics = 22
        swUseAlphaTransparency = 23
        swDynamicDrawingViewActivation = 24
        swAutoLoadPartsLightweight = 25
        swIGESStandardSetting = 26
        swIGESNurbsSetting = 27
        swTiffPrintScaleToFit = 28
        swDisplayVirtualSharps = 29
        swUpdateMassPropsDuringSave = 30
        swDisplayAnnotations = 31
        swDisplayFeatureDimensions = 32
        swDisplayReferenceDimensions = 33
        swDisplayAnnotationsUseAssemblySettings = 34
        swDisplayNotes = 35
        swDisplayGeometricTolerances = 36
        swDisplaySurfaceFinishSymbols = 37
        swDisplayWeldSymbols = 38
        swDisplayDatums = 39
        swDisplayDatumTargets = 40
        swDisplayCosmeticThreads = 41
        swDetailingDisplayWithBrokenLeaders = 42
        swDetailingDualDimensions = 43
        swDetailingDisplayDatumsPer1982 = 44
        swDetailingDisplayAlternateSection = 45
        swDetailingCenterMarkShowLines = 46
        swDetailingFixedSizeWeldSymbol = 47
        swDetailingDimsShowParenthesisByDefault = 48
        swDetailingDimsSnapTextToGrid = 49
        swDetailingDimsCenterText = 50
        swDetailingRadialDimsDisplay2ndOutsideArrow = 51
        swDetailingRadialDimsArrowsFollowText = 52
        swDetailingDimLeaderOverrideStandard = 53
        swDetailingNotesDisplayWithBentLeader = 54
        swDisplayTextAtSameSizeAlways = 55
        swDisplayOnlyInViewOfCreation = 56
        swGridDisplay = 57
        swGridDisplayDashed = 58
        swGridAutomaticScaling = 59
        swSnapToPoints = 60
        swSnapToAngle = 61
        swUnitsLinearRoundToNearestFraction = 62
        swUnitsLinearFeetAndInchesFormat = 63
        swFeatureManagerEnsureVisible = 64
        swFeatureManagerNameFeatureWhenCreated = 65
        swFeatureManagerKeyboardNavigation = 66
        swFeatureManagerDynamicHighlight = 67
        swColorsGradientPartBackground = 68
        swSTLBinaryFormat = 69
        swSTLShowInfoOnSave = 70
        swSTLDontTranslateToPositive = 71
        swSTLComponentsIntoOneFile = 72
        swSTLCheckForInterference = 73
        swOpenLastUsedDocumentAtStart = 74
        swSingleCommandPerPick = 75
        swShowDimensionNames = 76
        swShowErrorsEveryRebuild = 77
        swMaximizeDocumentOnOpen = 78
        swEditDesignTableInSeparateWindow = 80
        swEnablePropertyManager = 81
        swUseSystemSeparatorForDims = 82
        swUseEnglishLanguage = 83
        swDrawingAutomaticModelDimPlacement = 84
        swDrawingDisplayViewBorders = 85
        swAutomaticScaling3ViewDrawings = 86
        swDrawingAutomaticBomUpdate = 87
        swDrawingSelectHiddenEntities = 88
        swDrawingCreateDetailAsCircle = 89
        swAutomaticDrawingViewUpdate = 90
        swDrawingDetailInferCorner = 91
        swDrawingDetailInferCenter = 92
        swDrawingViewShowContentsWhileDragging = 93
        swSketchAlternateSplineCreation = 94
        swSketchInferFromModel = 95
        swSketchPromptToCloseSketch = 96
        swSketchCreateSketchOnNewPart = 97
        swSketchOverrideDimensionsOnDrag = 98
        swSketchDisplayPlaneWhenShaded = 99
        swSketchOverdefiningDimsPromptToSetState = 100
        swSketchOverdefiningDimsSetDrivenByDefault = 101
        swPerformanceVerifyOnRebuild = 102
        swPerformanceDynamicUpdateOnMove = 103
        swPerformanceAlwaysGenerateCurvature = 104
        swPerformanceWin95ZoomClipping = 105
        swIGESDuplicateEntities = 106
        swIGESHighTrimCurveAccuracy = 107
        swIGESExportSketchEntities = 108
        swIGESComponentsIntoOneFile = 109
        swIGESFlattenAssemHierarchy = 110
        swAlwaysUseDefaultTemplates = 111
        swUseSimpleOpenGL = 112
        swShowRefGeomName = 113
        swUseShadedPreview = 114
        swEdgesHiddenEdgeSelectionInWireframe = 115
        swEdgesHiddenEdgeSelectionInHLR = 116
        swEdgesRepaintAfterSelectionInHLR = 117
        swEdgesHighlightFeatureEdges = 118
        swEdgesDynamicHighlight = 119
        swEdgesHighQualityDisplay = 120
        swEdgesOpenEdgesDifferentColor = 121
        swEnableConfirmationCorner = 122
        swAutoShowPropertyManager = 123
        swIncontextFeatureHolderVisibility = 124
        swTransparencyHighQualityDynamic = 125
        swEdgesShadedEdgesDifferentColor = 126
        swEdgesAntiAlias = 127
        swPageSetupPrinterUsePrinterMargin = 128
        swPageSetupPrinterDrawingScaleToFit = 129
        swPageSetupPrinterPartAsmPrintWindow = 130
        swDisplayShadowsInShadedMode = 131
        swDrawingViewSmoothDynamicMotion = 132
        swDrawingEliminateDuplicateDimsOnInsert = 133
        swRapidDraftPrintOutOfSynchWaterMark = 134
        swDrawingViewAutoHideComponents = 135
        swEdgesDisplayShadedPlanes = 136
        swPlaneDisplayShowEdges = 137
        swPlaneDisplayShowIntersections = 138
        swColorsUseSpecifiedEditColors = 139
        swEnablePerformanceEmail = 141
        swSnapOnlyIfGridDisplayed = 142
        swDetailingBalloonsDisplayWithBentLeader = 143
        swBOMConfigurationLocked = 144
        swBOMConfigurationUseDocumentFont = 145
        swBOMConfigurationUseSummaryInfo = 146
        swBOMConfigurationAlignBottom = 147
        swBOMContentsDisplayAtTop = 148
        swBOMControlIdFromAssembly = 149
        swBOMControlMissingRows = 150
        swBOMControlSplitTable = 151
        swAutomaticDrawingViewUpdateDefault = 152
        swAutomaticDrawingViewUpdateForceOff = 153
        swAnnotationDisplayHideDanglingDim = 154
        swDetailingDimBreakAroundArrow = 155
        swDetailingDimensionsToleranceUseParentheses = 156
        swDetailingDimensionsToleranceUseDimensionFont = 157
        swImageQualityApplyToAllReferencedPartDoc = 158
        swPrintBackground = 159
        swEDrawingsCompression = 160
        swImportSolidSurface = 161
        swImportFreeCurves = 162
        swImport2dCurvesAs2dSketch = 163
        swLargeAsmModeAutoLoadLightweight = 166
        swLargeAsmModeUpdateMassPropsOnSave = 167
        swLargeAsmModeAutoRecover = 168
        swLargeAsmModeRemoveDetail = 169
        swLargeAsmModeHideAllItems = 170
        swLargeAsmModeDynHighlightFeatureMgr = 171
        swLargeAsmModeDynHighlightGraphicsView = 172
        swLargeAsmModeAntiAliasEdgesFastMode = 173
        swLargeAsmModeShadowsShadedMode = 174
        swLargeAsmModeTransparencyNormalViewMode = 175
        swLargeAsmModeTransparencyDynamicViewMode = 176
        swLargeAsmModeShowContentsDragDrawView = 177
        swLargeAsmModeSmoothDynamicMotionDrawView = 178
        swLargeAsmModeDrawingHLREdgesWhenShaded = 179
        swLargeAsmModeAutoHideCompsDrawViewCreation = 180
        swLargeAsmModeDrawingAutoLoadModels = 181
        swLargeAsmModeAlwaysGenerateCurvature = 182
        swImportStepConfigData = 183
        swIGESExportSolidAndSurface = 184
        swIGESExportFreeCurves = 185
        swIGESExportAsWireframe = 186
        swDetailingDimensionsAngularToleranceUseParentheses = 187
        swDetailingDimensionsToleranceFitTolUseDimensionFont = 188
        swDetailingAutoInsertCenterMarks = 189
        swDetailingAutoInsertCenterLines = 190
        swSTLPreview = 191
        swDetailingCenterMarkUseCenterLine = 192
        swMaterialPropertySolidFill = 193
        swSaveEModelData = 194
        swDisplayCurves = 195
        swDisplaySketches = 196
        swDisplayAllAnnotations = 197
        swViewDisplayHideAllTypes = 198
End Enum

' ---
'  User Preference Integer Values
'  The different User Preference Integer Values for GetUserPreferenceIntegerValue & SetUserPreferenceIntegerValue
' ---
Public Enum swUserPreferenceIntegerValue_e
        swDxfVersion = 0
        swDxfOutputFonts = 1
        swDxfMappingFileIndex = 2
        swAutoSaveInterval = 3
        swResolveLightweight = 4
        swAcisOutputVersion = 5
        swTiffScreenOrPrintCapture = 6
        swTiffImageType = 7
        swTiffCompressionScheme = 8
        swTiffPrintDPI = 9
        swTiffPrintPaperSize = 10
        swTiffPrintScaleFactor = 11
        swCreateBodyFromSurfacesOption = 12     '  Used by API CreateBodyFromSurfaces
        swDetailingDimensionStandard = 13
        swDetailingDualDimPosition = 14
        swDetailingDimTrailingZero = 15
        swDetailingArrowStyleForDimensions = 16
        swDetailingDimensionArrowPosition = 17
        swDetailingLinearDimLeaderStyle = 18
        swDetailingRadialDimLeaderStyle = 19
        swDetailingAngularDimLeaderStyle = 20
        swDetailingLinearToleranceStyle = 21
        swDetailingAngularToleranceStyle = 22
        swDetailingToleranceTextSizing = 23
        swDetailingLinearDimPrecision = 24
        swDetailingLinearTolPrecision = 25
        swDetailingAltLinearDimPrecision = 26
        swDetailingAltLinearTolPrecision = 27
        swDetailingAngularDimPrecision = 28
        swDetailingAngularTolPrecision = 29
        swDetailingNoteTextAlignment = 30
        swDetailingNoteLeaderSide = 31
        swDetailingBalloonStyle = 32
        swDetailingBalloonFit = 33
        swDetailingBOMBalloonStyle = 34
        swDetailingBOMBalloonFit = 35
        swDetailingBOMUpperText = 36
        swDetailingBOMLowerText = 37
        swDetailingArrowStyleForEdgeVertexAttachment = 38
        swDetailingArrowStyleForFaceAttachment = 39
        swDetailingArrowStyleForUnattached = 40
        swDetailingVirtualSharpStyle = 41
        swGridMinorLinesPerMajor = 42
        swSnapPointsPerMinor = 43
        swImageQualityShaded = 44
        swImageQualityWireframe = 45
        swImageQualityWireframeValue = 46
        swUnitsLinear = 47
        swUnitsLinearDecimalDisplay = 48
        swUnitsLinearDecimalPlaces = 49
        swUnitsLinearFractionDenominator = 50
        swUnitsAngular = 51
        swUnitsAngularDecimalPlaces = 52
        swLineFontVisibleEdgesThickness = 53
        swLineFontVisibleEdgesStyle = 54
        swLineFontHiddenEdgesThickness = 55
        swLineFontHiddenEdgesStyle = 56
        swLineFontSketchCurvesThickness = 57
        swLineFontSketchCurvesStyle = 58
        swLineFontDetailCircleThickness = 59
        swLineFontDetailCircleStyle = 60
        swLineFontSectionLineThickness = 61
        swLineFontSectionLineStyle = 62
        swLineFontDimensionsThickness = 63
        swLineFontDimensionsStyle = 64
        swLineFontConstructionCurvesThickness = 65
        swLineFontConstructionCurvesStyle = 66
        swLineFontCrosshatchThickness = 67
        swLineFontCrosshatchStyle = 68
        swLineFontTangentEdgesThickness = 69
        swLineFontTangentEdgesStyle = 70
        swLineFontDetailBorderThickness = 71
        swLineFontDetailBorderStyle = 72
        swLineFontCosmeticThreadThickness = 73
        swLineFontCosmeticThreadStyle = 74
        swStepAP = 75
        swHiddenEdgeDisplayDefault = 76
        swTangentEdgeDisplayDefault = 77
        swSTLQuality = 78
        swDrawingProjectionType = 79
        swDrawingPrintCrosshatchOutOfDateViews = 80
        swPerformanceAssemRebuildOnLoad = 81
        swLoadExternalReferences = 82
        swIGESRepresentation = 83
        swIGESSystem = 84
        swIGESCurveRepresentation = 85
        swViewRotationMouseSpeed = 86
        swBackupCopiesPerDocument = 87
        swCheckForOutOfDateLightweightComponents = 88
        swParasolidOutputVersion = 89
        swLineFontHideTangentEdgeThickness = 90
        swLineFontHideTangentEdgeStyle = 91
        swLineFontViewArrowThickness = 92
        swLineFontViewArrowStyle = 93
        swEdgesHiddenEdgeDisplay = 94
        swEdgesTangentEdgeDisplay = 95
        swEdgesShadedModeDisplay = 96
        swDetailingBOMStackedBalloonStyle = 97
        swDetailingBOMStackedBalloonFit = 98
        swSystemColorsViewportBackground = 99
        swSystemColorsTopGradientColor = 100
        swSystemColorsBottomGradientColor = 101
        swSystemColorsDynamicHighlight = 102
        swSystemColorsHighlight = 103
        swSystemColorsSelectedItem1 = 104
        swSystemColorsSelectedItem2 = 105
        swSystemColorsSelectedItem3 = 106
        swSystemColorsSelectedFaceShaded = 107
        swSystemColorsDrawingsVisibleModelEdge = 108
        swSystemColorsDrawingsHiddenModelEdge = 109
        swSystemColorsDrawingsPaperBorder = 110
        swSystemColorsDrawingsPaperShadow = 111
        swSystemColorsImportedDrivingAnnotation = 112
        swSystemColorsImportedDrivenAnnotation = 113
        swSystemColorsSketchOverDefined = 114
        swSystemColorsSketchFullyDefined = 115
        swSystemColorsSketchUnderDefined = 116
        swSystemColorsSketchInvalidGeometry = 117
        swSystemColorsSketchNotSolved = 118
        swSystemColorsGridLinesMinor = 119
        swSystemColorsGridLinesMajor = 120
        swSystemColorsConstructionGeometry = 121
        swSystemColorsDanglingDimension = 122
        swSystemColorsText = 123
        swSystemColorsAssemblyEditPart = 124
        swSystemColorsAssemblyEditPartHiddenLines = 125
        swSystemColorsAssemblyNonEditPart = 126
        swSystemColorsInactiveEntity = 127
        swSystemColorsTemporaryGraphics = 128
        swSystemColorsTemporaryGraphicsShaded = 129
        swSystemColorsActiveSelectionListBox = 130
        swSystemColorsSurfacesOpenEdge = 131
        swSystemColorsTreeViewBackground = 132
        swAcisOutputUnits = 133
        swSystemColorsShadedEdge = 134
        swDxfOutputLineStyles = 135
        swDxfOutputNoScale = 136
        swPageSetupPrinterOrientation = 138
        swPageSetupPrinterDrawingColor = 139
        swImportCheckAndRepair = 140
        swUseCustomizedImportTolerance = 141
        swStepExportPreference = 142
        swEdgesInContextEditTransparencyType = 143
        swEdgesInContextEditTransparency = 144
        swPlaneDisplayFrontFaceColor = 145
        swPlaneDisplayBackFaceColor = 146
        swPlaneDisplayTransparency = 147
        swPlaneDisplayIntersectionLineColor = 148
        swDetailingDatumDisplayType = 149
        swBOMConfigurationAnchorType = 150
        swBOMConfigurationWhatToShow = 151
        swBOMControlMissingRowDisplay = 152
        swBOMControlSplitDirection = 153
        swDetailingChamferDimLeaderStyle = 154
        swDetailingChamferDimTextStyle = 155
        swDetailingChamferDimXStyle = 156
        swDocumentColorFeatBend = 157
        swDocumentColorFeatBoss = 158
        swDocumentColorFeatCavity = 159
        swDocumentColorFeatChamfer = 160
        swDocumentColorFeatCut = 161
        swDocumentColorFeatLoftCut = 162
        swDocumentColorFeatSurfCut = 163
        swDocumentColorFeatSweepCut = 164
        swDocumentColorFeatWeldBead = 165
        swDocumentColorFeatExtrude = 166
        swDocumentColorFeatFillet = 167
        swDocumentColorFeatHole = 168
        swDocumentColorFeatLibrary = 169
        swDocumentColorFeatLoft = 170
        swDocumentColorFeatMidSurface = 171
        swDocumentColorFeatPattern = 172
        swDocumentColorFeatRefSurface = 173
        swDocumentColorFeatRevolution = 174
        swDocumentColorFeatShell = 175
        swDocumentColorFeatDerivedPart = 176
        swDocumentColorFeatSweep = 177
        swDocumentColorFeatThicken = 178
        swDocumentColorFeatRib = 179
        swDocumentColorFeatDome = 180
        swDocumentColorFeatForm = 181
        swDocumentColorFeatShape = 182
        swDocumentColorFeatReplaceFace = 183
        swDocumentColorWireFrame = 184
        swDocumentColorShading = 185
        swDocumentColorHidden = 186
        swLineFontExplodedLinesThickness = 187
        swLineFontExplodedLinesStyle = 188
        swSystemColorsRefTriadX = 189
        swSystemColorsRefTriadY = 190
        swSystemColorsRefTriadZ = 191
        swAcisOutputGeometryPreference = 192
        swSystemColorsDTDim = 193
        swLargeAsmModeThreshold = 194
        swLargeAsmModeAutoActivate = 195
        swLargeAsmModeCheckOutOfDateLightweight = 196
        swLargeAsmModeAutoRecoverCount = 197
        swLargeAsmModeDisplayModeForNewDrawViews = 198
        swLineFontBreakLineThickness = 199
        swLineFontBreakLineStyle = 200
        swSaveAssemblyAsPartOptions = 201
        swDetailingDimensionTextAlignmentVertical = 202
        swDetailingDimensionTextAlignmentHorizontal = 203
        swDetailingToleranceFitTolTextSizing = 204
        swImportUnitPreference = 205
        swImportCurvePreference = 206
        swImportUseBrep = 207
        swImportStlVrmlModelType = 208
        swSystemColorsSelectedItem4 = 209
End Enum

' ---
'  User Preference Double Values
'  The different User Preference Double Values for GetUserPreferenceDoubleValue & SetUserPreferenceDoubleValue
' ---
Public Enum swUserPreferenceDoubleValue_e
        swDetailingNoteFontHeight = 0
        swDetailingDimFontHeight = 1
        swSTLDeviation = 2
        swSTLAngleTolerance = 3
        swSpinBoxMetricLengthIncrement = 4
        swSpinBoxEnglishLengthIncrement = 5
        swSpinBoxAngleIncrement = 6
        swMaterialPropertyDensity = 7
        swTiffPrintPaperWidth = 8
        swTiffPrintPaperHeight = 9
        swTiffPrintDrawingPaperHeight = 8
        swTiffPrintDrawingPaperWidth = 9
        swDetailingCenterlineExtension = 10
        swDetailingBreakLineGap = 11
        swDetailingCenterMarkSize = 12
        swDetailingWitnessLineGap = 13
        swDetailingWitnessLineExtension = 14
        swDetailingObjectToDimOffset = 15
        swDetailingDimToDimOffset = 16
        swDetailingMaxLinearToleranceValue = 17
        swDetailingMinLinearToleranceValue = 18
        swDetailingMaxAngularToleranceValue = 19
        swDetailingMinAngularToleranceValue = 20
        swDetailingToleranceTextScale = 21
        swDetailingToleranceTextHeight = 22
        swDetailingNoteBentLeaderLength = 23
        swDetailingArrowHeight = 24
        swDetailingArrowWidth = 25
        swDetailingArrowLength = 26
        swDetailingSectionArrowHeight = 27
        swDetailingSectionArrowWidth = 28
        swDetailingSectionArrowLength = 29
        swGridMajorSpacing = 30
        swSnapToAngleValue = 31
        swImageQualityShadedDeviation = 32
        swDrawingDefaultSheetScaleNumerator = 33
        swDrawingDefaultSheetScaleDenominator = 34
        swDrawingDetailViewScale = 35
        swViewRotationArrowKeys = 36
        swMateAnimationSpeed = 37
        swViewAnimationSpeed = 38
        swDetailingDimBentLeaderLength = 39
        swMaterialPropertyCrosshatchScale = 40
        swMaterialPropertyCrosshatchAngle = 41
        swDrawingAreaHatchScale = 42
        swDrawingAreaHatchAngle = 43
        swPageSetupPrinterTopMargin = 44
        swPageSetupPrinterBottomMargin = 45
        swPageSetupPrinterLeftMargin = 46
        swPageSetupPrinterRightMargin = 47
        swPageSetupPrinterThinLineWeight = 48
        swPageSetupPrinterNormalLineWeight = 49
        swPageSetupPrinterThickLineWeight = 50
        swPageSetupPrinterThick2LineWeight = 51
        swPageSetupPrinterThick3LineWeight = 52
        swPageSetupPrinterThick4LineWeight = 53
        swPageSetupPrinterThick5LineWeight = 54
        swPageSetupPrinterThick6LineWeight = 55
        swPageSetupPrinterDrawingScale = 56
        swPageSetupPrinterPartAsmScale = 57
        swCustomizedImportTolerance = 58
        swDetailingBalloonBentLeaderLength = 60
        swBOMControlSplitHeight = 61
        swAnnotationTextScaleNumerator = 62
        swAnnotationTextScaleDenominator = 63
        swDetailingDimBreakGap = 64
        swCurvatureValue1 = 65
        swCurvatureValue2 = 66
        swCurvatureValue3 = 67
        swCurvatureValue4 = 68
        swCurvatureValue5 = 69
        swDetailingBreakLineExtension = 70
        swDetailingToleranceFitTolTextScale = 71
        swDetailingToleranceFitTolTextHeight = 72
        swDocumentColorAdvancedAmbient = 73
        swDocumentColorAdvancedDiffuse = 74
        swDocumentColorAdvancedSpecularity = 75
        swDocumentColorAdvancedShininess = 76
        swDocumentColorAdvancedTransparency = 77
        swDocumentColorAdvancedEmission = 78
        swDxfOutputScaleFactor = 79
End Enum

' ---
'  User Preference String Values
'  The different User Preference String Values for GetUserPreferenceStringValue & SetUserPreferenceStringValue
' ---
Public Enum swUserPreferenceStringValue_e
        swFileLocationsDocuments = 1
        swFileLocationsPaletteFeatures = 2
        swFileLocationsPaletteParts = 3
        swFileLocationsPaletteFormTools = 4
        swFileLocationsBlocks = 5
        swFileLocationsDocumentTemplates = 6
        swFileLocationsSheetFormat = 7
        swDefaultTemplatePart = 8
        swDefaultTemplateAssembly = 9
        swDefaultTemplateDrawing = 10
        swBackupDirectory = 11
        swFileLocationsBendTable = 12
        swMaterialPropertyCrosshatchPattern = 13
        swDrawingAreaHatchPattern = 14
        swDetailingNextDatumFeatureLabel = 15
        swFileSaveAsCoordinateSystem = 16
        swFileLocationsPaletteAssemblies = 17
End Enum

' ---
'  User Preference String List Values
'  The different User Preference String List Values for GetUserPreferenceStringListValue & SetUserPreferenceStringListValue
' ---
Public Enum swUserPreferenceStringListValue_e
        swDxfMappingFiles = 0
End Enum

' ---
'  User Preference Text Formats
'  The different User Preference Text Formats for Get/SetUserPreferenceTextFormat
' ---
Public Enum swUserPreferenceTextFormat_e
        swDetailingNoteTextFormat = 0
        swDetailingDimensionTextFormat = 1
        swDetailingSectionTextFormat = 2
        swDetailingDetailTextFormat = 3
        swDetailingViewArrowTextFormat = 4
        swDetailingSurfaceFinishTextFormat = 5
        swDetailingWeldSymbolTextFormat = 6
End Enum

' ---
'  View Display States
'  The different View Display States for IModelView::GetDisplayState
' ---
Public Enum swViewDisplayType_e
        swIsViewSectioned = 0
        swIsViewPerspective = 1
        swIsViewShaded = 2
        swIsViewWireFrame = 3
        swIsViewHiddenLinesRemoved = 4
        swIsViewHiddenInGrey = 5
        swIsViewCurvature = 6
End Enum

' ----
'  Control display of internal sketch points
' ----
Public Enum swSkInternalPntOpts_e
        swSkPntsOff = 0
        swSkPntsOn = 1
        swSkPntsDefault = 2
End Enum

' ----
'  DXF/DWG Output formats
' ----
Public Enum swDxfFormat_e
        swDxfFormat_R12 = 0
        swDxfFormat_R13 = 1
        swDxfFormat_R14 = 2
        swDxfFormat_R2000 = 3
End Enum

' ---
'  DXF/DWG output arrow directions
' ---
Public Enum swArrowDirection_e
        swINSIDE = 0
        swOUTSIDE = 1
        swSMART = 2
End Enum

' ---
'  Print Properties
'  The different property types for IModelDoc::SetPrintSetUp
' ---
Public Enum swPrintProperties_e
        swPrintPaperSize = 0
        swPrintOrientation = 1
        swPrintPaperLength = 2
        swPrintPaperWidth = 3
End Enum

' ---
'  Tiff Image types
' ---
Public Enum swTiffImageType_e
        swTiffImageBlackAndWhite = 0
        swTiffImageRGB = 1
End Enum

' ---
'  Tiff Image Compression schemes
' ---
Public Enum swTiffCompressionScheme_e
        swTiffUncompressed = 0
        swTiffPackbitsCompression = 1
        swTiffGroup4FaxCompression = 2
End Enum

' ----
'  Body operations.  For use with Body::Operations method.
' ----
Public Enum swBodyOperationError_e
        swBodyOperationUnknownError = -1
        swBodyOperationNoError = 0
        swBodyOperationNonApiBody = 1
        swBodyOperationWrongType = 2
        swBodyOperationBooleanFail = 1058
        swBodyOperationNoIntersect = 1067
        swBodyOperationNonManifold = 547
        swBodyOperationPartialCoincidence = 1040
        swBodyOperationIntersectSolidWithSheets = 972
        swBodyOperationUniteSolidSheet = 543
        swBodyOperationMissingGeom = 96
        swBodyOperationSameToolAndTarget = 545
        swBodyOperationFailGeomCondition = 3
        swBodyOperationFailToCutBody = 4
        swBodyOperationDisjointBodies = 5
        swBodyOperationEmptyBody = 6
        swBodyOperationEmptyInputBody = 7
End Enum

' ---
'  End Conditions.
'  These are used with FeatureBoss, FeatureCut, FeatureExtrusion, etc.
'  Not all types are valid for all body operations.  Some of these end conditions require additional
'  selections (ie - swEndCondUpToSurface, etc.) and some require additional data (ie - swEndCondOffsetFromSurface)
' ---
Public Enum swEndConditions_e
        swEndCondBlind = 0
        swEndCondThroughAll = 1
        swEndCondThroughNext = 2
        swEndCondUpToVertex = 3
        swEndCondUpToSurface = 4
        swEndCondOffsetFromSurface = 5
        swEndCondMidPlane = 6
        swEndCondUpToBody = 7
End Enum

Public Enum swChamferType_e
        swChamferAngleDistance = 1
        swChamferDistanceDistance = 2
        swChamferVertex = 3
        swChamferEqualDistance = 4
End Enum

' ---
'  Line weights
' ---
Public Enum swLineWeights_e
        swLW_NONE = -1
        swLW_THIN = 0
        swLW_NORMAL = 1
        swLW_THICK = 2
        swLW_THICK2 = 3
        swLW_THICK3 = 4
        swLW_THICK4 = 5
        swLW_THICK5 = 6
        swLW_THICK6 = 7
        swLW_NUMBER = 8
        swLW_LAYER = 9
End Enum

' ---
'  Toolbar States.  For use with ISldWorks::GetToolbarState()
' ---
Public Enum swToolbarStates_e
        swToolbarHidden = 0
End Enum

' ----
'  Summary info fields for use with IModelDoc::Get/SetSummaryInfo
' ----
Public Enum swSummInfoField_e
        swSumInfoTitle = 0
        swSumInfoSubject = 1
        swSumInfoAuthor = 2
        swSumInfoKeywords = 3
        swSumInfoComment = 4
        swSumInfoSavedBy = 5
        swSumInfoCreateDate = 6
        swSumInfoSaveDate = 7
        swSumInfoCreateDate2 = 8
        swSumInfoSaveDate2 = 9
End Enum

'  CPropertySheet enumerated types.
'  For use with the ISldWorks::PropertySheetCreateNotify notification
Public Enum swPropSheetType_e
        swPropSheetNotValid = 0
        swPropSheetLighting = 1
        swPropSheetToolsOptions = 2
        swPropSheetAmbientLight = 3
        swPropSheetDirectionalLight = 4
        swPropSheetPositionLight = 5
        swPropSheetSpotLight = 6
End Enum

Public Enum swWindowState_e
        swWindowNormal = 0
        swWindowMaximized = 1
        swWindowMinimized = 2
End Enum

'  Possible values for Witness Line visibility, for use by
'  auDisplayDimension_c::GetWitnessVisibility and SetWitnessVisibility.
Public Enum swWitnessLineVisibility_e
        swWitnessLineBoth = 0   '  BOTH witness lines are displayed
        swWitnessLineFirst = 1  '  only FIRST witness line is displayed
        swWitnessLineSecond = 2 '  only SECOND witness line is displayed
        swWitnessLineNone = 3   '  NEITHER witness line is displayed
End Enum

'  Possible values for Leader Line visibility, for use by
'  auDisplayDimension_c::GetLeaderVisibility and SetLeaderVisibility.
Public Enum swLeaderLineVisibility_e
        swLeaderLineBoth = 0    '  BOTH leader lines are displayed
        swLeaderLineFirst = 1   '  only FIRST leader line is displayed
        swLeaderLineSecond = 2  '  only SECOND leader line is displayed
        swLeaderLineNone = 3    '  NEITHER leader line is displayed
End Enum

'  Possible values for Arrow positions, for use by
'  auDisplayDimension_c::GetArrowSide and SetArrowSide.
Public Enum swDimensionArrowsSide_e
        swDimArrowsInside = 0   '  place arrows INSIDE of the witness lines
        swDimArrowsOutside = 1  '  place arrows OUTSIDE of the witness lines
        swDimArrowsSmart = 2    '  place arrows inside if the text and arrows fit, outside if not
        swDimArrowsFollowDoc = 3        '  place arrows the same as the document default for placing arrows
End Enum

'  The different parts of the dimension text, for use by
'  auDisplayDimension_c::GetText and SetText.
Public Enum swDimensionTextParts_e
        swDimensionTextAll = 0  '  all pieces of text (used only by SetText)
        swDimensionTextPrefix = 1       '  the prefix portion of the text
        swDimensionTextSuffix = 2       '  the suffix portion of the text
        swDimensionTextCalloutAbove = 3 '  the callout portion of the text, above the dimension
        swDimensionTextCalloutBelow = 4 '  the callout portion of the text, below the dimension
End Enum

Public Enum swTopology_e
        swTopoSolidBody = 1
        swTopoSheetBody = 2
        swTopoWireBody = 3
        swTopoMinimumBody = 4
End Enum

Public Enum swTopoEntity_e
        swTopoVertex = 1
        swTopoEdge = 2
        swTopoLoop = 3
        swTopoFace = 4
        swTopoShell = 5
        swTopoBody = 6
End Enum

'  The alignment information possible for Views.  For use with auDrView_c::GetAlignment.
Public Enum swViewAlignment_e
        swViewAlignNone = 0     '  this view has no alignment restrictions
        swViewAlignedChildren = 1       '  this view has children aligned with it
        swViewAligned = 2       '  this view is aligned with a parent view
        swViewAlignBoth = 3     '  this view is aligned and has aligned children
End Enum

' Toolbars
Public Enum swToolbar_e
        swSketchToolsToolbar = 0
        swMainToolbar = 1
        swStandardToolbar = 2
        swViewToolbar = 3
        swSketchRelationsToolbar = 4
        swMacroToolbar = 5
        swSketchToolbar = 6
        swAssemblyToolbar = 7
        swDrawingToolbar = 8
        swAnnotationToolbar = 9
        swWebToolbar = 10
        swFeatureToolbar = 11
        swFontToolbar = 12
        swLineToolbar = 13
        swSelectionFilterToolbar = 14
        swReferenceGeometryToolbar = 15
        swStandardViewsToolbar = 16
        swToolsToolbar = 17
        swCurvesToolbar = 18
        swMoldToolsToolbar = 19
        swSheetMetalToolbar = 20
        swSurfacesToolbar = 21
        swAlignToolbar = 22
        swLayerToolbar = 23
        sw2Dto3DToolbar = 24
        swRoutingToolbar = 25
        swSimulationToolbar = 26
End Enum

'  Annotations
Public Enum swInsertAnnotation_e
        swInsertCThreads = &H1
        swInsertDatums = &H2
        swInsertDatumTargets = &H4
        swInsertDimensions = &H8
        swInsertInstanceCounts = &H10
        swInsertGTols = &H20
        swInsertNotes = &H40
        swInsertSFSymbols = &H80
        swInsertWelds = &H100
        swInsertAxes = &H200
        swInsertCurves = &H400
        swInsertPlanes = &H800
        swInsertSurfaces = &H1000
        swInsertPoints = &H2000
        swInsertOrigins = &H4000
End Enum

'  MessageBox values
Public Enum swMessageBoxIcon_e
        swMbWarning = 1
        swMbInformation = 2
        swMbQuestion = 3
        swMbStop = 4
End Enum

Public Enum swMessageBoxBtn_e
        swMbAbortRetryIgnore = 1
        swMbOk = 2
        swMbOkCancel = 3
        swMbRetryCancel = 4
        swMbYesNo = 5
        swMbYesNoCancel = 6
End Enum

Public Enum swMessageBoxResult_e
        swMbHitAbort = 1
        swMbHitIgnore = 2
        swMbHitNo = 3
        swMbHitOk = 4
        swMbHitRetry = 5
        swMbHitYes = 6
        swMbHitCancel = 7
End Enum

'  Annotation types
Public Enum swAnnotationType_e
        swCThread = 1
        swDatumTag = 2
        swDatumTargetSym = 3
        swDisplayDimension = 4
        swGTol = 5
        swNote = 6
        swSFSymbol = 7
        swWeldSymbol = 8
        swCustomSymbol = 9
        swDowelSym = 10
        swLeader = 11
        swBlock = 12
        swCenterMarkSym = 13
        swTableAnnotation = 14
End Enum

'  The possible Driven States for Dimensions.  For use with auDimension_c::DrivenState.
Public Enum swDimensionDrivenState_e
        swDimensionDrivenUnknown = 0    '  the driven/driving state is unknown
        swDimensionDriven = 1   '  the dimension is a driven dimension
        swDimensionDriving = 2  '  the dimension is a driving dimension
End Enum

Public Enum swFileLoadError_e
        swGenericError = &H1
        swFileNotFoundError = &H2
        swIdMatchError = &H4    '  NO LONGER USED as of OpenDoc6, moved to swFileLoadWarning_e
        swReadOnlyWarn = &H8    '  NO LONGER USED as of OpenDoc6, moved to swFileLoadWarning_e
        swSharingViolationWarn = &H10   '  NO LONGER USED as of OpenDoc6, moved to swFileLoadWarning_e
        swDrawingANSIUpdateWarn = &H20  '  NO LONGER USED as of OpenDoc6, moved to swFileLoadWarning_e
        swSheetScaleUpdateWarn = &H40   '  NO LONGER USED as of OpenDoc6, moved to swFileLoadWarning_e
        swNeedsRegenWarn = &H80 '  NO LONGER USED as of OpenDoc6, moved to swFileLoadWarning_e
        swBasePartNotLoadedWarn = &H100 '  NO LONGER USED as of OpenDoc6, moved to swFileLoadWarning_e
        swFileAlreadyOpenWarn = &H200   '  NO LONGER USED as of OpenDoc6, moved to swFileLoadWarning_e
        swInvalidFileTypeError = &H400  '  the type argument passed into the API is not valid
        swDrawingsOnlyRapidDraftWarn = &H800    '  NO LONGER USED as of OpenDoc6, moved to swFileLoadWarning_e
        swViewOnlyRestrictions = &H1000 '  NO LONGER USED as of OpenDoc6, moved to swFileLoadWarning_e
        swFutureVersion = &H2000        '  document being opened is of a future version.
        swViewMissingReferencedConfig = &H4000  '  NO LONGER USED as of OpenDoc6, moved to swFileLoadWarning_e
        swDrawingSFSymbolConvertWarn = &H8000   '  NO LONGER USED as of OpenDoc6, moved to swFileLoadWarning_e
        swFileWithSameTitleAlreadyOpen = &H10000
End Enum

'  Warnings that occured during a Open API, but did NOT cause the save to fail.
Public Enum swFileLoadWarning_e
        swFileLoadWarning_IdMismatch = &H1
        swFileLoadWarning_ReadOnly = &H2
        swFileLoadWarning_SharingViolation = &H4
        swFileLoadWarning_DrawingANSIUpdate = &H8
        swFileLoadWarning_SheetScaleUpdate = &H10
        swFileLoadWarning_NeedsRegen = &H20
        swFileLoadWarning_BasePartNotLoaded = &H40
        swFileLoadWarning_AlreadyOpen = &H80
        swFileLoadWarning_DrawingsOnlyRapidDraft = &H100
        swFileLoadWarning_ViewOnlyRestrictions = &H200
        swFileLoadWarning_ViewMissingReferencedConfig = &H400
        swFileLoadWarning_DrawingSFSymbolConvert = &H800
        swFileLoadWarning_RevolveDimTolerance = &H1000
        swFileLoadWarning_ModelOutOfDate = &H2000
End Enum

'  Errors that caused the Save API to fail.
Public Enum swFileSaveError_e
        swGenericSaveError = &H1
        swReadOnlySaveError = &H2
        swFileNameEmpty = &H4   '  The filename must not be empty
        swFileNameContainsAtSign = &H8  '  The filename can not contain an '@' character
        swFileLockError = &H10
        swFileSaveFormatNotAvailable = &H20     '  The save as file type is not valid
        swFileSaveWithRebuildError = &H40       '  NO LONGER USED IN SW2001PLUS, moved to swFileSaveWarning_e
        swFileSaveAsDoNotOverwrite = &H80       '  The user chose not to overwrite an existing file
        swFileSaveAsInvalidFileExtension = &H100        '  The file extension differs from the Sw document type
End Enum

'  Warnings that occured during a Save API, but did NOT cause the save to fail.
Public Enum swFileSaveWarning_e
        swFileSaveWarning_RebuildError = &H1    '  The file was saved, but with a rebuild error
        swFileSaveWarning_NeedsRebuild = &H2    '  The file was saved, but needs a rebuild
        swFileSaveWarning_ViewsNeedUpdate = &H4 '  The file was saved, but views of inactive sheets need updating
End Enum

Public Enum swActivateDocError_e
        swGenericActivateError = &H1
        swDocNeedsRebuildWarning = &H2
End Enum

'  The suppression information possible for Components.  For use with auComponent_c::Suppression.
Public Enum swComponentSuppressionState_e
        swComponentSuppressed = 0       '  Fully suppressed - nothing is loaded
        swComponentLightweight = 1      '  Featherweight - only graphics data is loaded
        swComponentFullyResolved = 2    '  Fully resolved - model is completly loaded
End Enum

'  The visibility information possible for Components.  For use with auComponent_c::Visibility.
Public Enum swComponentVisibilityState_e
        swComponentHidden = 0
        swComponentVisible = 1
End Enum

'  Possible values for the solving option of components.
Public Enum swComponentSolvingOption_e
        swComponentRigidSolving = 0
        swComponentFlexibleSolving = 1
End Enum

Public Enum swCustomInfoType_e
        swCustomInfoUnknown = 0
        swCustomInfoText = 30   '  VT_LPSTR
        swCustomInfoDate = 64   '  VT_FILETIME
        swCustomInfoNumber = 3  '  VT_I4
        swCustomInfoYesOrNo = 11        '  VT_BOOL
End Enum

Public Enum swComponentResolveStatus_e
        swResolveOk = 0
        swResolveAbortedByUser = 1
        swResolveNotPerformed = 2
        swResolveError = 3
End Enum

Public Enum swSuppressionError_e
        swSuppressionBadComponent = 0
        swSuppressionBadState = 1
        swSuppressionChangeOk = 2
        swSuppressionChangeFailed = 3
End Enum

Public Enum swDynamicMode_e
        swNoDynamics = 0
        swSpinDynamics = 1
        swPanDynamics = 2
        swZoomDynamics = 3
        swUnknownDynamics = 4
        swAnimDynamics = 5
End Enum

'  The justification of text with respect to the note origin.
'  Used by auNote_c::Get and SetTextJustification[AtIndex].
Public Enum swTextJustification_e
        swTextJustificationLeft = 1     '  Text is Left Justified (Top Justified is assumed?)
        swTextJustificationCenter = 2   '  Text is Center Justified (Top Justified is assumed?)
        swTextJustificationRight = 3    '  Text is Top Justified (Top Justified is assumed?)
End Enum

Public Enum swComponentReloadOption_e
        swAlwaysReload = 0
        swDontReloadOldComponents = 1
End Enum

Public Enum swComponentReloadError_e
        swReloadOkay = 0
        swWriteAccessError = 1
        swFutureVersionError = 2
        swModifiedNotReloadedError = 3
        swInvalidOption = 4
        swFileNotSavedError = 5
        swInvalidComponentError = 6
        swUnexpectedError = 7
        swComponentLightWeightError = 8
        swFileDoesntExistError = 9
        swFileInvalidOrSameNameError = 10
        swDocumentHasNoView = 11
        swDocumentAlreadyOpenedError = 12
End Enum

Public Enum swIntersectionType_e
        swIntersectionSIMPLE = 1
        swIntersectionTANGENT = 2
        swIntersectionCOINCIDENCE_START = 3
        swIntersectionCOINCIDENCE_END = 4
End Enum

Public Enum swAddOrdinateDims_e
        swOrdinate = 1
        swVerticalOrdinate = 2
        swHorizontalOrdinate = 3
End Enum

Public Enum swSheetSewingOption_e
        swSewToSolid = 0
        swSewToSheets = 1
        swSewToSolidOrSheets = 2
End Enum

Public Enum swSheetSewingError_e
        swSewingOk = 0
        swBadArgument = 1
        swUnspecifiedError = 2
        swSewingFailed = 3
        swSewingIncomplete = 4
End Enum

Public Enum swBodyType_e
        swAllBodies = -1
        swSolidBody = 0
        swSheetBody = 1
        swWireBody = 2
        swMinimumBody = 3
        swGeneralBody = 4
        swEmptyBody = 5
End Enum

'  Which Configurations the Set Value applies to.
Public Enum swSetValueInConfiguration_e
        swSetValue_UseCurrentSetting = 0        '  Use the setting this parameter currently has
        swSetValue_InThisConfiguration = 1
        swSetValue_InAllConfigurations = 2
End Enum

'  Return status of the Set Value operation.
Public Enum swSetValueReturnStatus_e
        swSetValue_Successful = 0
        swSetValue_Failure = 1  '  failed for an unknown reason
        swSetValue_InvalidValue = 2     '  not a valid value for Change Parameter
        swSetValue_DrivenDimension = 3  '  can not be done on a dimension driven by geometry
        swSetValue_ModelNotLoaded = 4   '  the model must be loaded in order to set this value
End Enum

'  Possible values for the bendState value for the Get/SetBendState APIs.
Public Enum swSMBendState_e
        swSMBendStateNone = 0   '  No bend state - not a sheet metal part
        swSMBendStateSharps = 1 '  Bends are in the sharp state - bends currently not applied
        swSMBendStateFlattened = 2      '  Bends are flattened
        swSMBendStateFolded = 3 '  Bends are fully applied
End Enum

'  Possible return status of Sheet Metal APIs.
Public Enum swSMCommandStatus_e
        swSMErrorNone = 0
        swSMErrorUnknown = 1    '  failed for an unknown reason
        swSMErrorNotAPart = 2   '  Sheet Metal commands only apply to SW Parts
        swSMErrorNotASheetMetalPart = 3 '  the part contains no Sheet Metal features
        swSMErrorInvalidBendState = 4   '  an invalid bend state was specified (Set Bend State)
End Enum

'  Feature error code returned from Feature::GetErrorCode.
Public Enum swFeatureError_e
        swFeatureErrorNone = 0  '  No error
        swFeatureErrorUnknown = 1       '  Unknown error
        swFeatureErrorFilletNoLoop = 10 '  Loop for fillet/chamfer does not exist
        swFeatureErrorFilletNoFace = 11 '  face for fillet/chamfer does not exist
        swFeatureErrorFilletInvalidRadius = 12  '  invalid fillet radius or a face blend fillet recommended
        swFeatureErrorFilletNoEdge = 13 '  Edge for fillet/chamfer does not exist
        swFeatureErrorFilletModelGeometry = 14  '  Failed to create fillet due to model geometry
        swFeatureErrorFilletRadiusTooSmall = 15 '  Radius value is too small
        swFeatureErrorFilletCannotExtend = 16   '  Selected elements cannot be extended to intersect
        swFeatureErrorFilletRadiusEliminateElement = 17 '  Specified radius would eliminate one of the elements
        swFeatureErrorFilletRadiusTooBig = 18   '  Radius is too big or the elements are tangent or nearly tangent
        swFeatureErrorFilletRadiusTooBig2 = 19  '  The radius of the fillet is too large to fit the surrounding geometry. Try adjusting the input geometry and radius values or try using a face blend fillet.
        swFeatureErrorExtrusionDisjoint = 30    '  This feature would create a disjoint body. The direction may be wrong
        swFeatureErrorExtrusionNoEndFound = 31  '  Cannot locate end of feature
        swFeatureErrorExtrusionBadGeometricConditions = 32      '  Unable to create this extruded feature due to geometric conditions
        swFeatureErrorExtrusionCutContourOpenAndClosed = 33     '  Extruded cuts cannot have both open and closed contours
        swFeatureErrorExtrusionCutContourInvalid = 34   '  Extruded cuts require at least one closed or open contour which does not self-intersect
        swFeatureErrorExtrusionOpenCutContourInvalid = 35       '  Open extruded cuts require a single open contour which does not self-intersect
        swFeatureErrorExtrusionBossContourOpenAndClosed = 36    '  Bosses cannot have both open and closed contours
        swFeatureErrorExtrusionBossContourInvalid = 37  '  Bosses require one or more closed contours which do not self-intersect
End Enum

'  Possible values for the saveAsVersion argument of the SaveAs2 API
Public Enum swSaveAsVersion_e
        swSaveAsCurrentVersion = 0      '  default
        swSaveAsSW98plus = 1    '  save SW model in SW98plus model format - NO LONGER SUPPORTED
        swSaveAsFormatProE = 2  '  save Sw part as Pro/E format .prt/.asm extension (not as Sw .prt/.asm)
End Enum

'  Possible values for the leaderType argument of the Get/SetArcLengthLeader APIs
Public Enum swArcLengthLeaderType_e
        swArcLengthLeaderParallel = 1   '  Leaders are parallel to each other
        swArcLengthLeaderRadial = 2     '  Leaders are radial, from the arc center point
End Enum

'  Possible values for the condition argument of the Get/SetArcEndCondition APIs
'  These values should match up with the values in the moArcDimType_e enumeration.
Public Enum swArcEndCondition_e
        swArcEndConditionNone = 0       '  The end point is not related to an arc
        swArcEndConditionCenter = 1     '  The end point is the center of the arc
        swArcEndConditionMin = 2        '  The end point is the nearest point on the arc
        swArcEndConditionMax = 3        '  The end point is the furthest point on the arc
End Enum

Public Enum swDestroyNotifyType_e
        swDestroyNotifyDestroy = 0      '  The view is being destroyed
        swDestroyNotifyHidden = 1       '  The view is actually being hidden, not destroyed
End Enum

Public Enum swSketchSegments_e
        swSketchLINE = 0
        swSketchARC = 1
        swSketchELLIPSE = 2
        swSketchSPLINE = 3
        swSketchTEXT = 4
        swSketchPARABOLA = 5
End Enum

Public Enum swPipingPenetrationStatus_e
        swPenetrationSucceeded = 0
        swPenetrationFailed = 1
        swPenetrationFailedPipeTooWide = 2
        swPenetrationFailedDllNotLoaded = 3
        swPenetrationFailedNoSelection = 4
        swPenetrationFailedNotRouting = 5
        swPenetrationFailedBadSelection = 6
        swPenetrationFailedBadFitting = 7
        swPenetrationFailedAlreadyPenetrating = 8
        swPenetrationFailedMultiBody = 9
End Enum

'  Enumerate the possible different entity types for passing as a notification argument.
'  Currently used by AddItemNotify, RenameItemNotify, and DeleteItemNotify.
Public Enum swNotifyEntityType_e
        swNotifyConfiguration = 1       '  Configuration is being added, renamed, or deleted
        swNotifyComponent = 2
End Enum

Public Enum swRayPtsOpts_e
        swRayPtsOptsNORMALS = &H1
        swRayPtsOptsTOPOLS = &H2
        swRayPtsOptsENTRY_EXIT = &H4
        swRayPtsOptsUNBLOCK = &H8       ' alow the system to respond while waiting
End Enum

Public Enum swRayPtsResults_e
        swRayPtsResultsFACE = &H1
        swRayPtsResultsSILHOUETTE = &H2
        swRayPtsResultsEDGE = &H4
        swRayPtsResultsVERTEX = &H8
        swRayPtsResultsENTER = &H10
        swRayPtsResultsEXIT = &H20
End Enum

'  The different pieces of text within a weld annotation.  (WeldSymbol::GetText)
Public Enum swWeldSymbolTextTypes_e
        swWeldLeftTextAbove = 1 '  The text just to the left of the weld symbol, above the horizontal line
        swWeldSymbolTextAbove = 2       '  The weld symbol, above the horizontal line
        swWeldRightTextAbove = 3        '  The text just to the right of the weld symbol, above the horizontal line
        swWeldStaggerTextAbove = 4      '  The text related to the stagger characteristic, above the horizontal line
        swWeldLeftTextBelow = 5 '  The text just to the left of the weld symbol, below the horizontal line
        swWeldSymbolTextBelow = 6       '  The weld symbol, below the horizontal line
        swWeldRightTextBelow = 7        '  The text just to the right of the weld symbol, below the horizontal line
        swWeldStaggerTextBelow = 8      '  The text related to the stagger characteristic, below the horizontal line
        swWeldProcessText = 9   '  The text related to the process indicators characteristic
End Enum

'  The different cases for contour symbols of a weld annotation.  (WeldSymbol::GetContour/SetText)
Public Enum swWeldSymbolContourTypes_e
        swWeldContourNone = 1
        swWeldContourFlat = 2
        swWeldContourConvex = 3
        swWeldContourConcave = 4
End Enum

'  The different cases for symmetric characteristic of a weld annotation. (WeldSymbol::Get/SetSymmetric)
Public Enum swWeldSymbolSymmetric_e
        swWeldSymmetric = 1     '  The symbol is symmetric on this weld annotation
        swWeldDashedLineOnTop = 2       '  The symbol is not symmetric, with the dashed horizontal line above
        swWeldDashedLineOnBottom = 3    '  The symbol is not symmetric, with the dashed horizontal line below
End Enum

'  The different cases for field or site characteristic of a weld annotation. (WeldSymbol::Get/SetFieldWeld)
Public Enum swWeldSymbolField_e
        swFieldWeldNone = 1     '  No field-site weld marking on this annotation
        swFieldWeldUp = 2       '  The field-site weld marking is pointing up
        swFieldWeldDown = 3     '  The field-site weld marking is pointing down
End Enum

'  The different cases for whether or not a Display Dimension leader is broken or not, and how the text is
'  placed relative to the leader.
Public Enum swDisplayDimensionLeaderText_e
        swSolidLeaderAlignedText = 1
        swBrokenLeaderHorizontalText = 2
        swBrokenLeaderAlignedText = 3
End Enum

Public Enum swLineStyles_e
        swLineCONTINUOUS = 0
        swLineHIDDEN = 1
        swLinePHANTOM = 2
        swLineCHAIN = 3
        swLineCENTER = 4
        swLineSTITCH = 5
        swLineCHAINTHICK = 6
        swLineDEFAULT = 7
End Enum

'  The different types of drawing views.
Public Enum swDrawingViewTypes_e
        swDrawingSheet = 1
        swDrawingSectionView = 2
        swDrawingDetailView = 3
        swDrawingProjectedView = 4
        swDrawingAuxiliaryView = 5
        swDrawingStandardView = 6
        swDrawingNamedView = 7
        swDrawingRelativeView = 8
End Enum

'  For the Sketch Fillet command, the different actions that can be taken when the fillet is being
'  applied to a corner that has constraints.
Public Enum swConstrainedCornerAction_e
        swConstrainedCornerInteract = 0 '  Ask the user whether to Delete the geometry or Stop Processing
        swConstrainedCornerKeepGeometry = 1     '  Keep the constrained geometry in the part
        swConstrainedCornerDeleteGeometry = 2   '  Delete the constrained geometry from the part
        swConstrainedCornerStopProcessing = 3   '  Do not do anything, stop processing immediately
End Enum

'  The different command mode that can be in effect.  Used by the SldWorks::GetMouseDragMode API.
Public Enum swMouseDragMode_e
        swTranslateAssemblyComponent = 1        '  Assembly Component Move mode
        swRotateAssemblyComponentAboutCenter = 2        '  Assembly Component Rotate mode
        swRotateAssemblyComponentAboutAxis = 3  '  Assembly Component Rotate About Axis mode
        swAssemblySmartMates = 4        '  Assembly Component Smart Mate mode
        swRotateView = 5        '  View Rotate mode
        swTranslateView = 6     '  View Translate mode
        swZoomView = 7  '  View Zoom mode
        swZoomToAreaOfView = 8  '  View Zoom To mode
        swInsertDimension = 9   '  Insert Dimension mode
End Enum

'  The different possibilities for types of Datum Target area shapes. (DatumTargetSym::Get/SetTargetShape)
Public Enum swDatumTargetAreaShape_e
        swDatumTargetAreaNone = 0
        swDatumTargetAreaPoint = 1
        swDatumTargetAreaCircle = 2
        swDatumTargetAreaRectangle = 3
End Enum

'  Possible status values for the Edit Part command.
Public Enum swEditPartCommandStatus_e
        swEditPartFailure = -1
        swEditPartAsmMustBeSaved = -2
        swEditPartCompMustBeSelected = -3
        swEditPartCompMustBeResolved = -4
        swEditPartCompMustHaveWriteAccess = -5
        swEditPartSuccessful = 0
        swEditPartCompNotPositioned = &H1
End Enum

'  Possible values for the visibility state of an annotation. (Annotation::Visible property)
Public Enum swAnnotationVisibilityState_e
        swAnnotationVisibilityUnknown = 0
        swAnnotationVisible = 1
        swAnnotationHalfHidden = 2
        swAnnotationHidden = 3
End Enum

'  This is used by the notification LightweightComponentOpenNotify()
Public Enum swOutOfDateStatus_e
        swUnknownState = 0
        swModelUpToDate = 1
        swModelOutOfDate = 2
End Enum

'  This is used by the API GetLocalizedMenuName()
Public Enum swMenuIdentifiers_e
        swFileMenu = 0
        swEditMenu = 1
        swViewMenu = 2
        swInsertMenu = 3
        swToolsMenu = 4
        swWindowMenu = 5
        swHelpMenu = 6
        swDeveloperToolsMenu = 7
        swViewToolbarsMenu = 8
End Enum

'  For InsertScale
Public Enum swScaleType_e
        swScaleAboutCentroid = 0
        swScaleAboutOrigin = 1
        swScaleAboutCoordinateSystem = 2
End Enum

'  For InsertCavity4
Public Enum swCavityScaleType_e
        swAboutCentroid = 0
        swAboutOrigin = 1
        swAboutMoldBaseOrigin = 2
        swAboutCoordinateSystem = 3
End Enum

Public Enum swFeatMgrPane_e
        swFeatMgrPaneTop = 0
        swFeatMgrPaneBottom = 1
        swFeatMgrPaneTopHidden = 2
        swFeatMgrPaneBottomHidden = 3
End Enum

'  Possible values for the swDetailingDualDimPosition User Preference setting.
Public Enum swDetailingDualDimPosition_e
        swDualDimensionsSideBySide = 1
        swDualDimensionsAboveAndBelow = 2
End Enum

'  Possible values for the swDetailingDimTrailingZero User Preference setting
Public Enum swDetailingDimTrailingZero_e
        swDimSmartTrailingZeroes = 0
        swDimShowTrailingZeroes = 1
        swDimRemoveTrailingZeroes = 2
End Enum

'  Possible values for the swDetailingToleranceTextSizing User Preference setting
Public Enum swDetailingToleranceTextSizing_e
        swToleranceTextSizeUsingScaleValue = 1
        swToleranceTextSizeUsingHeightValue = 2
End Enum

'  Possible values for the swDetailingDimensionStandard User Preference setting
Public Enum swDetailingStandard_e
        swDetailingStandardANSI = 1
        swDetailingStandardISO = 2
        swDetailingStandardDIN = 3
        swDetailingStandardJIS = 4
        swDetailingStandardBS = 5
        swDetailingStandardGOST = 6
        swDetailingStandardGB = 7
End Enum

'  Possible values for the swDetailingBOMUpperText and LowerText User Preference settings
Public Enum swDetailingNoteTextContent_e
        swDetailingNoteTextCustom = 1
        swDetailingNoteTextItemNumber = 2
        swDetailingNoteTextQuantity = 3
End Enum

'  Possible values for the swDetailingVirtualSharpStyle User Preference settings
Public Enum swDetailingVirtualSharp_e
        swDetailingVirtualSharpNone = 0
        swDetailingVirtualSharpPlus = 1
        swDetailingVirtualSharpStar = 2
        swDetailingVirtualSharpWitness = 3
        swDetailingVirtualSharpDot = 4
End Enum

'  Different types of dimensions.  Used by DisplayDimension::GetType.
Public Enum swDimensionType_e
        swDimensionTypeUnknown = 0
        swOrdinateDimension = 1
        swLinearDimension = 2
        swAngularDimension = 3
        swArcLengthDimension = 4
        swRadialDimension = 5
End Enum

'  Possible values for the swImageQualityShaded User Preference setting
Public Enum swImageQualityShaded_e
        swShadedImageQualityCoarse = 1
        swShadedImageQualityFine = 2
        swShadedImageQualityCustom = 3
End Enum

'  Possible values for the swImageQualityWireframe User Preference setting
Public Enum swImageQualityWireframe_e
        swWireframeImageQualityOptimal = 1
        swWireframeImageQualityCustom = 2
End Enum

Public Enum swLoadDetachedModelRules_e
        swLoadDetachedModelPrompt = 0
        swLoadDetachedModelAuto = 1
        swDoNotLoadDetachedModel = 2
End Enum

'  Possible value for different methods of display tangent edges.  (View::Get/SetDisplayTangentEdges2)
Public Enum swDisplayTangentEdges_e
        swTangentEdgesHidden = 0
        swTangentEdgesVisibleAndFonted = 1
        swTangentEdgesVisible = 2
End Enum

'  Possible values for the swSTLQuality User Preference setting
Public Enum swSTLQuality_e
        swSTLQuality_Coarse = 1
        swSTLQuality_Fine = 2
        swSTLQuality_Custom = 3
End Enum

'  Possible values for the swDrawingProjectionType User Preference setting
Public Enum swDrawingProjectionType_e
        swDrawing1stAngleProjection = 1
        swDrawing3rdAngleProjection = 2
End Enum

Public Enum swPromptAlwaysNever_e
        swResponsePrompt = 0
        swResponseAlways = 1
        swResponseNever = 2
End Enum

'  Possible values for the swIGESRepresentation User Preference setting.
Public Enum swIGESRepresentation_e
        swIGES_TRMSRF = 0       '  Trimmed surface representation
        swIGES_CURVES = 1       '   WireFrame representation
        swIGES_TRMSRFANDCURVES = 2      '   Both trimmed surface and wireframe representation
        swIGES_BREP = 3
End Enum

'  Possible values for the swIGESSystem User Preference setting.
Public Enum swIGESPreferredSystem_e
        swIGES_STANDARD = 0
        swIGES_NURBS = 1
        swIGES_ANSYS = 2
        swIGES_COSMOS = 3
        swIGES_MASCAM = 4
        swIGES_SURFCAM = 5
        swIGES_SMARTCAM = 6
        swIGES_TEKSOFT = 7
        swIGES_ALPHACAM = 8
        swIGES_MULTICAM = 9
        swIGES_ALIAS = 10
End Enum

'  Possible values for the swIGESCurveRepresentation User Preference setting.
Public Enum swIGESCurveRepresentation_e
        swIGES_CURVES_BSPLINE = 0       '  free form curves as bspline representation
        swIGES_CURVES_PSPLINE = 1       '  free form curves as parametric spline representation
End Enum

' Possible value for the constraint status of a sketch
Public Enum swConstrainedStatus_e
        swUnknownConstraint = 1
        swUnderConstrained = 2
        swFullyConstrained = 3
        swOverConstrained = 4
        swNoSolution = 5
        swInvalidSolution = 6
        swAutosolveOff = 7
End Enum

'  Suppression actions for features.
Public Enum swFeatureSuppressionAction_e
        swSuppressFeature = 0   '  Suppress the feature.
        swUnSuppressFeature = 1 '  Unsuppress the feature.
        swUnSuppressDependent = 2       '  Unsuppress the children of the features.
End Enum

'  HLR Quality Settings...
Public Enum swHlrQuality_e
        swPreciseHlr = 0
        swFastHlr = 1
End Enum

'  Possible values for the entityType argument of the Sketch::SetEntityCount API.
Public Enum swSketchEntityType_e
        swSketchEntityPoint = 1
        swSketchEntityLine = 2
        swSketchEntityArc = 3
        swSketchEntityEllipse = 4
        swSketchEntityParabola = 5
        swSketchEntitySpline = 6
End Enum

Public Enum swWzdGeneralHoleTypes_e
        swWzdCounterBore = 0
        swWzdCounterSink = 1
        swWzdHole = 2
        swWzdPipeTap = 3
        swWzdTap = 4
        swWzdLegacy = 5
End Enum

' if additional hole types are added to the hole wizard dialog, this will be incremented.
Public Enum swWzdHoleTypes_e
        swSimple = 0
        swTapered = 1
        swCounterBored = 2
        swCounterSunk = 3
        swCounterDrilled = 4
        swSimpleDrilled = 5
        swTaperedDrilled = 6
        swCounterBoredDrilled = 7
        swCounterSunkDrilled = 8
        swCounterDrilledDrilled = 9
        swCounterBoreBlind = 10
        swCounterBoreBlindCounterSinkMiddle = 11
        swCounterBoreBlindCounterSinkTop = 12
        swCounterBoreBlindCounterSinkTopmiddle = 13
        swCounterBoreThru = 14
        swCounterBoreThruCounterSinkBottom = 15
        swCounterBoreThruCounterSinkMiddle = 16
        swCounterBoreThruCounterSinkMiddleBottom = 17
        swCounterBoreThruCounterSinkTop = 18
        swCounterBoreThruCounterSinkTopBottom = 19
        swCounterBoreThruCounterSinkTopMiddle = 20
        swCounterBoreThruCounterSinkTopMiddleBottom = 21
        swHoleBlind = 22
        swHoleBlindCounterSinkTop = 23
        swCounterSinkBlind = 24
        swHoleThru = 25
        swHoleThruCounterSinkBottom = 26
        swHoleThruCounterSinkTop = 27
        swHoleThruCounterSinkTopBottom = 28
        swCounterSinkThru = 29
        swCounterSinkThruCounterSinkBottom = 30
        swTapBlind = 31
        swTapBlindCounterSinkTop = 32
        swTapThru = 33
        swTapThruCounterSinkBottom = 34
        swTapThruCounterSinkTop = 35
        swTapThruCounterSinkTopBottom = 36
        swPipeTapBlind = 37
        swPipeTapBlindCounterSinkTop = 38
        swPipeTapThru = 39
        swPipeTapThruCounterSinkBottom = 40
        swPipeTapThruCounterSinkTop = 41
        swPipeTapThruCounterSinkTopBottom = 42
        swCounterSinkBlindWithoutHeadClearance = 43
        swCounterSinkThruWithoutHeadClearance = 44
        swCounterSinkThruCounterSinkBottomWithoutHeadClearance = 45
        swTapBlindCosmeticThread = 46
        swTapBlindCosmeticThreadCounterSinkTop = 47
        swTapThruCosmeticThread = 48
        swTapThruCosmeticThreadCounterSinkTop = 49
        swTapThruCosmeticThreadCounterSinkBottom = 50
        swTapThruCosmeticThreadCounterSinkTopBottom = 51
        swTapThruThreadThru = 52
        swTapThruThreadThruCounterSinkTop = 53
        swTapThruThreadThruCounterSinkBottom = 54
        swTapThruThreadThruCountersinkTopBottom = 55
End Enum

'  Update this when you add new hole types.
Public Enum swWzdHoleStandards_e
        swStandardAnsiInch = 0
        swStandardAnsiMetric = 1
        swStandardBSI = 2
        swStandardDME = 3
        swStandardDIN = 4
        swStandardHascoMetric = 5
        swStandardHelicoilInch = 6
        swStandardHelicoilMetric = 7
        swStandardISO = 8
        swStandardJIS = 9
        swStandardPCS = 10
        swStandardProgressive = 11
        swStandardSuperior = 12
End Enum

Public Enum swWzdHoleStandardFastenerTypes_e
        swStandardAnsiInchBinding = 0
        swStandardAnsiInchButton = 1
        swStandardAnsiInchFillister = 2
        swStandardAnsiInchHexBolt = 3
        swStandardAnsiInchHexBoltFinished = 4
        swStandardAnsiInchHexBoltHeavy = 5
        swStandardAnsiInchHexScrew = 6
        swStandardAnsiInchHexWasherScrew = 7
        swStandardAnsiInchPan = 8
        swStandardAnsiInchSocketCapScrew = 9
        swStandardAnsiInchSocketShoulderScrew = 10
        swStandardAnsiInchSquare = 11
        swStandardAnsiInchTruss = 12
        swStandardAnsiInchFlatSocket82 = 13
        swStandardAnsiInchFlatHead100 = 14
        swStandardAnsiInchFlatHead82 = 15
        swStandardAnsiInchOval = 16
        swStandardAnsiInchHcoilTapDrills = 17
        swStandardAnsiInchAllDrillSizes = 18
        swStandardAnsiInchFractionalDrillSizes = 19
        swStandardAnsiInchLetterDrillSizes = 20
        swStandardAnsiInchPipeTapDrills = 21
        swStandardAnsiInchScrewClearances = 22
        swStandardAnsiInchTapDrills = 23
        swStandardAnsiInchNumberDrillSizes = 24
        swStandardAnsiInchTaperedPipeTap = 25
        swStandardAnsiInchBottomingTappedHole = 26
        swStandardAnsiInchTappedHole = 27
        swStandardAnsiMetricButton = 28
        swStandardAnsiMetricHexBolt = 29
        swStandardAnsiMetricHexCapScrew = 30
        swStandardAnsiMetricHexScrewFormed = 31
        swStandardAnsiMetricPan = 32
        swStandardAnsiMetricSocketHeadCapScrew = 33
        swStandardAnsiMetricSocketShoulderScrew = 34
        swStandardAnsiMetricFlatSocket82 = 35
        swStandardAnsiMetricFlatHead82 = 36
        swStandardAnsiMetricOval = 37
        swStandardAnsiMetricHcoilTapDrills = 38
        swStandardAnsiMetricDrillSizes = 39
        swStandardAnsiMetricScrewClearances = 40
        swStandardAnsiMetricTapDrills = 41
        swStandardAnsiMetricBottomingTappedHole = 42
        swStandardAnsiMetricTappedHole = 43
        swStandardBSICheese = 44
        swStandardBSIHexBolt = 45
        swStandardBSIHexCapScrew = 46
        swStandardBSIHexMachineScrew = 47
        swStandardBSIPanHead = 48
        swStandardBSISocketCapScrew = 49
        swStandardBSIFlatSocketCap = 50
        swStandardBSIFlatHead = 51
        swStandardBSIOvalHead = 52
        swStandardBSIHcoilTapDrills = 53
        swStandardBSIDrillSizes = 54
        swStandardBSIScrewClearances = 55
        swStandardBSITapDrills = 56
        swStandardBSITappedHoleBottoming = 57
        swStandardBSITappedHole = 58
        swStandardBSITaperedPipeTap = 59
        swStandardDINHeavyHexBolt = 60
        swStandardDINHexFlangeBolt = 61
        swStandardDINCheeseHead = 62
        swStandardDINHexBolt = 63
        swStandardDINHexCapScrew = 64
        swStandardDINHexMachineScrew = 65
        swStandardDINPan = 66
        swStandardDINSocketHeadCap = 67
        swStandardDINSocketCTSKFlatHead = 68
        swStandardDINCTSKFlatHead = 69
        swStandardDINCTSKRaisedHead = 70
        swStandardDINHcoilTapDrills = 71
        swStandardDINDrillSizes = 72
        swStandardDINScrewClearances = 73
        swStandardDINTapDrills = 74
        swStandardDINTappedHoleBottoming = 75
        swStandardDINTappedHole = 76
        swStandardDINTaperedPipeTap = 77
        swStandardDMECCorePins = 78
        swStandardDMECXCorePins = 79
        swStandardDMETHXEjectorPins = 80
        swStandardDMEStandardLeaderPins = 81
        swStandardDMEReturnPins = 82
        swStandardDMESocketCapScrew = 83
        swStandardDMESupportPillarSHCS = 84
        swStandardDMESpruPullerPins = 85
        swStandardDMEStripperBolt = 86
        swStandardDMEFlatSocket82 = 87
        swStandardDMEFlatHead100 = 88
        swStandardDMEFlatHead82 = 89
        swStandardDMEOval = 90
        swStandardDMESupportPillarClearance = 91
        swStandardDMEFractionalDrillSizes = 92
        swStandardDMEHcoilTapDrills = 93
        swStandardDMEAllDrillSizes = 94
        swStandardDMELetterDrillSizes = 95
        swStandardDMENumberDrillSizes = 96
        swStandardDMEPipeTapDrills = 97
        swStandardDMEScrewClearances = 98
        swStandardDMECCorePinClearances = 99
        swStandardDMECXCorePinClearances = 100
        swStandardDMETHXEjectorPinClearances = 101
        swStandardDMELeaderPinClearances = 102
        swStandardDMEReturnPinClearances = 103
        swStandardDMESpruPullerPinClearances = 104
        swStandardDMETapDrills = 105
        swStandardDMEBottomingTappedHole = 106
        swStandardDMETappedHole = 107
        swStandardDMETaperedPipeTap = 108
        swStandardHascoMetricCCorePins = 109
        swStandardHascoMetricGuideBushings = 110
        swStandardHascoMetricGuidePillars = 111
        swStandardHascoMetricLocatingGuideBushings = 112
        swStandardHascoMetricLocatingGuidePillars = 113
        swStandardHascoMetricSocketCapScrew = 114
        swStandardHascoMetricShoulderScrew = 115
        swStandardHascoMetricCTSKFlatHead = 116
        swStandardHascoMetricDrillSizes = 117
        swStandardHascoMetricScrewClearances = 118
        swStandardHascoMetricCorePinClearances = 119
        swStandardHascoMetricCenteringSleeve = 120
        swStandardHascoMetricEjectorRodClearances = 121
        swStandardHascoMetricBottomingTappedHole = 122
        swStandardHascoMetricTappedHole = 123
        swStandardHcoilInchInsert10Dia = 124
        swStandardHcoilInchInsert15Dia = 125
        swStandardHcoilInchInsert20Dia = 126
        swStandardHcoilInchInsert25Dia = 127
        swStandardHcoilInchInsert30Dia = 128
        swStandardHcoilMetricInsert10Dia = 129
        swStandardHcoilMetricInsert15Dia = 130
        swStandardHcoilMetricInsert20Dia = 131
        swStandardHcoilMetricInsert25Dia = 132
        swStandardHcoilMetricInsert30Dia = 133
        swStandardISOCheeseHead = 134
        swStandardISOHexBolt = 135
        swStandardISOHexCapScrew = 136
        swStandardISOHexMachineScrew = 137
        swStandardISOPan = 138
        swStandardISOSocketHeadCap = 139
        swStandardISOSocketCTSKFlatHead = 140
        swStandardISOCTSKFlatHead = 141
        swStandardISOCTSKRaisedHead = 142
        swStandardISODrillSizes = 143
        swStandardISOScrewClearances = 144
        swStandardISOTapDrills = 145
        swStandardISOTappedHoleBottoming = 146
        swStandardISOTappedHole = 147
        swStandardISOTaperedPipeTap = 148
        swStandardJISCheeseHead = 149
        swStandardJISFillisterHead = 150
        swStandardJISButton = 151
        swStandardJISHexBolt = 152
        swStandardJISHexCapScrew = 153
        swStandardJISHexMachineScrew = 154
        swStandardJISPan = 155
        swStandardJISSocketHeadCap = 156
        swStandardJISSocketShoulderScrew = 157
        swStandardJISFlatCTSKHead = 158
        swStandardJISRaisedCTSKHead = 159
        swStandardJISDrillSizes = 160
        swStandardJISScrewClearances = 161
        swStandardJISTapDrills = 162
        swStandardJISTappedHoleBottoming = 163
        swStandardJISTappedHole = 164
        swStandardJISTaperedPipeTap = 165
        swStandardPCSReturnPins = 166
        swStandardPCSCorePins = 167
        swStandardPCSEjectorPins = 168
        swStandardPCSStandardLeaderPins = 169
        swStandardPCSSocketCapScrew = 170
        swStandardPCSStripperBolt = 171
        swStandardPCSSupportPillarSHCS = 172
        swStandardPCSFlatHead100 = 173
        swStandardPCSFlatHead82 = 174
        swStandardPCSOval = 175
        swStandardPCSFlatSocket82 = 176
        swStandardPCSHcoilTapDrills = 177
        swStandardPCSFractionalDrillSizes = 178
        swStandardPCSNumberDrillSizes = 179
        swStandardPCSPipeTapDrills = 180
        swStandardPCSScrewClearances = 181
        swStandardPCSAllDrillSizes = 182
        swStandardPCSEjectorPinClearances = 183
        swStandardPCSLetterDrillSizes = 184
        swStandardPCSSupportPillarClearances = 185
        swStandardPCSCorePinClearances = 186
        swStandardPCSLeaderPinClearances = 187
        swStandardPCSReturnPinClearances = 188
        swStandardPCSTapDrills = 189
        swStandardPCSBottomingTappedHole = 190
        swStandardPCSTappedHole = 191
        swStandardPCSTaperedPipeTap = 192
        swStandardProgressiveSocketCapScrew = 193
        swStandardProgressiveReturnPins = 194
        swStandardProgressiveCorePins = 195
        swStandardProgressiveEjectorPins = 196
        swStandardProgressiveSpruePullerPins = 197
        swStandardProgressiveSupportPillarSHCS = 198
        swStandardProgressiveStripperBolt = 199
        swStandardProgressiveStandardLeaderPins = 200
        swStandardProgressiveFlatSocket82 = 201
        swStandardProgressiveOval = 202
        swStandardProgressiveFlatHead100 = 203
        swStandardProgressiveFlatHead82 = 204
        swStandardProgressiveHcoilTapDrills = 205
        swStandardProgressiveFractionalDrillSizes = 206
        swStandardProgressiveNumberDrillSizes = 207
        swStandardProgressivePipeTapDrills = 208
        swStandardProgressiveScrewClearances = 209
        swStandardProgressiveAllDrillSizes = 210
        swStandardProgressiveEjectorPinClearances = 211
        swStandardProgressiveLetterDrillSizes = 212
        swStandardProgressiveSupportPillarClearances = 213
        swStandardProgressiveCorePinClearances = 214
        swStandardProgressiveLeaderPinClearances = 215
        swStandardProgressiveSpruePullerPinClearances = 216
        swStandardProgressiveReturnPinClearances = 217
        swStandardProgressiveTapDrills = 218
        swStandardProgressiveTappedHole = 219
        swStandardProgressiveBottomingTappedHole = 220
        swStandardProgressiveTaperedPipeTap = 221
        swStandardSuperiorReturnPins = 222
        swStandardSuperiorEjectorPins = 223
        swStandardSuperiorSpruePullerPins = 224
        swStandardSuperiorSupportPillarSHCS = 225
        swStandardSuperiorStripperBolt = 226
        swStandardSuperiorSocketCapScrew = 227
        swStandardSuperiorStandardLeaderPins = 228
        swStandardSuperiorFlatHead100 = 229
        swStandardSuperiorFlatHead82 = 230
        swStandardSuperiorOval = 231
        swStandardSuperiorFlatSocket82 = 232
        swStandardSuperiorHcoilTapDrills = 233
        swStandardSuperiorFractionalDrillSizes = 234
        swStandardSuperiorNumberDrillSizes = 235
        swStandardSuperiorPipeTapDrills = 236
        swStandardSuperiorScrewClearances = 237
        swStandardSuperiorAllDrillSizes = 238
        swStandardSuperiorEjectorPinClearances = 239
        swStandardSuperiorLetterDrillSizes = 240
        swStandardSuperiorSupportPillarClearances = 241
        swStandardSuperiorLeaderPinClearances = 242
        swStandardSuperiorSpruePullerPinClearances = 243
        swStandardSuperiorReturnPinClearances = 244
        swStandardSuperiorTapDrills = 245
        swStandardSuperiorTappedHole = 246
        swStandardSuperiorBottomingTappedHole = 247
        swStandardSuperiorTaperedPipeTap = 248
End Enum

Public Enum swWzdHoleCounterSinkHeadClearanceTypes_e
        swHeadClearanceIncreasedCsink = 0
        swHeadClearanceAddToCbore = 1
End Enum

Public Enum swWzdHoleHcoilTapTypes_e
        swTapTypePlug = 0
        swTapTypeBottom = 1
End Enum

Public Enum swWzdHoleScrewClearanceTypes_e
        swScrewClearanceClose = 0
        swScrewClearanceNormal = 1
        swScrewClearanceLoose = 2
End Enum

Public Enum swWzdHoleCosmeticThreadTypes_e
        swCosmeticThreadNone = 0
        swCosmeticThreadWithCallout = 1
        swCosmeticThreadWithoutCallout = 2
End Enum

Public Enum swWzdHoleThreadEndCondition_e
        swEndThreadTypeBLIND = 0
        swEndThreadTypeTHROUGH_ALL = 1
        swEndThreadTypeTHROUGH_NEXT = 2
End Enum

Public Enum swCreateFacesBodyAction_e
        swCreateFacesBodyActionCap = 1
        swCreateFacesBodyActionGrow = 2
        swCreateFacesBodyActionGrowFromParent = 3
        swCreateFacesBodyActionLeaveRubber = 4
End Enum

'  Used for APIs that allow bitwise ORing of document types, like
'  SldWorks::AddToolbar2()
Public Enum swDocTemplateTypes_e
        swDocTemplateTypeNONE = &H1
        swDocTemplateTypePART = &H2
        swDocTemplateTypeASSEMBLY = &H4
        swDocTemplateTypeDRAWING = &H8
End Enum

Public Enum swCreateFeatureBodyOpts_e
        swCreateFeatureBodyCheck = &H1
        swCreateFeatureBodySimplify = &H2
End Enum

Public Enum swToolbarDockStatePosition_e
        swDockNoToolbar = -1
        swNoDock = 0
        swDockTop = 1
        swDockBottom = 2
        swDockRight = 3
        swDockLeft = 4
End Enum

Public Enum swImprintingFacesOpts_e
        swImprintingFacesOnTool = &H1
        swImprintingFacesOnOverlapping = &H2
        swImprintingFacesOnExtendFace = &H4
End Enum

'  A list of feature types to be used with the Sketch::CheckFeatureUse API
Public Enum swSketchCheckFeatureProfileUsage_e
        swSketchCheckFeature_UNSET = 0
        swSketchCheckFeature_BASEEXTRUDE = 1
        swSketchCheckFeature_BASEEXTRUDETHIN = 2
        swSketchCheckFeature_BOSSEXTRUDE = 3
        swSketchCheckFeature_BOSSEXTRUDETHIN = 4
        swSketchCheckFeature_SURFACEEXTRUDE = 5
        swSketchCheckFeature_BASEREVOLVE = 6
        swSketchCheckFeature_BASEREVOLVETHIN = 7
        swSketchCheckFeature_BOSSREVOLVE = 8
        swSketchCheckFeature_BOSSREVOLVETHIN = 9
        swSketchCheckFeature_SURFACEREVOLVE = 10
        swSketchCheckFeature_CUTEXTRUDE = 11
        swSketchCheckFeature_CUTEXTRUDETHIN = 12
        swSketchCheckFeature_CUTREVOLVE = 13
        swSketchCheckFeature_CUTREVOLVETHIN = 14
        swSketchCheckFeature_SWEEPSECTION = 15
        swSketchCheckFeature_SURFACESWEEPSECTION = 16
        swSketchCheckFeature_SWEEPPATHORGUIDE = 17
        swSketchCheckFeature_LOFTSECTION = 18
        swSketchCheckFeature_SURFACELOFTSECTION = 19
        swSketchCheckFeature_LOFTGUIDE = 20
        swSketchCheckFeature_RIBSECTION = 21
        swSketchCheckFeature_SHEETMETAL_BASEFLANGE = 22
End Enum

'  A list of return status values for the Sketch::CheckFeatureUse API
Public Enum swSketchCheckFeatureStatus_e
        swSketchCheckFeatureStatus_UnknownError = -1
        swSketchCheckFeatureStatus_OK = 0
        swSketchCheckFeatureStatus_EntXEnt = 1
        swSketchCheckFeatureStatus_EntXSelf = 2
        swSketchCheckFeatureStatus_EntUnspecBad = 3
        swSketchCheckFeatureStatus_ThreeEnts = 4
        swSketchCheckFeatureStatus_EmptySketch = 5
        swSketchCheckFeatureStatus_WrongOpen = 6
        swSketchCheckFeatureStatus_WrongManyContours = 7
        swSketchCheckFeatureStatus_ZeroLengthEnt = 8
        swSketchCheckFeatureStatus_ManyOpen = 9
        swSketchCheckFeatureStatus_NoOpen = 10
        swSketchCheckFeatureStatus_MixedContours = 11
        swSketchCheckFeatureStatus_CturXCtur = 12
        swSketchCheckFeatureStatus_DisjCturs = 13
        swSketchCheckFeatureStatus_OpenWantClosed = 14
        swSketchCheckFeatureStatus_ClosedWantOpen = 15
        swSketchCheckFeatureStatus_DoubleContainment = 16
        swSketchCheckFeatureStatus_MoreThanOneContour = 17
        swSketchCheckFeatureStatus_OneOpenContourExpected = 18
        swSketchCheckFeatureStatus_OneClosedContourExpected = 19
        swSketchCheckFeatureStatus_WantSingleOpenOrMultiClosedDisjoint = 20
        swSketchCheckFeatureStatus_NeedsAxis = 21
        swSketchCheckFeatureStatus_OpenOrUnclear = 22
        swSketchCheckFeatureStatus_ContourIntersectsCenterLine = 23
End Enum

'  A list of return status values for the ModelDoc::GetMassProperties API
Public Enum swMassPropertiesStatus_e
        swMassPropertiesStatus_OK = 0
        swMassPropertiesStatus_UnknownError = 1
        swMassPropertiesStatus_NoBody = 2
End Enum

'  A list of possible arc types when using CreateTangentArc2
Public Enum swTangentArcTypes_e
        swForward = 1
        swLeft = 2
        swBack = 3
        swRight = 4
End Enum

'  Possible values for the options argument of SldWorks::OpenDoc4.
Public Enum swOpenDocOptions_e
        swOpenDocOptions_Silent = &H1   '  Open document silently or not
        swOpenDocOptions_ReadOnly = &H2 '  Open document read only or not
        swOpenDocOptions_ViewOnly = &H4 '  Open document view only or not
        swOpenDocOptions_RapidDraft = &H8       '  Convert document to RapidDraft format or not (drawings only)
        swOpenDocOptions_LoadModel = &H10       '  Load detached models automatically or not (drawings only)
        swOpenDocOptions_AutoMissingConfig = &H20       '  Automatically handle missing configs of drawing views (drawings only)
End Enum

'  Possible values for the options argument of ModelDoc::SaveAs3.
Public Enum swSaveAsOptions_e
        swSaveAsOptions_Silent = &H1    '  Save document silently or not
        swSaveAsOptions_Copy = &H2      '  Save document as a copy or not
        swSaveAsOptions_SaveReferenced = &H4    '  Save referenced documents or not (drawings and parts only)
        swSaveAsOptions_AvoidRebuildOnSave = &H8        '  Avoid rebuild on Save or SaveAs, if swSaveAsOptions_Silent
        swSaveAsOptions_UpdateInactiveViews = &H10      '  Update views of inactive sheets, if swSaveAsOptions_Silent
        swSaveAsOptions_OverrideSaveEmodel = &H20       '  Override system setting for saving emodel data of document
        swSaveAsOptions_SaveEmodelData = &H40   '  If OverrideSaveEmodel is True, use this as the value instead
End Enum

Public Enum swInConfigurationOpts_e
        swConfigPropertySuppressFeatures = 0
        swThisConfiguration = 1
        swAllConfiguration = 2
        swSpecifyConfiguration = 3
End Enum

Public Enum swKernelErrorCode_e
        swErrorSuccess = 1
        swErrorError = 0
        swErrorNotEntity = -100022
        swErrorInvalidParameter = -100120
        swErrorSurfaceDiscontinuous = -100129
        swErrorCurveDiscontinuous = -100131
        swErrorInvalidEntity = -100914
        swErrorInvalidSharing = -100921
        swErrorInvalidKnots = -100978
        swErrorInvalidGeometry = -100999
        swErrorHasInvalidentity = -101004
        swErrorBodyDontKnit = -101041
        swErrorInvalidPattern = -101042
        swErrorCurveShort = -101057
        swErrorFailed = -101063
        swErrorCheckFailed = -105061
        swErrorGeometryMissing = -113803
        swErrorTopologySelfx = -113804
        swErrorGeometrySelfx = -113805
        swErrorGeometryDegenerate = -113806
        swErrorInvalidGeometry2 = -113808
        swErrorCheckFailed2 = -113812
        swErrorFaceFaceInconsistent = -113816
        swErrorVertexNotOnCurve = -113818
        swErrorVerticesTouch = -113821
        swErrorLoopsInconsistent = -113826
        swErrorGeometryDiscontinuous = -113827
        swErrorFacecheckFailed = -113829
        swErrorFaceRedundant = -116402
        swErrorInconsistentDirs = -116403
        swErrorEdgeisectInvalid = -116404
        swErrorInvalidLoop = -116405
        swErrorEdgeIncorrectOrder = -116406
        swErrorUnknown = -1
End Enum

'  Different buttons that can be displayed on the PropertyManagerPage.
Public Enum swPropertyManagerButtonTypes_e
        swPropertyManager_OkayButton = &H1
        swPropertyManager_CancelButton = &H2
        swPropertyManager_HelpButton = &H4
End Enum

'  Return status values for the various PropertyManagerPage APIs.
Public Enum swPropertyManagerStatus_e
        swPropertyManagerStatus_Okay = 0
        swPropertyManagerStatus_Failed = -1
        swPropertyManagerStatus_Disconnected = -2
End Enum

'  Possible values for the swParasolidOutputVersion User Preference setting.
Public Enum swParasolidOutputVersion_e
        swParasolidOutputVersion_latest = 0
        swParasolidOutputVersion_80 = 1
        swParasolidOutputVersion_90 = 2
        swParasolidOutputVersion_91 = 3
        swParasolidOutputVersion_100 = 4
        swParasolidOutputVersion_110 = 5
        swParasolidOutputVersion_111 = 6
        swParasolidOutputVersion_120 = 7
        swParasolidOutputVersion_121 = 8
End Enum

'  Possible values for what action to take when setting the selected object mark.
Public Enum swSelectionMarkAction_e
        swSelectionMarkSet = 0
        swSelectionMarkAppend = 1
        swSelectionMarkRemove = 2
        swSelectionMarkClear = 3
End Enum

'  Possible values for the swEdgesHiddenEdgeDisplay integer user preference option
Public Enum swEdgesHiddenEdgeDisplay_e
        swEdgesHiddenEdgeDisplaySolid = 1
        swEdgesHiddenEdgeDisplayDashed = 2
End Enum

'  Possible values for the swEdgesTangentEdgeDisplay integer user preference option
Public Enum swEdgesTangentEdgeDisplay_e
        swEdgesTangentEdgeDisplayVisible = 1
        swEdgesTangentEdgeDisplayPhantom = 2
        swEdgesTangentEdgeDisplayRemoved = 3
End Enum

'  Possible values for the swEdgesShadedModeDisplay integer user preference option
Public Enum swEdgesShadedModeDisplay_e
        swEdgesShadedModeDisplayNone = 1
        swEdgesShadedModeDisplayHLR = 2
        swEdgesShadedModeDisplayWireframe = 3
End Enum

Public Enum swSplitFaceOnParam_e
        swSplitFaceOnParamU = 1
        swSplitFaceOnParamV = 2
End Enum

Public Enum swTbCommand_e
        swTbCONTROL = -2
        swTbACTIVE = -1
        swTbNONE = 0
        swTbPART = 1
        swTbASSEMBLY = 2
        swTbDRAWING = 3
End Enum

Public Enum swTbSaveModes_e
        swTbSAVE = 0
        swTbLOAD = 1
End Enum

Public Enum swTbControlModes_e
        swTbSTOP = 0
        swTbCONTINUE = 1
        swTbOleInplaceMode = 2
End Enum

Public Enum swBendAllowanceTypes_e
        swBendAllowanceBendTable = 1
        swBendAllowanceKFactor = 2
        swBendAllowanceDirect = 3
        swBendAllowanceDeduction = 4
End Enum

Public Enum swSheetMetalReliefTypes_e
        swSheetMetalReliefRectangular = 1
        swSheetMetalReliefTear = 2
        swSheetMetalReliefObround = 3
        swSheetMetalReliefNone = 4
        swSheetMetalReliefTearBend = 5
End Enum

Public Enum swUserUnitsType_e
        swLengthUnit = 0
        swAngleUnit = 1
End Enum

Public Enum swFlangeOffsetTypes_e
        swFlangeOffsetBlind = 1
        swFlangeOffsetUpToVertex = 2
        swFlangeOffsetUpToSurface = 3
        swFlangeOffsetFromSurface = 4
        swFlangeOffsetMidPlane = 5
End Enum

Public Enum swFlangeDimTypes_e
        swFlangeDimTypeOuterVirtualSharp = 1
        swFlangeDimTypeInnerVirtualSharp = 2
End Enum

Public Enum swFlangePositionTypes_e
        swFlangePositionTypeMaterialInside = 1
        swFlangePositionTypeMaterialOutside = 2
        swFlangePositionTypeBendOutside = 3
        swFlangePositionTypeBendCenterLine = 4
        swFlangePositionTypeBendSharp = 5
End Enum

Public Enum swReliefTearTypes_e
        swReliefTearTypeRip = 1
        swReliefTearTypeExtend = 2
End Enum

Public Enum swClosedCornerTypes_e
        swClosedCornerTypeButt = 1
        swClosedCornerTypeOverlap = 2
        swClosedCornerTypeUnderlap = 3
End Enum

Public Enum swSelectionReferenceTypes_e
        swReferenceTypeVertex = 1
        swReferenceTypeEdge = 2
        swReferenceTypeFace = 3
        swReferenceTypeRefSurface = 4
        swReferenceTypeRefPlan = 5
        swReferenceTypeSketchPoint = 6
        swReferenceTypeBody = 7
End Enum

Public Enum swPatternReferenceTypes_e
        swPatternReferenceTypeAxis = 0
        swPatternReferenceTypeEdge = 1
        swPatternReferenceTypeRefDim = 2
End Enum

Public Enum swThinWallType_e
        swThinWallOneDirection = 0
        swThinWallOppDirection = 1
        swThinWallMidPlane = 2
        swThinWallTwoDirection = 3
End Enum

Public Enum swTextInBoxStyle_e
        swTextInBoxStyleNone = 0
        swTextInBoxStyleWrap = 1
        swTextInBoxStyleFit = 2
End Enum

'  Possible values for the swPageSetupPrinterOrientation integer user preference value. (SPR 100576)
Public Enum swPageSetupOrientation_e
        swPageSetupOrient_Portrait = 1
        swPageSetupOrient_Landscape = 2
End Enum

'  Possible values for the swPageSetupPrinterDrawingColor integer user preference value. (SPR 100576)
Public Enum swPageSetupDrawingColor_e
        swPageSetup_AutomaticDrawingColor = 1
        swPageSetup_ColorGrey = 2
        swPageSetup_BlackAndWhite = 3
End Enum

'  PropertyManagerPage2 status codes.
Public Enum swPropertyManagerPageStatus_e
        swPropertyManagerPage_Okay = 0
        swPropertyManagerPage_UnsupportedHandler = 1
        swPropertyManagerPage_CreationFailure = -1
        swPropertyManagerPage_NoDocument = -2
End Enum

'  Possible values for the options argument of SldWorks::CreatePropertyManagerPage.
Public Enum swPropertyManagerPageOptions_e
        swPropertyManagerOptions_OkayButton = &H1
        swPropertyManagerOptions_CancelButton = &H2
        swPropertyManagerOptions_LockedPage = &H4
        swPropertyManagerOptions_CloseDialogButton = &H8
        swPropertyManagerOptions_MultiplePages = &H10
End Enum

'  Possible values for the options argument of PropertyManagerPage2::AddGroupBox.
Public Enum swAddGroupBoxOptions_e
        swGroupBoxOptions_Checkbox = &H1
        swGroupBoxOptions_Checked = &H2
        swGroupBoxOptions_Visible = &H4
        swGroupBoxOptions_Expanded = &H8
End Enum

'  Possible values for the message box visibility argument of PropertyManagerPage2::SetMessage.
Public Enum swPropertyManagerPageMessageVisibility
        swNoMessageBox = 1
        swMessageBoxHidden = 2
        swMessageBoxVisible = 3
        swImportantMessageBox = 4
End Enum

'  Possible values for the control type argument of PropertyManagerPageGroup::AddControl.
Public Enum swPropertyManagerPageControlType_e
        swControlType_Label = 1
        swControlType_Checkbox = 2
        swControlType_Button = 3
        swControlType_Option = 4
        swControlType_Textbox = 5
        swControlType_Listbox = 6
        swControlType_Combobox = 7
        swControlType_Numberbox = 8
        swControlType_Selectionbox = 9
        swControlType_ActiveX = 10
End Enum

'  Possible values for the options argument of PropertyManagerPageGroup::AddControl.
Public Enum swAddControlOptions_e
        swControlOptions_Visible = &H1
        swControlOptions_Enabled = &H2
        swControlOptions_SmallGapAbove = &H4
End Enum

'  Possible values for the left edge alignment argument of PropertyManagerPageGroup::AddControl.
Public Enum swPropertyManagerPageControlLeftAlign_e
        swControlAlign_LeftEdge = 1
        swControlAlign_Indent = 2
End Enum

'  Possible values for the default unit type of a PropertyManagerPageNumberbox.
Public Enum swNumberboxUnitType_e
        swNumberBox_UnitlessInteger = 1
        swNumberBox_UnitlessDouble = 2
        swNumberBox_Length = 3
        swNumberBox_Angle = 4
End Enum

'  Possible values for predefined bitmap label types for a PropertyManagerPage control.
Public Enum swControlBitmapLabelType_e
        swBitmapLabel_LinearDistance = 1
        swBitmapLabel_AngularDistance = 2
        swBitmapLabel_SelectEdgeFaceVertex = 3
        swBitmapLabel_SelectFaceSurface = 4
        swBitmapLabel_SelectVertex = 5
        swBitmapLabel_SelectFace = 6
        swBitmapLabel_SelectEdge = 7
        swBitmapLabel_SelectFaceEdge = 8
        swBitmapLabel_SelectComponent = 9
End Enum

'  Possible values for the argument in the OnClose handler method call.
Public Enum swPropertyManagerPageCloseReasons_e
        swPropertyManagerPageClose_UnknownReason = 0
        swPropertyManagerPageClose_Okay = 1
        swPropertyManagerPageClose_Cancel = 2
End Enum

'  Possible values in the PropertyManagerPageListbox::Style property.
Public Enum swPropMgrPageListBoxStyle_e
        swPropMgrPageListBoxStyle_Sorted = &H1
End Enum

'  Possible values in the PropertyManagerPageCombobox::Style property.
Public Enum swPropMgrPageComboBoxStyle_e
        swPropMgrPageComboBoxStyle_Sorted = &H1
        swPropMgrPageComboBoxStyle_EditableText = &H2
End Enum

'  Possible values for the button type argument of PropertyManagerPage2::EnableButton.
Public Enum swPropertyManagerPageButtons_e
        swPropertyManagerPageButton_Ok = 1
        swPropertyManagerPageButton_Cancel = 2
        swPropertyManagerPageButton_Help = 3
        swPropertyManagerPageButton_Next = 4
        swPropertyManagerPageButton_Back = 5
End Enum

'  Possible values in the PropertyManagerPageLabel::Style property.
Public Enum swPropMgrPageLabelStyle_e
        swPropMgrPageLabelStyle_LeftText = &H1
        swPropMgrPageLabelStyle_CenterText = &H2
        swPropMgrPageLabelStyle_RightText = &H4
End Enum

'  Possible return values for the OnActiveXControlCreated PropertyManagerPage2 handler method.
Public Enum swHandleActiveXCreationFailure_e
        swHandleActiveXCreationFailure_Cancel = 1
        swHandleActiveXCreationFailure_Retry = 2
        swHandleActiveXCreationFailure_Continue = 3
End Enum

'  Possible values in the PropertyManagerPageOption::Style property.
Public Enum swPropMgrPageOptionStyle_e
        swPropMgrPageOptionStyle_FirstInGroup = &H1
End Enum

'  Possible values in the PropertyManagerPageSelectionbox::Style property.
Public Enum swPropMgrPageSelectionBoxStyle_e
        swPropMgrPageSelectionBoxStyle_HScroll = &H1
End Enum

Public Enum swInsertEdgeFlangeOptions_e
        swInsertEdgeFlangeUseDefaultRadius = &H1
        swInsertEdgeFlangeFlipDir = &H2
        swInsertEdgeFlangeDoOffset = &H4
        swInsertEdgeFlangeReverseOffsetDir = &H8
        swInsertEdgeFlangeTearClip = &H10
        swInsertEdgeFlangeTrimSideBend = &H20
        swInsertEdgeFlangeUseReliefRatio = &H40
        swInsertEdgeFlangeUseDefaultRelief = &H80
End Enum

' Twist control type used for creating Sweep
Public Enum swTwistControlType_e
        swTwistControlFollowPath = 0
        swTwistControlKeepNormalConstant = 1
        swTwistControlFollowPathFirstGuideCurve = 2
        swTwistControlFollowFirstSecondGuideCurves = 3
End Enum

'  Sweep and Loft tangency options
Public Enum swTangencyType_e
        swTangencyNone = 0
        swTangencyNormalToProfile = 1
        swTangencyDirectionVector = 2
        swTangencyAllFaces = 3
End Enum

' Step type for Step Draft
Public Enum swDraftStepType_e
        swDraftTaperedStep = 3
        swDraftPerpendicular = 6
End Enum

' Face propagation type in draft
Public Enum swDraftFacePropagationType_e
        swFacePropNone = 0
        swFacePropTangent = 1
        swFacePropAllLoops = 2
        swFacePropInnerLoops = 3
        swFacePropOuterLoops = 4
End Enum

' thickness type in Thicken
Public Enum swThickenThicknessType_e
        swThickenSideOne = 0
        swThickenSideTwo = 1
        swThickenSideBoth = 2
End Enum

' External reference status
Public Enum swExternalReferenceStatus_e
        swExternalReferenceBroken = 0
        swExternalReferenceLocked = 1
        swExternalReferenceInContext = 3
        swExternalReferenceOutOfContext = 4
        swExternalReferenceDangling = 5
End Enum

Public Enum swReplaceComponentError_e
        swReplaceComponent_Undefined = 0
        swReplaceComponent_Success = 1
        swReplaceComponent_EmptyName = 2
        swReplaceComponent_InvalidFileName = 3
        swReplaceComponent_SameModelDifferentPath = 4
        swReplaceComponent_SameFile = 5
        swReplaceComponent_NotTopLevelComponent = 6
        swReplaceComponent_UnknownError = 7
End Enum

Public Enum swInContextEditTransparencyType_e
        swInContextEditTransparencyOpaque = 0
        swInContextEditTransparencyForce = 1
        swInContextEditTransparencyMaintain = 2
End Enum

Public Enum swDraftType_e
        swNeutralPlaneDraft = 0
        swPartingLineDraft = 1
        swStepDraft = 3
End Enum

Public Enum swMacroFeatureParamType_e
        swMacroFeatureParamTypeString = 0
        swMacroFeatureParamTypeDouble = 1
        swMacroFeatureParamTypeInteger = 2
End Enum

'  Possible values for the swDetailingDatumDisplayType document user preference.
Public Enum swDatumDisplayType_e
        swDatumDisplayType_Default = 0
        swDatumDisplayType_Square = 1
        swDatumDisplayType_Roundgb = 2
End Enum

Public Enum swCurveDrivenPatternCurveMethod_e
        swCurvePatternTransformCurve = 0
        swCurvePatternOffsetCurve = 1
End Enum

Public Enum swCurveDrivenPatternAlignment_e
        swCurvePatternTangentToCurve = 0
        swCurvePatternAlignToSeed = 1
End Enum

'  Possible values for the InsertBomTable error return argument.
Public Enum swBOMConfigurationCreationErrors_e
        swBOMTableCreation_Okay = 0
        swBOMTableCreation_UnspecifiedError = -1
        swBOMTableCreation_MustBeDrawingView = -2
        swBOMTableCreation_AlreadyExists = -3
        swBOMTableCreation_ExcelDisabled = -4
        swBOMTableCreation_Failed = -5
        swBOMTableCreation_NoModelForView = -6
End Enum

'  Possible values for the swBOMConfigurationAnchorType environment user preference.
Public Enum swBOMConfigurationAnchorType_e
        swBOMConfigurationAnchor_TopLeft = 1
        swBOMConfigurationAnchor_TopRight = 2
        swBOMConfigurationAnchor_BottomLeft = 3
        swBOMConfigurationAnchor_BottomRight = 4
End Enum

'  Possible values for the swBOMConfigurationWhatToShow environment user preference.
Public Enum swBOMConfigurationWhatToShow_e
        swBOMConfiguration_ShowPartsOnly = 1
        swBOMConfiguration_ShowPartsAndTopLevelAsm = 2
        swBOMConfiguration_ShowAllInIndentedList = 3
End Enum

'  Possible values for the swBOMControlMissingRowDisplay environment user preference.
Public Enum swBOMControlMissingRowDisplay_e
        swBOMControlShowMissingRow = 1
        swBOMControlHideMissingRow = 2
        swBOMControlStrikeMissingRow = 3
End Enum

'  Possible values for the swBOMControlSplitDirection environment user preference.
Public Enum swBOMControlSplitDirection_e
        swBOMControlSplitRight = 1
        swBOMControlSplitLeft = 2
End Enum

'  Possible values for the Options argument of AddConfiguration3 and EditConfiguration3.
Public Enum swConfigurationOptions2_e
        swConfigOption_UseAlternateName = &H1
        swConfigOption_DontShowPartsInBOM = &H2
        swConfigOption_SuppressByDefault = &H4
        swConfigOption_HideByDefault = &H8
        swConfigOption_MinFeatureManager = &H10
        swConfigOption_InheritProperties = &H20
End Enum

'  Possible values for the stacked balloon direction.
Public Enum swStackedBalloonDirection_e
        swStackedBalloonDir_None = 0
        swStackedBalloonDir_Up = 1
        swStackedBalloonDir_Down = 2
        swStackedBalloonDir_Left = 3
        swStackedBalloonDir_Right = 4
End Enum

'  Possible values for the draft analysis
Public Enum swDraftAnalysisOptions_e
        swDraftAnalysisFlipDir = &H1
        swDraftAnalysisFindSteep = &H2
End Enum

Public Enum swDraftAnalysisShow_e
        swDraftAnalysisShowPositive = &H1
        swDraftAnalysisShowNegative = &H2
        swDraftAnalysisShowDraftRequired = &H4
        swDraftAnalysisShowStraddle = &H8
        swDraftAnalysisShowPositiveSteep = &H10
        swDraftAnalysisShowNegativeSteep = &H20
        swDraftAnalysisShowSurface = &H40
End Enum

'  Possible values for getting the standard strings that can be part of a page header or footer.
Public Enum swStandardHeaderFooterPageSetupTexts_e
        swHeaderFooterText_PageNumber = 1
        swHeaderFooterText_PageCount = 2
        swHeaderFooterText_Date = 3
        swHeaderFooterText_Time = 4
        swHeaderFooterText_Filename = 5
End Enum

Public Enum swDetailingChamferDimLeaderTextStyle_e
        swDetailChamferDimDistDist = 1
        swDetailChamferDimDistAng = 2
        swDetailChamferDimAngDist = 3
        swDetailChamferDimCDist = 4
End Enum

Public Enum swDetailingChamferDimLeaderStyle_e
        swDetailChamferDimLeaderHorizBeside = 1
        swDetailChamferDimLeaderHorizAbove = 2
        swDetailChamferDimLeaderAngBeside = 3
        swDetailChamferDimLeaderAngAbove = 4
End Enum

Public Enum swDetailingChamferDimXStyle_e
        swDetailingChamferDimXStyleUpperCaseX = 1
        swDetailingChamferDimXStyleLowerCaseX = 2
End Enum

Public Enum swHemPositionTypes_e
        swHemPositionTypeInside = 0
        swHemPositionTypeOutside = 1
End Enum

Public Enum swHemTypes_e
        swHemTypeOpen = 0
        swHemTypeClosed = 1
        swHemTypeTearDrop = 2
        swHemTypeRolled = 3
        swHemTypeDouble = 4
End Enum

Public Enum swBreakCornerTypes_e
        swBreakCornerTypeFillet = 1
        swBreakCornerTypeChamfer = 2
End Enum

Public Enum swJogDimensionPositionType_e
        swJogDimensionPositionInsideOffset = 1
        swJogDimensionPositionOutsideOffset = 2
        swJogDimensionPositionOverallPosition = 3
End Enum

Public Enum swJogPositionType_e
        swJogPositionBendCenterline = 1
        swJogPositionMaterialInside = 2
        swJogPositionMaterialOutside = 3
        swJogPositionBendOutside = 4
End Enum

Public Enum swJogOffsetTypes_e
        swJogOffsetBlind = 1
        swJogOffsetUpToVertex = 2
        swJogOffsetUpToSurface = 3
        swJogOffsetFromSurface = 4
        swJogOffsetMidPlane = 5
End Enum

Public Enum swSurfaceTrimType_e
        swTypeTrimTool = 0
        swTypeMutualTrim = 1
End Enum

Public Enum swRevolveType_e
        swRevolveTypeOneDirection = 0
        swRevolveTypeMidPlane = 1
        swRevolveTypeTwoDirection = 2
End Enum

Public Enum swSurfaceExtendEndCond_e
        swSurfaceExtendEndCondDistance = 0
        swSurfaceExtendEndCondUpToPoint = 1
        swSurfaceExtendEndCondUpToSurface = 2
End Enum

'  Possible values for callout target style.
Public Enum swCalloutTargetStyle_e
        swCalloutTargetStyle_None = 0
        swCalloutTargetStyle_Square = 1
        swCalloutTargetStyle_Circle = 2
        swCalloutTargetStyle_Triangle = 3
        swCalloutTargetStyle_Arrow = 4
End Enum

Public Enum swBendType_e
        swSharpBend = 0
        swRoundBend = 1
        swFlatBend = 2
        swNoneBend = 3
        swBaseBend = 4
        swMiterBend = 5
        swFlat3dBend = 6
        swMirrorBend = 7
        swEdgeFlangeBend = 8
        swHemBend = 9
        swFreeFormBend = 10     '  Obsolete; this type was never actually used by Sw
        swRuledBend = 11        '  Obsolete; this type was never actually used by Sw
        swLoftedBend = 12
End Enum

'  Possible values for the block definition external reference state, returned by the
'  BlockDefinition::SetUseExternalFile and SetExternalFileName APIs.
Public Enum swBlockDefinitionExtFileStatus_e
        swBlockDefinitionExtFile_Failed = -1
        swBlockDefinitionExtFile_Success = 0
        swBlockDefinitionExtFile_NotLinked = 1
        swBlockDefinitionExtFile_MissingReference = 2
        swBlockDefinitionExtFile_OutOfDateReference = 3
End Enum

'  Project 4545 XHatch Export to DXF-DWG
Public Enum swCrossHatchFilter_e
        swCrossHatchInclude = 0
        swCrossHatchExclude = 1
        swCrossHatchOnly = 2
End Enum

'  Possible values for the swLargeAsmModeCheckOutOfDateLightweight user preference.
Public Enum swCheckOutOfDate_e
        swCheckOutOfDate_DoNotCheck = 0
        swCheckOutOfDate_Indicate = 1
        swCheckOutOfDate_AlwaysResolve = 2
End Enum

'  Possible values for the break line style.  Used by BreakLine::Style property.
Public Enum swBreakLineStyle_e
        swBreakLine_Straight = 1
        swBreakLine_ZigZag = 2
        swBreakLine_Curve = 3
        swBreakLine_SmallZigZag = 4
End Enum

'  Possible values for the break line orientation.  Used by BreakLine::Orientation property.
Public Enum swBreakLineOrientation_e
        swBreakLineHorizontal = 1
        swBreakLineVertical = 2
End Enum

'  Possible values for the swSaveAssemblyAsPartOptions user preference.
Public Enum swSaveAsmAsPartOptions_e
        swSaveAsmAsPart_AllComponents = 1
        swSaveAsmAsPart_ExteriorComponents = 2
        swSaveAsmAsPart_ExteriorFaces = 3
End Enum

'  Possible values for the swPrompForFilenameCause_e user preference.
Public Enum swPrompForFilenameCause_e
        swUnused = 0
        swGeneric = 1
        swMirrorComponent = 2
        swWeldBead = 3
        swDerivedPart = 4
        swSplitAssembly = 5
        swSplitPart = 6
        swInsertEnvelopeFromFile = 7
        swMirrorComponentBrowse = 8
        swCreateNamedViewFromFile = 9
        swComponentPropsReplace = 10
        swOpenAssociatedDrawing = 11
        swFileReloadReplace = 12
        swDrawingAddViewFromFile = 13
        swDrawingInsert3ViewFromFile = 14
        swAddComponent = 15
End Enum

Public Enum swRefPlaneType_e
        swRefPlaneInvalid = 0
        swRefPlaneUndefined = 1
        swRefPlaneLinePoint = 2
        swRefPlaneThreePoint = 3
        swRefPlaneLineLine = 4
        swRefPlaneDistance = 5
        swRefPlaneParallel = 6
        swRefPlaneAngle = 7
        swRefPlaneNormal = 8
        swRefPlaneOnSurface = 9
End Enum

Public Enum swOnSurfacePlaneProjectType_e
        swOnSurfacePlaneProjecttoNearestLocation = 0
        swOnSurfacePlaneProjectAlongSketchNormal = 1
End Enum

Public Enum swFileSaveTypes_e
        swFileSave = 1
        swFileSaveAs = 2
        swFileSaveAsCopy = 3
End Enum

'  Possible values for the center mark style.  Used by CenterMark::Style property and
'  DrawingDoc::InsertCenterMark.
Public Enum swCenterMarkStyle_e
        swCenterMark_NonAnnotation = 1
        swCenterMark_Single = 2
        swCenterMark_LinearGroup = 3
        swCenterMark_CircularGroup = 4
End Enum

'  Possible values for the center mark connection line visibility.
'  Used by CenterMark::ConnectionLines property.
Public Enum swCenterMarkConnectionLine_e
        swCenterMark_ShowNoConnectLines = &H0
        swCenterMark_ShowLinearConnectLines = &H1       '  applies only to linear pattern style
        swCenterMark_ShowCircularConnectLines = &H2     '  applies only to circular pattern style
        swCenterMark_ShowRadialConnectLines = &H4       '  applies only to circular pattern style
        swCenterMark_ShowBaseCenterMarkLines = &H8      '  applies only to circular pattern style
End Enum

Public Enum swTextAlignmentVertical_e
        swTextAlignmentTop = 0
        swTextAlignmentMiddle = 1
        swTextAlignmentBottom = 2
End Enum

Public Enum swMacroFeatureEntityIdType_e
        swMacroFeatureEntityIdNotApplied = -1   '  id is not owned by macro feature or entity does not require id assignment.
        swMacroFeatureEntityIdUndefined = 0     '  no id on entity
        swMacroFeatureEntityIdUnique = 1        '  id was uniquely assigned
        swMacroFeatureEntityIdDerived = 2       '  id was derived from parent feature id
        swMacroFeatureEntityIdUserDefined = 3   '  id was assigned by user
End Enum

Public Enum swCornerReliefType_e
        swCornerCircularRelief = 0
        swCornerSquareRelief = 1
        swCornerBendWaistRelief = 2
End Enum

'  Possible values for the EditRollback3 API to use for indicating where the rollback bar is going.
Public Enum swMoveRollbackBarTo_e
        swMoveRollbackBarToEnd = 1
        swMoveRollbackBarToPreviousPosition = 2
        swMoveRollbackBarToBeforeFeature = 3
        swMoveRollbackBarToAfterFeature = 4
End Enum

'  Possible values for the block instance text status.  Used by BlockInstance::TextDisplay property.
Public Enum swBlockInstanceTextDisplay_e
        swBlockInstanceTextDisplayNone = 1
        swBlockInstanceTextDisplayAll = 2
        swBlockInstanceTextDisplayNormal = 3
End Enum

Public Enum swTranslationNotifyOptions_e
        swTranslationNotifySilentMode = &H1
End Enum

Public Enum swMacroFeatureOptions_e
        swMacroFeatureByDefault = &H0
        swMacroFeatureAlwaysAtEnd = &H1
End Enum

Public Enum swMacroFeatureSecurityOptions_e
        swMacroFeatureSecurityByDefault = &H0
        swMacroFeatureSecurityCannotBeDeleted = &H1
        swMacroFeatureSecurityNotEditable = &H2
        swMacroFeatureSecurityCannotBeSuppressed = &H4
End Enum

'  Possible values for the which page setup object to use.  Used by ModelDoc::UsePageSetup property.
Public Enum swPageSetupInUse_e
        swPageSetupInUse_Application = 1
        swPageSetupInUse_Document = 2
        swPageSetupInUse_DrawingSheet = 3
End Enum

Public Enum swAutodimEntities_e
        swAutodimEntitiesAll = 1
        swAutodimEntitiesSelected = 2
End Enum

Public Enum swAutodimMark_e
        swAutodimMarkEntities = &H1
        swAutodimMarkHorizontalDatum = &H2
        swAutodimMarkVerticalDatum = &H4
End Enum

Public Enum swAutodimScheme_e
        swAutodimSchemeBaseline = 1
        swAutodimSchemeOrdinate = 2
        swAutodimSchemeChain = 3
        swAutodimSchemeCenterline = 4
End Enum

Public Enum swAutodimHorizontalPlacement_e
        swAutodimHorizontalPlacementBelow = -1
        swAutodimHorizontalPlacementAbove = 1
End Enum

Public Enum swAutodimVerticalPlacement_e
        swAutodimVerticalPlacementLeft = -1
        swAutodimVerticalPlacementRight = 1
End Enum

Public Enum swAutodimStatus_e
        swAutodimStatusSuccess = 0
        swAutodimStatusBadOptionValue = 1
        swAutodimStatusNoActiveDoc = 2
        swAutodimStatusDocTypeNotSupported = 3
        swAutodimStatusNoActiveSketch = 4
        swAutodimStatus3DSketchNotSupported = 5
        swAutodimStatusSketchIsEmpty = 6
        swAutodimStatusSketchIsOverDefined = 7
        swAutodimStatusNoEntities = 8
        swAutodimStatusEntitiesNotValid = 9
        swAutodimStatusCenterlineNotAllowed = 10
        swAutodimStatusDatumNotSupplied = 11
        swAutodimStatusDatumNotUnique = 12
        swAutodimStatusDatumNotValidType = 13
        swAutodimStatusDatumLineNotCenterline = 14
        swAutodimStatusDatumLineNotVertical = 15
        swAutodimStatusDatumLineNotHorizontal = 16
        swAutodimStatusAlgorithmFailed = 17
End Enum

Public Enum swFeatureFilletOptions_e
        swFeatureFilletPropagate = &H1
        swFeatureFilletUniformRadius = &H2
        swFeatureFilletVarRadiusType = &H4
        swFeatureFilletUseHelpPoint = &H8
        swFeatureFilletUseTangentHoldLine = &H10
        swFeatureFilletCornerType = &H20
        swFeatureFilletAttachEdges = &H40
End Enum

Public Enum swRibExtrusionDirection_e
        swRibParallelToSketch = 0
        swRibNormalToSketch = 1
End Enum

Public Enum swRibType_e
        swRibLinear = 0
        swRibNatural = 1
End Enum

Public Enum swSimpleFilletType_e
        swConstRadiusFillet = 0
        swFaceFillet = 2
        swFullRoundFillet = 3
End Enum

Public Enum swSimpleFilletWhichFaces_e
        swSimpleFilletSingleRadius = 0
        swFaceFilletSet1 = 1
        swFaceFilletSet2 = 2
        swFullRoundFilletSet1 = 3
        swFullRoundFilletCenterSet = 4
        swFullRoundFilletSet2 = 5
End Enum

Public Enum swHelixDefinedBy_e
        swHelixDefinedByPitchAndRevolution = 0
        swHelixDefinedByHeightAndRevolution = 1
        swHelixDefinedByHeightAndPitch = 2
        swHelixDefinedBySpiral = 3
End Enum

Public Enum swCreateWireBodyOptions_e
        swCreateWireBodyByDefault = 0
        swCreateWireBodyMergeCurves = &H1
End Enum

Public Enum swBoundingBoxOptions_e
        swBoundingBoxIncludeRefPlanes = &H1
        swBoundingBoxIncludeSketches = &H2
End Enum

' Used by Dimension::GetType
Public Enum swDimensionParamType_e
        swDimensionParamTypeUnknown = -1
        swDimensionParamTypeDoubleLinear = 0
        swDimensionParamTypeDoubleAngular = 1
        swDimensionParamTypeInteger = 2
End Enum

'
'  To combine options for a specific API call, you can use bitwise addition
'  on the numbers within a particular enum.
'
'  The following enum represents the option bits that can be set
'  for FeatureRevolve2 and FeatureCutRevolve2
Public Enum swRevolveOptions_e
        swAutoCloseSketch = &H1
End Enum

'  The following enum represents the option bits that can be set
'  for AddConfiguration and EditConfiguration
Public Enum swConfigurationOptions_e
        swUseAlternateName = &H1
        swDontShowPartsInBOM = &H2
End Enum

'  The following enum represents the option bits that can be set
'  for SetBlockingState
Public Enum swBlockingStates_e
        swNoBlock = &H0
        swFullBlock = &H1
        swModifyBlock = &H2
        swPartialModifyBlock = &H3
End Enum

'  The following enum can be used with the ModelDoc::Rebuild function.  Be aware that
'  certain options are only valid for particular document types.  For example,
'  swUpdateMates is only valid for ModelDoc objects which are assemblies.
Public Enum swRebuildOptions_e
        swRebuildAll = &H1
        swForceRebuildAll = &H2
        swUpdateMates = &H4
        swCurrentSheetDisp = &H8
        swUpdateDirtyOnly = &H10
End Enum
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).