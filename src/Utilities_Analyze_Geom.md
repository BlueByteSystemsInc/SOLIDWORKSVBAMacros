# Geometry Analysis Macro for SolidWorks

## Description
This macro leverages SolidWorks Utilities to perform an extensive geometry analysis on the active document. It identifies common geometric issues such as short edges, knife edges, discontinuous edges, small faces, sliver faces, discontinuous faces, and knife vertices, which can impact the manufacturability and integrity of the design.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - A SolidWorks part or assembly must be open.
> - SolidWorks Utilities add-in must be enabled to access the geometry analysis tools.

## Results
> [!NOTE]
> - Identifies and lists various geometric defects.
> - Optional: Saves a detailed report to a specified directory.
> - Updates the Design Checker with the results if necessary.

## Steps to Setup the Macro

1. **Prepare SolidWorks**:
   - Ensure SolidWorks is running and that a document is open.
   - Enable the SolidWorks Utilities add-in via **Tools** > **Add-Ins**.

2. **Configure and Run the Macro**:
   - Open the SolidWorks VBA editor by pressing (`Alt + F11`).
   - Insert a new module and copy the provided macro code into this module.
   - Run the macro from within SolidWorks by navigating to **Tools** > **Macro** > **Run**, then select your saved macro file.

3. **Review Results**:
   - The macro will execute and identify potential issues based on predefined thresholds.
   - Review the output directly in SolidWorks or in the optional saved report.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

' Define module-level SolidWorks application variables
Dim swApp As SldWorks.SldWorks                                ' SolidWorks application object
Dim swUtil As SWUtilities.gtcocswUtilities                    ' SolidWorks Utilities interface
Dim swUtilGeometryAnalysis As SWUtilities.gtcocswGeometryAnalysis ' Geometry analysis tool interface
Dim longStatus As gtError_e                                   ' Status variable for error tracking
Dim bCheckStatus As Boolean                                   ' Status of the geometry check
Dim lTotalCount, nFailedItemCount As Long                     ' Counters for total and failed items
Dim FailedItemsArr() As String                                ' Array to store failed item details
Dim i As Long                                                 ' Loop counter

' Main subroutine to perform geometry analysis
Sub main()
    ' Connect to SolidWorks
    Set swApp = Application.SldWorks

    ' Get the SolidWorks Utilities interface
    Set swUtil = swApp.GetAddInObject("Utilities.UtilitiesApp")
    Set swUtilGeometryAnalysis = swUtil.GetToolInterface(gtSwToolGeomCheck, longStatus)

    ' Initialize the geometry analysis tool
    longStatus = swUtilGeometryAnalysis.Init()

    ' Perform geometric checks

    ' Check for short edges
    Dim lShortEdgeCount As Long
    lShortEdgeCount = swUtilGeometryAnalysis.GetShortEdgesCount(0.0005, longStatus) ' Threshold: 0.0005
    Call BadGeomList(lShortEdgeCount, "Short Edge", "EDGE") ' Log short edges

    ' Check for knife edges
    Dim lKnifeEdgeCount As Long
    lKnifeEdgeCount = swUtilGeometryAnalysis.GetKnifeEdgesCount(5#, longStatus) ' Threshold: 5 degrees
    Call BadGeomList(lKnifeEdgeCount, "Knife Edge", "EDGE") ' Log knife edges

    ' Optionally save the report to a file
    ' Uncomment and specify the desired path to enable report saving
    ' longStatus = swUtilGeometryAnalysis.SaveReport2("C:\Path\To\Save\Report", False, True)

    ' Cleanup after the analysis
    longStatus = swUtilGeometryAnalysis.Close()

    ' Update the Design Checker with the results
    Dim dcApp As Object
    Set dcApp = swApp.GetAddInObject("SWDesignChecker.SWDesignCheck")
    Call dcApp.SetCustomCheckResult(bCheckStatus, FailedItemsArr)
End Sub

' Helper subroutine to handle bad geometry listing
Public Sub BadGeomList(iNum As Long, geomType As String, eSelectType As String)
    ' Iterate over each failed geometry and log details
    For i = 1 To iNum
        nFailedItemCount = nFailedItemCount + 1
        ' Expand the FailedItemsArr array to accommodate the new failed item
        ReDim Preserve FailedItemsArr(1 To 2, 1 To nFailedItemCount)
        ' Log the failed geometry type and selection type
        FailedItemsArr(1, nFailedItemCount) = geomType & Str(i)
        FailedItemsArr(2, nFailedItemCount) = eSelectType
    Next i
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).