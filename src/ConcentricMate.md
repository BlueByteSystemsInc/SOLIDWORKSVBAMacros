# Auto-Mate Addition Macro for SolidWorks
<img src="../images/concentric_mate.mkv" autoplay muted controls style="width: 100%; border-radius: 12px;"></video>
## Description
This SolidWorks macro automatically adds a mate relationship to the active document, specifically targeting a precise alignment with defined parameters. It's ideal for automating assembly setup processes where specific mating conditions are frequently required.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - A document (part or assembly) must be actively open in SolidWorks.
> - The entities to be mated should be pre-selected in the correct order for the mate to apply correctly.

## Results
> [!NOTE]
> - A mate is added to the entities selected in the active document based on the specified parameters.
> - All selections are cleared after the mate is added to prevent clutter and accidental modifications.

## Steps to Setup the Macro

1. **Open SolidWorks**:
   - Start SolidWorks and open the document (part or assembly) you wish to modify.
   - Pre-select the entities that need to be mated.

2. **Load and Run the Macro**:
   - Open the SolidWorks VBA editor by pressing (`Alt + F11`).
   - Insert a new module and paste the provided macro code into this module.
   - Run the macro directly from the VBA editor or save the macro and run it from **Tools** > **Macro** > **Run**.

3. **Using the Macro**:
   - The macro executes automatically to add a mate based on the predefined parameters.
   - The selection is cleared post-execution to tidy up the workspace.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

' Declare variables for SolidWorks application and document
Dim swApp As Object               ' SolidWorks application object
Dim Part As Object                ' Active document object
Dim Gtol As Object                ' Placeholder for potential future use (e.g., geometric tolerances)

' Main subroutine
Sub main()
    ' Initialize SolidWorks application
    Set swApp = CreateObject("SldWorks.Application")
    
    ' Get the active document from SolidWorks
    Set Part = swApp.ActiveDoc

    ' Ensure an active document exists
    If Part Is Nothing Then
        MsgBox "No active document found. Please open a part or assembly."
        Exit Sub
    End If

    ' Add a mate with specified parameters
    ' Parameters: 
    ' - Type: 1 (Concentric)
    ' - Alignment: 2 (Aligned)
    ' - Distance/Angle: 0 (No offset)
    ' - Minimum distance: 0.01 meters
    ' - Angle: 0.5235987755983 radians (~30 degrees)
    Part.AddMate 1, 2, 0, 0.01, 0.5235987755983 ' Example parameters for a concentric mate
    
    ' Clear any active selections to tidy up
    Part.ClearSelection

    ' Notify the user about the operation
    MsgBox "Mate added successfully."
End Sub
```
You can download the macro from [here](../images/concentric_mate.swp)
## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).