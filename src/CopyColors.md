# Copy Colors Macro

## Description
This macro copies part-level colors from the first selected assembly component to all other selected components within an assembly. It ensures components are resolved and does not require opening them in separate windows.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer  
- **Operating System**: Windows 7 or later  

## Pre-Conditions
> [!NOTE]
> - An assembly document must be open in SolidWorks.
> - At least two assembly components must be selected.
> - The first selected component serves as the source for the color properties.

## Results
> [!NOTE]
> - The color properties of the first selected component will be applied to all other selected components.
> - Affected components will be saved after the changes are applied.

## Steps to Use the Macro

### **1. Prepare the Assembly**
   - Open an assembly document in SolidWorks.
   - Select the components you wish to copy the color to. Ensure the source component (from which the color will be copied) is selected first.

### **2. Execute the Macro**
   - Run the macro in SolidWorks. It will resolve components, copy the color properties from the first selected component, and apply them to all other selected components.

### **3. Verify Changes**
   - Check the components to ensure the color has been successfully applied.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

Sub main()
    On Error GoTo ErrorHandler
    
    ' Declare variables
    Dim swApp As SldWorks.SldWorks
    Dim swDoc As SldWorks.ModelDoc2
    Dim swAssy As SldWorks.AssemblyDoc
    Dim swComp As SldWorks.Component2
    Dim swDoc2 As SldWorks.ModelDoc2
    Dim swComponents() As SldWorks.Component2
    Dim swSelMgr As SldWorks.SelectionMgr
    Dim SelCount As Integer
    Dim MatProps As Variant
    Dim i As Integer
    Dim Errors As Long, Warnings As Long
    
    ' Initialize SolidWorks application and active document
    Set swApp = Application.SldWorks
    Set swDoc = swApp.ActiveDoc
    
    ' Validate the active document type
    If swDoc Is Nothing Or swDoc.GetType <> swDocASSEMBLY Then
        MsgBox "Please open an assembly and select components to copy colors.", vbExclamation, "Error"
        Exit Sub
    End If
    
    Set swAssy = swDoc
    Set swSelMgr = swDoc.SelectionManager
    SelCount = swSelMgr.GetSelectedObjectCount
    
    ' Ensure at least two components are selected
    If SelCount < 2 Then
        MsgBox "Please select at least two components in the assembly.", vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Store selected components in an array
    ReDim swComponents(SelCount - 1)
    For i = 0 To SelCount - 1
        Set swComp = swSelMgr.GetSelectedObjectsComponent3(i + 1, -1)
        If swComp Is Nothing Then
            MsgBox "Invalid selection detected. Ensure only components are selected.", vbExclamation, "Error"
            Exit Sub
        End If
        Set swComponents(i) = swComp
    Next i
    
    ' Process each selected component
    For i = 0 To UBound(swComponents)
        Set swComp = swComponents(i)
        
        ' Resolve the component if suppressed
        If swComp.GetSuppression <> swComponentFullyResolved Then
            swComp.SetSuppression2 swComponentFullyResolved
        End If
        
        Set swDoc2 = swComp.GetModelDoc2
        
        If i = 0 Then
            ' Retrieve material properties from the first component
            MatProps = swDoc2.MaterialPropertyValues
        Else
            ' Apply material properties to the other components
            swDoc2.MaterialPropertyValues = MatProps
            swDoc2.Save3 swSaveAsOptions_Silent, Errors, Warnings
        End If
    Next i
    
    ' Notify user of successful operation
    MsgBox "Colors copied successfully to the selected components.", vbInformation, "Success"
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).