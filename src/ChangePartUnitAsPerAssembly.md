# Match Unit System of All Sub-Parts and Sub-Assemblies with Main Assembly

## Description
This macro changes the unit system of all sub-parts and sub-assemblies in the active assembly to match the unit system of the main assembly. The macro ensures that all components in the assembly have a consistent unit system, which is crucial for accurate measurement and interoperability.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - The active document must be an assembly.
> - The macro should be run with all necessary permissions to modify and save the components.

## Results
> [!NOTE]
> - All sub-parts and sub-assemblies in the assembly will have their unit systems changed to match the main assembly's unit system.
> - The changes will be saved, and a message box will display the updated unit system.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

' Declare global variables
Dim swApp As SldWorks.SldWorks                    ' SolidWorks application object
Dim swmodel As SldWorks.ModelDoc2                 ' Active model document (assembly)
Dim swasm As SldWorks.AssemblyDoc                 ' Assembly document object
Dim swconf As SldWorks.Configuration              ' Configuration object
Dim swrootcomp As SldWorks.Component2             ' Root component of the assembly
Dim usys As Long                                  ' Main assembly unit system
Dim usys1 As Long                                 ' Main assembly linear units
Dim dunit As Long                                 ' Dual linear unit system value
Dim bret As Boolean                               ' Boolean return status variable
Dim err As Long, war As Long                      ' Error and warning variables

' --------------------------------------------------------------------------
' Main subroutine to initialize the process and update unit systems
' --------------------------------------------------------------------------
Sub main()

    ' Initialize SolidWorks application and get the active document
    Set swApp = Application.SldWorks
    Set swmodel = swApp.ActiveDoc

    ' Check if there is an active document open
    If swmodel Is Nothing Then
        MsgBox "No active document found. Please open an assembly and try again.", vbCritical, "No Active Document"
        Exit Sub
    End If

    ' Check if the active document is an assembly
    If swmodel.GetType <> swDocASSEMBLY Then
        MsgBox "This macro only works on assemblies. Please open an assembly and try again.", vbCritical, "Invalid Document Type"
        Exit Sub
    End If

    ' Get the active configuration and root component of the assembly
    Set swconf = swmodel.GetActiveConfiguration
    Set swrootcomp = swconf.GetRootComponent3(True)

    ' Get the main assembly's unit system and dual units
    usys = swmodel.GetUserPreferenceIntegerValue(swUnitSystem)         ' Unit system (CGS, MKS, IPS, etc.)
    dunit = swmodel.GetUserPreferenceIntegerValue(swUnitsDualLinear)   ' Dual linear unit system
    If usys = 4 Then
        usys1 = swmodel.GetUserPreferenceIntegerValue(swUnitsLinear)   ' Custom linear units
    End If

    ' Traverse through all sub-components and update their unit systems
    Traverse swrootcomp, 1

    ' Notify the user about the updated unit system
    Select Case usys
        Case 1
            swApp.SendMsgToUser2 "Unit system changed to CGS", swMbInformation, swMbOk
        Case 2
            swApp.SendMsgToUser2 "Unit system changed to MKS", swMbInformation, swMbOk
        Case 3
            swApp.SendMsgToUser2 "Unit system changed to IPS", swMbInformation, swMbOk
        Case 4
            swApp.SendMsgToUser2 "Unit system changed to Custom Unit", swMbInformation, swMbOk
        Case 5
            swApp.SendMsgToUser2 "Unit system changed to MMGS", swMbInformation, swMbOk
    End Select

End Sub

' --------------------------------------------------------------------------
' Recursive function to traverse through the assembly and update unit systems
' --------------------------------------------------------------------------
Sub Traverse(swcomp As SldWorks.Component2, nlevel As Long)

    ' Declare necessary variables
    Dim vChildComp As Variant                       ' Array of child components in the assembly
    Dim swChildComp As SldWorks.Component2          ' Individual child component object
    Dim swCompConfig As SldWorks.Configuration      ' Component configuration object
    Dim swpmodel As SldWorks.ModelDoc2              ' Model document object of the component
    Dim path As String                              ' Path of the component file
    Dim sPadStr As String                           ' String for formatting debug output
    Dim i As Long                                   ' Loop counter for iterating through child components

    ' Format padding for debug output based on level
    For i = 0 To nlevel - 1
        sPadStr = sPadStr + "  "
    Next i

    ' Get child components of the current component
    vChildComp = swcomp.GetChildren

    ' Loop through each child component
    For i = 0 To UBound(vChildComp)
        Set swChildComp = vChildComp(i)    ' Set the child component

        ' Recursively traverse through sub-components
        Traverse swChildComp, nlevel + 1

        ' Check if the child component is valid
        If Not swChildComp Is Nothing Then
            path = swChildComp.GetPathName ' Get the path of the component

            ' Open the part or assembly based on file extension
            If (LCase(Right(path, 3)) = "prt") Then
                Set swpmodel = swApp.OpenDoc6(path, swDocPART, 0, swChildComp.ReferencedConfiguration, err, war)
            ElseIf (LCase(Right(path, 3)) = "asm") Then
                Set swpmodel = swApp.OpenDoc6(path, swDocASSEMBLY, 0, swChildComp.ReferencedConfiguration, err, war)
            End If

            ' If the component is successfully opened, update its unit system
            If Not swpmodel Is Nothing Then
                bret = swpmodel.SetUserPreferenceIntegerValue(swUnitSystem, usys)
                bret = swpmodel.SetUserPreferenceIntegerValue(swUnitsDualLinear, dunit)
                If usys = 4 Then
                    bret = swpmodel.SetUserPreferenceIntegerValue(swUnitsLinear, usys1)
                End If

                ' Save the component after updating the unit system
                swpmodel.Save3 0, err, war
                Set swpmodel = Nothing  ' Release the object
            End If
        End If
    Next i

End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).

