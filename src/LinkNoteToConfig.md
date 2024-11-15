# Link Note to Configuration Specific Property in SolidWorks

## Description
This macro allows users to link a note in a SolidWorks document to a configuration-specific property. The macro pushes the value of the note into a custom property in the configuration, but it does not update the note with changes in the property value. The macro prompts the user to enter the property name to which the selected note will be linked. Additionally, a macro feature is created, ensuring that the note and custom property linkage is maintained.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - The user must preselect a note in the SolidWorks model before running the macro.
> - The active document must support configuration-specific properties.
> - A valid name for the property should be provided in the input box on the UserForm.

## Results
> [!NOTE]
> - The selected note's text is linked to a configuration-specific property with the specified name.
> - A macro feature is added to maintain this link, allowing the custom property to update as the note text changes.
> - A message box appears if no note is selected or if an invalid name is entered.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

' Main subroutine to display the UserForm
Sub Main()
    UserForm1.Show
End Sub


' Create a Module named FeatureModule and paste the code Below
vbnet
Copy code
Option Explicit

' Rebuild routine for the macro feature
Function swmRebuild(varApp As Variant, varDoc As Variant, varFeat As Variant) As Variant
    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swCusPropMgr As CustomPropertyManager
    Dim swCfgPropMgr As CustomPropertyManager
    
    Set swApp = varApp
    Set swModel = varDoc
    Set swCusPropMgr = swModel.Extension.CustomPropertyManager("")
    
    Dim CustomPropNames As Variant
    Dim CustomPropValues As Variant
    swCusPropMgr.GetAll CustomPropNames, Nothing, CustomPropValues
    
    Dim i As Integer
    If swCusPropMgr.Count = 0 Then Exit Function
    
    For i = LBound(CustomPropNames) To UBound(CustomPropNames)
        If Left(CustomPropNames(i), 9) = "Linked - " Then
            Set swCfgPropMgr = swModel.Extension.CustomPropertyManager(swModel.ConfigurationManager.ActiveConfiguration.name)
            swCfgPropMgr.Delete CustomPropValues(i)
            swCfgPropMgr.Add2 CustomPropValues(i), swCustomInfoText, GetNoteTextByName(Mid(CustomPropNames(i), 10))
        End If
    Next
End Function

' Function to get the text of a note by name
Function GetNoteTextByName(ByVal name As String)
    Dim swApp As SldWorks.SldWorks
    Dim swModel As ModelDoc2
    Dim swNote As Note
    Dim swSelMgr As SelectionMgr
    
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set swSelMgr = swModel.SelectionManager
    
    swModel.Extension.SelectByID2 name & "@Annotations", "NOTE", 0, 0, 0, False, -1, Nothing, swSelectOptionDefault
    Set swNote = swSelMgr.GetSelectedObject6(1, -1)
    If Not swNote Is Nothing Then GetNoteTextByName = swNote.GetText
End Function

' Edit definition routine for the macro feature
Function swmEditDefinition(varApp As Variant, varDoc As Variant, varFeat As Variant) As Variant
    UserForm1.Show
End Function

' Security routine for the macro feature
Function swmSecurity(varApp As Variant, varDoc As Variant, varFeat As Variant) As Variant
    swmSecurity = SwConst.swMacroFeatureSecurityOptions_e.swMacroFeatureSecurityByDefault
End Function

' Subroutine to create the macro feature if it doesn't already exist
Sub CreateNewMacroFeature(ByRef swApp As SldWorks.SldWorks)
    If Not LinkFeatureExists(swApp) Then
        Dim swModel As SldWorks.ModelDoc2
        Dim feat As Feature
        Dim Methods(8) As String
        Dim Names As Variant, Types As Variant, Values As Variant
        Dim options As Long
        Dim icons(2) As String

        Set swModel = swApp.ActiveDoc
        ThisFile = swApp.GetCurrentMacroPathName
        Methods(0) = ThisFile: Methods(1) = "FeatureModule": Methods(2) = "swmRebuild"
        Methods(3) = ThisFile: Methods(4) = "FeatureModule": Methods(5) = "swmEditDefinition"

        options = swMacroFeatureAlwaysAtEnd
        Set feat = swModel.FeatureManager.InsertMacroFeature3("Link Properties", "", Methods, Names, Types, Values, Empty, Empty, Empty, icons, options)
        swModel.ForceRebuild3 False
    End If
End Sub

' Function to check if the macro feature already exists
Function LinkFeatureExists(ByRef swApp As SldWorks.SldWorks)
    Dim swModel As ModelDoc2
    Dim swFeat As Feature
    Set swModel = swApp.ActiveDoc
    LinkFeatureExists = False
    Set swFeat = swModel.FirstFeature
    Do Until swFeat Is Nothing
        If Left(swFeat.Name, Len("Link Properties")) = "Link Properties" Then
            LinkFeatureExists = True
            Exit Function
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
End Function

'End of Feature Module

'Create User Form With TestBox and 2 Command Buttons as Below
'textBox1
'CommandButton1
'CommandButton2

'Paste the below in the backcode of the User Form
Option Explicit

Private Sub CommandButton1_Click()
    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks

    If TextBox1.Text <> "" Then
        CreateNoteProperties TextBox1.Text
        CreateNewMacroFeature swApp
    Else
        MsgBox "Enter a name first"
    End If
End Sub

' Function to create custom properties for the note
Function CreateNoteProperties(ByVal OutputName As String) As String
    Dim swApp As SldWorks.SldWorks
    Dim swModel As ModelDoc2
    Dim swSelMgr As SelectionMgr
    Dim swNote As Note

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set swSelMgr = swModel.SelectionManager

    If swSelMgr.GetSelectedObjectCount2(-1) = 0 Or swSelMgr.GetSelectedObjectType3(1, -1) <> 15 Then
        MsgBox "You must select a note to link first"
        Exit Function
    End If

    Set swNote = swSelMgr.GetSelectedObject6(1, -1)
    If Not swNote Is Nothing Then
        CreateNoteProperties = swNote.GetText
        swModel.Extension.CustomPropertyManager("").Delete "Linked - " & swNote.GetName
        swModel.Extension.CustomPropertyManager("").Add2 "Linked - " & swNote.GetName, swCustomInfoText, OutputName
    End If
End Function

```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).