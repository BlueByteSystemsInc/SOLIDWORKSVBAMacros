# Traverse and Export SOLIDWORKS Components to DXF

## Description
This VBA macro automates traversing through all components of an active SOLIDWORKS assembly and exporting each part as a DXF file. It handles traversing, exporting flat patterns for sheet metal parts, and saving to a specified location.

## System Requirements
- **SOLIDWORKS Version**: SOLIDWORKS 2018 or later
- **VBA Environment**: Pre-installed with SOLIDWORKS (Access via Tools > Macro > New or Edit)
- **Operating System**: Windows 7, 8, 10, or later

## VBA Code:
```vbnet
Option Explicit

' DISCLAIMER: 
' This macro is provided "as is" without any warranty. Blue Byte Systems Inc. is not liable for any issues that arise 
' from its use. Always test the macro in a safe environment before applying it to production data.

Sub Main()
    ' Initialize SOLIDWORKS application and set active document
    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Set swApp = CreateObject("SldWorks.Application")
    Set swModel = swApp.ActiveDoc

    ' Prompt user for save path
    Dim savePath As String
    savePath = InputBox("Where do you want to save the files?")

    ' Traverse the active document to process components
    TraverseComponents swApp.ActiveDoc, savePath
End Sub

' Traverse through components and process each one
Sub TraverseComponents(swModel As ModelDoc2, savePath As String)
    Dim swApp As SldWorks.SldWorks
    Dim swRootComp As SldWorks.Component2
    Dim swConf As SldWorks.Configuration
    Dim swConfMgr As SldWorks.ConfigurationManager
    Dim vChildComp As Variant
    Dim i As Long
    Dim swChildComp As SldWorks.Component2
    
    ' Set the application object
    Set swApp = CreateObject("SldWorks.Application")
    Set swConfMgr = swModel.ConfigurationManager
    Set swConf = swConfMgr.ActiveConfiguration
    Set swRootComp = swConf.GetRootComponent3(True)
    
    ' Get child components
    vChildComp = swRootComp.GetChildren
    
    ' Loop through each child component
    For i = 0 To UBound(vChildComp)
        Set swChildComp = vChildComp(i)
        Set swModel = swChildComp.GetModelDoc2
        
        ' Check if the model exists
        If Not swModel Is Nothing Then
            If swModel.GetType = swDocASSEMBLY Then
                ' Recursively traverse sub-assemblies
                TraverseComponents swModel, savePath
            Else
                ' Process part (e.g., save as STL or DXF)
                ProcessPartToDXF swModel, savePath
            End If
        End If
    Next i
End Sub

' Process and export flat pattern of the part as DXF
Sub ProcessPartToDXF(swModel As SldWorks.ModelDoc2, savePath As String)
    Dim swFeat As SldWorks.Feature
    Dim swFlatFeat As SldWorks.Feature
    
    ' Iterate through features to find flat pattern
    Set swFeat = swModel.FirstFeature
    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName = "FlatPattern" Then
            Set swFlatFeat = swFeat
            swFeat.Select (True)
            swModel.EditUnsuppress2
            
            ' Export the flat pattern as DXF
            ExportToDXF swModel, savePath
            
            ' Suppress the flat pattern after exporting
            swFlatFeat.Select (True)
            swModel.EditSuppress2
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
End Sub

' Export the flat pattern to DXF
Sub ExportToDXF(swModel As SldWorks.ModelDoc2, savePath As String)
    Dim swPart As SldWorks.PartDoc
    Dim sModelName As String
    Dim sPathName As String
    Dim options As Long
    Dim dataAlignment(11) As Double
    
    ' Setup default alignment for export
    dataAlignment(0) = 0#: dataAlignment(1) = 0#: dataAlignment(2) = 0#
    dataAlignment(3) = 1#: dataAlignment(4) = 0#: dataAlignment(5) = 0#
    dataAlignment(6) = 0#: dataAlignment(7) = 1#: dataAlignment(8) = 0#
    dataAlignment(9) = 0#: dataAlignment(10) = 0#: dataAlignment(11) = 1#
    
    ' Get model and path names
    sModelName = swModel.GetPathName
    sPathName = savePath & "\" & swModel.GetTitle & ".dxf"
    
    ' Set export options
    options = 13 ' Export flat pattern geometry, bend lines, and sketches
    
    ' Perform DXF export
    Set swPart = swModel
    swPart.ExportToDWG sPathName, sModelName, 1, True, dataAlignment, False, False, options, Null
End Sub

' Function to extract the title from a file path
Public Function GetTitle(filePath As String) As String
    Dim pathParts As Variant
    pathParts = Split(filePath, "\")
    GetTitle = Left(pathParts(UBound(pathParts)), InStr(pathParts(UBound(pathParts)), ".") - 1)
End Function
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).