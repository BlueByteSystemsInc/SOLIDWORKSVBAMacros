# Traverse Assembly Components in SOLIDWORKS Using VBA

## Macro Description

This article explains how to write a VBA macro in SOLIDWORKS to traverse through an assembly's components and list their hierarchy. The macro explores the assembly structure, identifies components, and logs their names with proper indentation to reflect their parent-child relationship within the assembly.


## VBA Macro Code

```vbnet
' ********************************************************************
' DISCLAIMER: 
' This code is provided as-is with no warranty or liability by 
' Blue Byte Systems Inc. The company assumes no responsibility for 
' any issues arising from the use of this code in production.
' ********************************************************************
  Dim swApp As SldWorks.SldWorks
Dim swRootAssemblyModelDoc As ModelDoc2
 

Sub main()

    Set swApp = Application.SldWorks
    
    swApp.CommandInProgress = True
    
    Set swRootAssemblyModelDoc = swApp.ActiveDoc
    
    Dim swFeature As Feature
    
    Set swFeature = swRootAssemblyModelDoc.FirstFeature
           
    While Not swFeature Is Nothing
     TraverseFeatureForComponents swFeature
     Set swFeature = swFeature.GetNextFeature
    Wend
    
    
    swApp.CommandInProgress = False
    
End Sub

Private Sub TraverseFeatureForComponents(ByVal swFeature As Feature)
    Dim swSubFeature As Feature
                
    Dim swComponent As Component2
    
    Dim typeName As String
    
    typeName = swFeature.GetTypeName2
   
    If typeName = "Reference" Then
        Set swComponent = swFeature.GetSpecificFeature2
         
        If Not swComponent Is Nothing Then
         
         LogComponentName swComponent
           
           Set swSubFeature = swComponent.FirstFeature()
             While Not swSubFeature Is Nothing
                TraverseFeatureForComponents swSubFeature
                Set swSubFeature = swSubFeature.GetNextFeature()
             Wend
        End If
    End If
End Sub

Private Sub LogComponentName(ByVal swComponent As Component2)
    Dim parentCount As Long
    
    Dim swParentComponent As Component2
    Set swParentComponent = swComponent.GetParent()
    
    While Not swParentComponent Is Nothing
     parentCount = parentCount + 1
     Set swParentComponent = swParentComponent.GetParent()
    Wend
     
    Dim indentation As String
    indentation = Replicate(" ", parentCount)
    Debug.Print indentation & Split(swComponent.GetPathName(), "\")(UBound(Split(swComponent.GetPathName(), "\")))
End Sub
        
Public Function Replicate(RepeatString As String, ByVal NumOfTimes As Long)

    If NumOfTimes = 0 Then
     Replicate = ""
     Exit Function
    End If

    Dim s As String
    Dim c As Long
    Dim l As Long
    Dim i As Long

    l = Len(RepeatString)
    c = l * NumOfTimes
    s = Space$(c)

    For i = 1 To c Step l
        Mid(s, i, l) = RepeatString
    Next

    Replicate = s
 
End Function
```

## System Requirements
To run this VBA macro, ensure that your system meets the following requirements:

- SOLIDWORKS Version: SOLIDWORKS 2017 or later
- VBA Environment: Pre-installed with SOLIDWORKS (Access via Tools > Macro > New or Edit)
- Operating System: Windows 7, 8, 10, or later


## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).