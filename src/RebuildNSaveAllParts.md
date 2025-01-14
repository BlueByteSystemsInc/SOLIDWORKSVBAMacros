# Rebuild And Save All Parts Macro

## Description
This macro rebuilds and saves a SolidWorks assembly document and all its dependent parts and assemblies. It iterates through dependencies, applies changes, and ensures that all associated files are updated.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer  
- **Operating System**: Windows 7 or later  

## Pre-Conditions
> [!NOTE]
> - An assembly document must be active in SolidWorks.
> - Ensure that all referenced files are accessible during the operation.

## Results
> [!NOTE]
> - The assembly and its dependencies are rebuilt and saved.
> - A message is displayed upon successful completion.

## Steps to Use the Macro

### 1. Add a UserForm
   - Open the Visual Basic for Applications (VBA) editor (Alt + F11).
   - Insert a new UserForm:
     1. Go to **Insert** > **UserForm**.
     2. Rename the UserForm to `frmRebuild`.
     3. Add a `Label` to the UserForm and name it `Label1`.
     4. Adjust the layout as shown in the attached image.
     5. Add the UserForm Code to the UserForm.

### 2. Update the Macro Code
   - Ensure the main macro integrates the UI for rebuild status updates.
   - Use the `ChangeLabel` method to update the status during processing.

### 3. Execute the Macro
   - Open an assembly document in SolidWorks.
   - Run the macro to rebuild and save the assembly and its dependencies.
   - The UI will display the status for each file being processed.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

' Public variables for SolidWorks application and global data
Public swApp As SldWorks.SldWorks
Public swBlisterItem As String
Public swProjectNumber As String
Dim Mainsourcefiles() As String
Dim changedfiles() As String
Dim vRebuild As New frmRebuild ' Form instance for rebuild progress display

' Subroutine to rebuild and save the current part and its dependencies
Public Sub Change_Stamp(ByRef swpart As ModelDoc2)
    Dim swDocName As String
    Dim hDname As Variant
    Dim vDocName As Variant
    Dim isChanged As Boolean
    Dim boolstatus As Boolean
    Dim sModel As SldWorks.ModelDoc2
    Dim nRetVal As Long
    Dim nErrors As Long
    Dim nWarnings As Long

    ' Show progress form
    vRebuild.Show vbModeless

    ' Check if the document is a part
    If swpart.GetType = swDocPART Then
        ' Force rebuild for part documents
        swpart.ForceRebuild3 True
        Debug.Print "Rebuilt part: " & swpart.GetPathName
    Else
        ' Get all dependent files for the assembly
        Mainsourcefiles = Get_Depends(swpart)
        Dim AssemblyName As String
        ReDim Preserve changedfiles(0)
        AssemblyName = swpart.GetPathName
        changedfiles(0) = AssemblyName

        ' Loop through all dependent files
        For Each vDocName In Mainsourcefiles
            Dim pType As String
            pType = UCase(Mid(vDocName, InStrRev(vDocName, ".", -1) + 1))
            swDocName = vDocName
            DoEvents
            vRebuild.ChangeLabel swDocName
            DoEvents
            Set sModel = swApp.ActivateDoc2(swDocName, True, nRetVal)

            ' If the document is a part
            If pType = "SLDPRT" Then
                isChanged = False
                ' Check if the file is already rebuilt
                For Each hDname In changedfiles
                    If UCase(hDname) = UCase(vDocName) Then
                        isChanged = True
                        Exit For
                    End If
                Next
                ' Rebuild if not already done
                If Not isChanged Then
                    If Not sModel Is Nothing Then
                        sModel.ForceRebuild3 True
                    End If
                End If
            End If

            ' If the document is an assembly, process its dependencies recursively
            If Not sModel Is Nothing Then
                If pType = "SLDASM" Then
                    Call Cycle_Through_Dependents(sModel)
                End If
                sModel.Visible = False
            End If
            Set sModel = Nothing
        Next

        ' Hide progress form
        vRebuild.Hide

        ' Rebuild and save the main assembly
        Set swpart = swApp.ActivateDoc2(AssemblyName, True, nRetVal)
        swpart.ForceRebuild3 True
        swpart.Save3 swSaveAsOptions_SaveReferenced, nErrors, nWarnings
    End If
End Sub

' Recursive function to process dependents of an assembly
Private Sub Cycle_Through_Dependents(ByRef swpart As ModelDoc2)
    Dim sourcefiles() As String
    Dim AssemblyName As String
    Dim vDocName As Variant
    Dim swDocName As String
    Dim nRetVal As Long
    Dim sModel As ModelDoc2
    Dim pType As String
    Dim isInTheAssembly As Boolean

    sourcefiles = Get_Depends(swpart)
    AssemblyName = swpart.GetPathName

    ' Loop through all dependent files
    For Each vDocName In sourcefiles
        pType = UCase(Mid(vDocName, InStrRev(vDocName, ".", -1) + 1))
        swDocName = vDocName
        isInTheAssembly = False

        ' Check if the file is already processed
        For Each hDname In Mainsourcefiles
            If UCase(hDname) = UCase(vDocName) Then
                isInTheAssembly = True
                Exit For
            End If
        Next

        If Not isInTheAssembly Then
            DoEvents
            vRebuild.ChangeLabel swDocName
            DoEvents
            Set sModel = swApp.ActivateDoc2(swDocName, True, nRetVal)

            ' If the document is a part
            If pType = "SLDPRT" Then
                sModel.ForceRebuild3 True
            End If

            ' If the document is an assembly, process its dependencies recursively
            If pType = "SLDASM" Then
                Call Cycle_Through_Dependents(sModel)
            End If

            ' Close the document after processing
            sModel.Visible = False
            swApp.CloseDoc sModel.GetPathName
            Set sModel = Nothing
        End If
    Next

    ' Rebuild the main assembly after processing dependents
    Set swpart = swApp.ActivateDoc2(AssemblyName, True, nRetVal)
    swpart.ForceRebuild3 True
End Sub

' Function to get all dependent files of a given document
Private Function Get_Depends(ByRef swpart As ModelDoc2) As Variant
    Dim depends As Variant
    Dim sourcefiles() As String
    Dim d As Integer

    depends = swApp.GetDocumentDependencies2(swpart.GetPathName, True, True, False)
    For d = 1 To UBound(depends) Step 2
        ReDim Preserve sourcefiles(UBound(sourcefiles) + 1)
        sourcefiles(UBound(sourcefiles)) = depends(d)
    Next
    Get_Depends = sourcefiles
End Function

' Main entry point for the macro
Sub main()
    Dim swpart As SldWorks.ModelDoc2
    Dim intResp As Integer

    ' Initialize SolidWorks application and get the active document
    Set swApp = Application.SldWorks
    Set swpart = swApp.ActiveDoc

    If swpart Is Nothing Then
        swApp.SendMsgToUser2 "No active file.", swMbInformation, swMbOk
        Exit Sub
    End If

    ' Prompt user for confirmation
    intResp = swApp.SendMsgToUser2("The macro will rebuild and save the document and all dependents. Do you wish to continue?", swMbInformation, swMbYesNo)
    If intResp = swMbHitNo Then Exit Sub

    ' Start the processing
    Change_Stamp swpart
    swApp.SendMsgToUser2 "Macro finished successfully.", swMbInformation, swMbOk
End Sub
```

## VBA UserForm Code

```vbnet
Option Explicit

' Public variable to hold the label text
Public lblText As String

' Event triggered when the UserForm is activated
Private Sub UserForm_Activate()
    ' This is a placeholder for additional UI activation logic
    ' Add any initialization logic here if required
End Sub

' Public subroutine to change the caption of Label1 dynamically
' Parameters:
'   vText (String) - The text to set as the label caption
Public Sub ChangeLabel(ByVal vText As String)
    ' Update the label's caption with the provided text
    Me.Label1.Caption = vText
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).