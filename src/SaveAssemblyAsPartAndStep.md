# Save Assembly as Part Macro for SolidWorks

## Description
This macro saves the active assembly document as a part file containing only exterior faces. This is useful for creating simplified versions of assemblies for sharing or lightweight referencing.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - An assembly document must be open in SolidWorks.
> - The user must provide a file path where the part file will be saved.

## Results
> [!NOTE]
> - The assembly is saved as a new part file at the specified path.
> - The part file will contain only the exterior faces of the assembly.

## Steps to Setup the Macro

### 1. **Open the Assembly**:
   - Ensure that the assembly document you want to save as a part file is open in SolidWorks.

### 2. **Load and Execute the Macro**:
   - Load the macro into SolidWorks using the VBA editor (`Alt + F11`).
   - Execute the macro from within SolidWorks by navigating to **Tools** > **Macro** > **Run**.

### 3. **Interact with the User Form**:
   - Upon running the macro, a user form will appear requesting the file path for saving the new part file.
   - Use the form to enter or browse to the desired save location and specify the file name.

### 4. **User Form Creation**:
   - In the VBA editor, insert a new UserForm.
   - Add a TextBox for inputting the file path.
   - Include a "Browse" button to open a file dialog allowing the user to select a directory.
   - Add a "Save" button to initiate the save process.
   - Optionally, add an "Exit" button to close the form without saving.

## VBA Macro Code

### Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

' ******************************************************************************
' Description:
' This macro checks if an assembly document is open and extracts its file path 
' and file name without extension. It then displays this information in a user form.
' ******************************************************************************

' Global variables
Public swApp                As SldWorks.SldWorks           ' SolidWorks application object
Public swModel              As ModelDoc2                  ' Active SolidWorks document
Public FilePathDefault      As String                     ' Default file path for the active document

Sub Main()

    ' Declare variables for document properties
    Dim docType                 As swDocumentTypes_e       ' Type of the active document
    Dim sModelNameSize          As Long                    ' Length of the file name
    Dim sModelName              As String                  ' File name with extension
    Dim sModelNameNoExtension   As String                  ' File name without extension
    Dim FilePathLength          As Long                    ' Length of the full file path

    ' Initialize SolidWorks application
    Set swApp = Application.SldWorks
    Set swDoc = swApp.ActiveDoc
200: _
    ' Get the active document
    Set swModel = swApp.ActiveDoc
    
    ' Check if a document is open
    If (swModel Is Nothing) Then
        Call swApp.SendMsgToUser("No assembly is open. Please open an assembly and restart the macro.") ' No assembly open
    Else
        ' Determine the type of the active document
        docType = swModel.GetType
        If (docType = swDocASSEMBLY) Then
            ' Extract the file name and its components
            sModelName = Mid(swModel.GetPathName, InStrRev(swModel.GetPathName, "\") + 1) ' File name with extension
            sModelNameNoExtension = Left(sModelName, InStrRev(sModelName, ".") - 1)      ' File name without extension
            sModelNameSize = Strings.Len(sModelName)                                     ' Length of the file name
            FilePathLength = Strings.Len(swModel.GetPathName)                            ' Length of the full file path
            
            ' Extract the directory path
            Dim n As Long
            n = FilePathLength - sModelNameSize
            FilePathDefault = Strings.Left(swModel.GetPathName, n)                       ' Directory path
            
            ' Load and display the user form
            Load UserForm1
            UserForm1.TextBox1.Text = FilePathDefault                                    ' Set the default file path in the form
            UserForm1.Show
        Else
            ' Inform the user that the macro works only with assemblies
            Call swApp.SendMsgToUser("This macro is only for processing assemblies. Please open an assembly and restart the macro.")
        End If
    End If
End Sub
```

## VBA UserForm Code

```vbnet
Option Explicit

' Constants for folder browsing options
Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Private Const MAX_PATH As Long = 260

' Function to calculate the length of a file name
Function GetFileLength(ByVal CurrentFileName As String) As Integer
    Dim PathSize As Integer
    PathSize = Strings.Len(CurrentFileName)
    GetFileLength = PathSize
End Function

' Function to browse for a folder and return its path
Function BrowseFolder(Optional Caption As String, Optional InitialFolder As String) As String
    Dim SH As Shell32.Shell
    Dim F As Shell32.Folder
    Set SH = New Shell32.Shell
    Set F = SH.BrowseForFolder(0&, Caption, BIF_RETURNONLYFSDIRS, InitialFolder)
    If Not F Is Nothing Then
        BrowseFolder = F.Items.Item.Path
    End If
End Function

' Function to list all files in a specified folder
Function listfiles(ByVal sPath As String) As Variant
    Dim vaArray As Variant
    Dim i As Integer
    Dim oFile As Object
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFiles As Object

    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(sPath)
    Set oFiles = oFolder.Files

    If oFiles.Count = 0 Then Exit Function

    ReDim vaArray(1 To oFiles.Count)
    i = 1
    For Each oFile In oFiles
        vaArray(i) = oFile.Name
        i = i + 1
    Next

    listfiles = vaArray
End Function

' Subroutine to cancel the user form
Private Sub Cancel_Click()
    Unload Me
End Sub

' Subroutine to handle folder selection
Private Sub CB1_Click()
    Dim Path As String
    Dim n As String
    n = UserForm1.TextBox1.Text

    Path = BrowseFolder("Select A Folder/Path", n)
    If Path = "" Then
        MsgBox "Please select the path and try again", vbOKOnly
        Exit Sub
    Else
        Path = Path & "\"
        TextBox1.Text = Path
    End If
End Sub

' Main subroutine to process file save and conversion operations
Private Sub CB3_Click()
    ' Declare variables for file paths and extensions
    Dim FilePath As String
    Dim sFilePath As String
    Dim PathSize As Long
    Dim PathNoExtension As String
    Dim NewFilePath As String
    Dim nErrors As Long
    Dim nWarnings As Long
    Dim swModelDocExt As SldWorks.ModelDocExtension
    Dim sModelName As String
    Dim Step As Long
    Dim Response As Variant
    Dim ListOfExistingFilesOriginal() As Variant
    Dim LengthOfFileListOriginal As Integer

    ' Initialize SolidWorks application and active document
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set swModelDocExt = swModel.Extension

    FilePath = UserForm1.TextBox1.Text
    PathSize = Strings.Len(FilePath)

    If FilePath = "" Then
        MsgBox "Target path not specified, please specify a target path!", vbExclamation, "Error"
        Exit Sub
    End If

    PathNoExtension = Strings.Left(FilePath, PathSize)

    ' Check if the directory contains files
    If Dir(PathNoExtension & "*.*") = "" Then
        sModelName = Mid(swModel.GetPathName, InStrRev(swModel.GetPathName, "\") + 1)
        sModelName = Left(sModelName, InStrRev(sModelName, ".") - 1)
        NewFilePath = PathNoExtension & sModelName & ".SLDPRT"
        Response = MsgBox("The target folder contains no files." & vbCrLf & vbCrLf & "Default file name is: " & vbCrLf & vbCrLf & sModelName _
                          & vbCrLf & vbCrLf & "Would you like to proceed with this name?", vbYesNo)
        If Response = vbYes Then
            GoTo ProcessFile
        ElseIf Response = vbNo Then
            Exit Sub
        End If
    Else
        ' List existing files in the directory
        ListOfExistingFilesOriginal() = listfiles(PathNoExtension)
        LengthOfFileListOriginal = UBound(ListOfExistingFilesOriginal, 1) - LBound(ListOfExistingFilesOriginal, 1) + 1
        sModelName = Mid(swModel.GetPathName, InStrRev(swModel.GetPathName, "\") + 1)
        sModelName = Left(sModelName, InStrRev(sModelName, ".") - 1)

        Response = MsgBox("The following files already exist in the folder:" & vbCrLf & vbCrLf & Join(ListOfExistingFilesOriginal, vbCrLf & vbCrLf) _
                          & vbCrLf & vbCrLf & "Default file name: " & vbCrLf & sModelName & vbCrLf & vbCrLf & "Proceed? Yes/No", vbYesNo)
        If Response = vbNo Then
            Dim nmbText As String
            nmbText = InputBox("Enter a new file name", "Bohaa Inc.", sModelName)
            If nmbText = vbNullString Then Exit Sub
            sModelName = nmbText
            NewFilePath = PathNoExtension & sModelName & ".SLDPRT"
        ElseIf Response = vbYes Then
            NewFilePath = PathNoExtension & sModelName & ".SLDPRT"
            GoTo ProcessFile
        End If
    End If

ProcessFile:
    ' Save the assembly as a part file and optionally generate a STEP file
    swApp.SetUserPreferenceIntegerValue swSaveAssemblyAsPartOptions, swSaveAsmAsPart_ExteriorFaces
    swModelDocExt.SaveAs NewFilePath, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, nErrors, nWarnings
    Response = MsgBox("Generate STEP file?", vbYesNo)
    If Response = vbYes Then
        Set swModel = swApp.OpenDoc(NewFilePath, swDocPART)
        Set swModel = swApp.ActiveDoc
        Step = swApp.SetUserPreferenceIntegerValue(swStepAP, 214)
        sFilePath = PathNoExtension & sModelName & ".STEP"
        swModel.SaveAs sFilePath
        swApp.CloseDoc NewFilePath
    End If
End Sub

Private Sub CB4_Click()
    Set swApp = Nothing
    Unload Me
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).