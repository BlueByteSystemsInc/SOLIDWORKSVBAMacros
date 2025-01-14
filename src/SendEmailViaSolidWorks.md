# Send Email via SolidWorks Macro

## Description
This macro allows users to send an email from within SolidWorks with the assembly name included in the email subject. It's particularly useful for quick updates or notifications about specific assemblies.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later
- **Additional Requirements**: An email client installed on the user's system that supports mailto: links (e.g., Microsoft Outlook).

## Pre-Conditions
> [!NOTE]
> - An assembly document must be currently open in SolidWorks.
> - The user's default email client must be configured to handle mailto: links.

## Results
> [!NOTE]
> - An email draft will be opened with the subject containing the name of the currently active assembly.
> - The body of the email can be customized within the macro.

## Steps to Setup the Macro

### 1. **Configure Email Details**:
   - Modify the email recipient, subject prefix, and body message in the macro code to fit your specific needs.

### 2. **Run the Macro**:
   - Execute the macro while an assembly document is active in SolidWorks. The macro checks the type of the document and proceeds only if it's an assembly.
   - The system's default email client will open a new email draft with pre-filled details.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long

Dim Email As String, Subj As String
Dim Msg As String, URL As String
Dim swApp As Object
Dim Model As Object

Sub Main()
    ' Initialize SolidWorks application and get the active document
    Set swApp = CreateObject("SldWorks.Application")
    Set Model = swApp.ActiveDoc

    ' Check if there is an active document
    If Model Is Nothing Then
        MsgBox "No active file, please open a SolidWorks file and try again.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Ensure the active document is an assembly
    If Model.GetType <> swDocASSEMBLY Then
        MsgBox "This macro works only with assemblies. Please open an assembly and try again.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Set the recipient email address (modify as needed)
    Email = "123@123.com"

    ' Compose the subject line using the assembly name
    Subj = "Assembly To Work On: " & Model.GetTitle

    ' Compose the email body (customize as needed)
    Msg = ""
    Msg = Msg & "Dear Boss," & vbCrLf & vbCrLf
    Msg = Msg & "I want to inform you about the following assembly work:" & vbCrLf
    Msg = Msg & "Your Name Here" & vbCrLf

    ' Replace spaces and line breaks with URL-encoded equivalents
    Msg = Replace(Msg, " ", "%20")
    Msg = Replace(Msg, vbCrLf, "%0D%0A")

    ' Create the mailto URL
    URL = "mailto:" & Email & "?subject=" & Subj & "&body=" & Msg

    ' Open the default email client with the pre-composed email
    ShellExecute 0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).