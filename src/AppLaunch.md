# AppLaunch Macro for SolidWorks

## Description
This macro is designed to launch an external application or file based on configurations specified in an INI file. It reads paths and options from the INI file to dynamically set the execution parameters.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later
- **Permissions**: Requires permissions to execute external applications on the host machine.

## Pre-Conditions
- **INI File**: An INI file named after the macro but with an .ini extension must be present in the same directory as the macro.
- **Correct Formatting**: The INI file should be correctly formatted with sections labeled `[APPLICATION]`, `[LAUNCH]`, and `[OPTIONS]`, and it should contain valid paths and options.

## Results
- **Successful Execution**: If all paths are correct and the specified application or file exists, it will be launched using the options provided.
- **Error Handling**: The macro will provide error messages if the INI file is not found, if specified files do not exist, or if there are any other issues preventing the launch of the application.

## VBA Macro Code

```vbnet
'--------------------------------------------------------------------10/10/2003
' AppLaunch.swb - Copyright 2003 Leonard Kikstra
'------------------------------------------------------------------------------

Sub main()
  FileError = False
  Set swApp = CreateObject("SldWorks.Application")
  Source = swApp.GetCurrentMacroPathName             ' Get macro path+filename
  Source = Left$(Source, Len(Source) - 3) + "ini"    ' Change file extension to .ini
  Set FileSys = CreateObject("Scripting.FileSystemObject")
  If FileSys.FileExists(Source) Then                 ' Check if source file exists
    LaunchProgram = ""
    LaunchFile = ""
    Options = ""
    Open Source For Input As #1                      ' Open INI file for reading
    Do While Not EOF(1)                              ' Read until the end of the file
      Input #1, Reader                               ' Read a line
      If Reader = "[APPLICATION]" Then               ' Look for the [APPLICATION] section
        Do While Not EOF(1)
          Input #1, LineItem                         ' Read next line
          If LineItem <> "" Then
            If FileSys.FileExists(LineItem) Then
              LaunchProgram = LineItem               ' Set the program to launch
            End If
          Else
            GoTo EndRead1                            ' Skip to end if empty line
          End If
        Loop
EndRead1:
      ElseIf Reader = "[LAUNCH]" Then                ' Look for the [LAUNCH] section
        Do While Not EOF(1)
          Input #1, LineItem
          If LineItem <> "" Then
            If FileSys.FileExists(LineItem) Then
              LaunchFile = LineItem                  ' Set the file to launch
            Else
              MsgBox "Could not find file to launch." & Chr$(10) & LineItem
              FileError = True
            End If
          Else
            GoTo EndRead2
          End If
        Loop
EndRead2:
      ElseIf Reader = "[OPTIONS]" Then               ' Look for the [OPTIONS] section
        Do While Not EOF(1)
          Input #1, LineItem
          If LineItem <> "" Then
            Options = LineItem                       ' Read launch options
          Else
            GoTo EndRead3
          End If
        Loop
EndRead3:
      End If
    Loop
    Close #1                                         ' Close the INI file
  Else
    MsgBox "Source file " & Source & " not found."
  End If
  If LaunchProgram <> "" And Not FileError Then
    If LaunchFile <> "" Then
      Shell LaunchProgram & " " & Options & " " & LaunchFile, 1
    Else
      Shell LaunchProgram & " " & Options, 1
    End If
  ElseIf FileError Then
  Else
    MsgBox "Could not find application to launch."
  End If
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).