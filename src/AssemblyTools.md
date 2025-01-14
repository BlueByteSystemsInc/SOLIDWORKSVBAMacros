# Assembly Tools Macro for SolidWorks

## Description
This macro provides functionalities to manipulate assembly components in SolidWorks. It can unsuppress all components, copy them to the current directory, or reload the assembly model.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
- An assembly document must be open in SolidWorks.
- The macro should be run from the SolidWorks environment with proper permissions to access file system operations.

## Results
- Depending on the selected operation, components will either be unsuppressed, copied, or the assembly will be reloaded.
- Provides a simple UI to interact with the operations and monitor their status.

## Steps to Setup the Macro

1. **Create the UserForm**:
   - Open the VBA editor in SolidWorks by pressing (`Alt + F11`).
   - In the Project Explorer, right-click on the project (e.g., `AssemblyTools`) and select **Insert** > **UserForm**.
     - Rename the form to `Form_AssemblyTools`.
     - Design the form with the following:
       - Add a `ListBox` named `ListBoxAssemblyTools` for displaying operation options.
       - Add a `TextBox` named `TextBoxDescribe` for displaying descriptions of the selected operation.
       - Add two buttons:
         - Launch: Set Name = `CommandLaunch` and Caption = `Launch`.
         - Close: Set Name = `CommandButtonClose` and Caption = `Close`.


2. **Add VBA Code**:
   - Copy the **Macro Code** provided below into a new module.
   - Copy the **UserForm Code** into the `Form_AssemblyTools` code-behind.
   - In the Project Explorer, right-click on the project (e.g., `AssemblyTools`) and select **Insert** > **Module**.
   - Rename the form to `SubCommon`.
   - Copy the **SubCommon Code** into the `SubCommon` code-behind.

3. **Save and Run the Macro**:
   - Save the macro file (e.g., `AssemblyTools.swp`).
   - Run the macro by going to **Tools** > **Macro** > **Run** in SolidWorks, then select your saved macro.

4. **Using the Macro**:
   - The macro will open the Assembly Tools UserForm.
   - Select the operation you want to perform from the ListBox.
   - Click **Launch** to execute the selected operation.
   - Use the **Close** button to exit the UserForm.

## VBA Macro Code

```vbnet

' -------------------------------------------------------------------12/05/2005
' AssemblyTools.swp - Copyright 2005, Leonard Kikstra
' -----------------------------------------------------------------------------
  Global swApp              As Object
  Global ModelDoc2          As Object
  Global FileTyp            As Long
  Global Bool               As Boolean
  Global Intg               As Integer
  Global Var                As Variant
  Global Lng                As Long
  Global Str                As String
  Global Dbl                As Double
  Global ConfigQty          As Integer
  Global ConfigNam          As Variant
  Global Configuration      As Object
  Global ConfigurationName  As String
  Global Counter            As Integer
  Global OpenRO             As Boolean
  Global FileRO             As Boolean
  Global FileSys            As Object
  Global GetPath            As Boolean
  Global MasterPath         As String

' -----------------------------------------------------------------------------
' Let's get out of here
' -----------------------------------------------------------------------------
Public Sub ProgExit()
  End
End Sub

' -----------------------------------------------------------------------------
' Close forms, reload model, re-enter macro
' -----------------------------------------------------------------------------
Public Sub Reload()
  Dim longstatus   As Long
  Dim longstat     As Long
  ThisDoc = ModelDoc2.GetPathName
  SubCommon.CloseAll
  On Error Resume Next
  Set ModelDoc2 = swApp.OpenDoc6(ThisDoc, _
                   swDocASSEMBLY, swOpenDocOptions_Silent, _
                   "Default", longstatus, longstat)
  If Not ModelDoc2 Is Nothing Then
    ThisDoc = ModelDoc2.GetPathName
    MyPath = SubCommon.NewParseString(UCase(ThisDoc), "\", 1, 0)
    DocName = SubCommon.NewParseString(UCase(ThisDoc), "\", 1, 1)
    GetBom = True
  End If
End Sub

' -----------------------------------------------------------------------------
' Program control to define procedures when program is launced.
' Attach to SolidWorks, Set objects, Determine file type, launch form.
' -----------------------------------------------------------------------------
Public Sub main()
  ' ---------------------------------------------------------------------------
  ' Connect with SolidWorks and get the active document
  ' Setup appropriate document objects
  ' ---------------------------------------------------------------------------
  Set swApp = CreateObject("SldWorks.Application")
  swApp.Visible = True
  swApp.UserControl = True
  Set ModelDoc2 = swApp.ActiveDoc       ' Grab currently active document
  If ModelDoc2 Is Nothing Then          ' Check to see if a document
    MsgBox "No model loaded."
  Else
    FileTyp = ModelDoc2.GetType
    If FileTyp = swDocASSEMBLY Then
      ThisDoc = ModelDoc2.GetPathName
      MyPath = SubCommon.NewParseString(UCase(ThisDoc), "\", 1, 0)
      DocName = SubCommon.NewParseString(UCase(ThisDoc), "\", 1, 1)
      FormAssemblyTools.Show
    Else
      MsgBox "Active file is not an assembly."
    End If                              ' document type
  End If                                ' document loaded
End Sub
```

## VBA UserForm Code

```vbnet
Option Compare Binary
' -----------------------------------------------------------------------------
' AssemblyTools.swp
' -----------------------------------------------------------------------------

Function TraverseAssembly(Level As Integer, Component2 As Object, _
                           MultiMode As Integer)
  ' mode    0 - Unsuppress all components
  '         1  - Copy to current directory
  Dim QtyChilden    As Integer
  Dim Children      As Variant
  Dim Child         As Object
  Dim ChildCount    As Integer
  Dim FilePath      As String
  Dim ModDoc2       As Object
  Dim MyBoole       As Boolean
  Dim FileSysObj As Object
  Set FileSysObj = CreateObject("Scripting.FileSystemObject")
  If Level > 0 Then
    ' Routines for 'Child' components only
    Select Case MultiMode
           Case 0
                  If Component2.IsSuppressed() Then
                    Sourcefile = Component2.GetPathName
                    SimpleStatus "Unsuppressing: " & Sourcefile
                    Component2.SetSuppression2 swComponentFullyResolved
                    If Component2.IsSuppressed() Then _
                      MsgBox "Cound not unsuppress component: " _
                             & Chr$(13) & Sourcefile _
                             & Chr$(13) & "Suppression state may be " _
                             & "controlled by Design Table."
                  End If
           Case 1
                  If Not Component2.IsSuppressed() Then
                    FilePath = Component2.GetPathName
                    ' Are paths equal? - No if file loaded from different dir.
                    If NewParseString(FilePath, "\", 1, 0) <> MasterPath Then
                      ' See if destination file exists
                      FilePath = Component2.GetPathName
                      If Not FileSysObj.FileExists(MasterPath & "\" _
                             & NewParseString(FilePath, "\", 1, 1)) Then
                        ' Destination does not exist
                        Sourcefile = Component2.GetPathName
                        If Not FileSysObj.FileExists(Sourcefile) Then
                          MsgBox "File: " & Sourcefile & " not found."
                        Else
                          SimpleStatus "Copying: " & Sourcefile
                          On Error GoTo FileCopyError
                          MyBoole = FileSysObj.CopyFile(Sourcefile, _
                                    MasterPath & "\")
                          On Error GoTo 0
                        End If
                      End If
                    End If
                  End If
           Case Else
    End Select
  End If
  ' Traverse children parent assembly
  Children = Component2.GetChildren                         ' Get list / children
  ChildCount = UBound(Children) + 1                         ' Get qty children
  Level = Level + 1                                         ' Increment level
  For ChildCount = 0 To (ChildCount - 1)                    ' Each child in sub
    Set Child = Children(ChildCount)                        ' Get child comp obj
    TraverseAssembly Level, Child, MultiMode                ' Trav child's comp
  Next ChildCount
Exit Function
FileCopyError:
  MsgBox "File: " & Sourcefile & " not found."
  Resume Next
End Function

Private Sub CommandButtonClose_Click()
  End
End Sub

Private Sub CommandLaunch_Click()
  Dim Configuration     As Object
  Dim RootComponent     As Object
  ControlStatus False
  Select Case ListBoxAssemblyTools.ListIndex
         Case 0 ' Unsuppress All Components
                Set Configuration = ModelDoc2.GetActiveConfiguration
                ModelDoc2.ShowConfiguration2 (Configuration.Name)
                Set RootComponent = Configuration.GetRootComponent()
                ' Recursively traverse the component and build temporary BOM
                If Not RootComponent Is Nothing Then
                  TraverseAssembly 0, RootComponent, 0
               End If
         Case 1 ' Copy All Components Here
                MasterPath = NewParseString(ModelDoc2.GetPathName + "", "\", 1, 0)
                Set Configuration = ModelDoc2.GetActiveConfiguration
                ModelDoc2.ShowConfiguration2 (Configuration.Name)
                Set RootComponent = Configuration.GetRootComponent()
                ' Recursively traverse the component and build temporary BOM
                If Not RootComponent Is Nothing Then
                  TraverseAssembly 0, RootComponent, 1
                End If
         Case 2 ' Reload Assembly Model
                main.Reload
         Case Else
  End Select
  SimpleStatus "Done."
  ControlStatus True
  Me.Repaint
End Sub

Private Sub ListBoxAssemblyTools_Click()
  Dim Configuration     As Object
  Dim RootComponent     As Object
  Dim Action            As String
  Select Case ListBoxAssemblyTools.ListIndex
         Case 0 ' Unsuppress All Components
                Action = Action & "Traverses thru the assembly and unsuppresses "
                Action = Action & "all sub components (parts and assemblies)."
                Action = Action & Chr$(13) & Chr$(13)
                Action = Action & "NOTE: Time it takes to complete this task is "
                Action = Action & "dependant on assembly size and number of "
                Action = Action & "suppressed components."
         Case 1 ' Copy All Components Here
                Action = Action & "Traverses thru the assembly and copies all "
                Action = Action & "sub components (parts and assemblies) files "
                Action = Action & "to the current directory."
                Action = Action & Chr$(13) & Chr$(13)
                Action = Action & "NOTE: Time it takes to complete this task is "
                Action = Action & "dependant on assembly size of the assembly "
                Action = Action & "and the number of files to copy."
         Case 2 ' Reload Assembly Model
                Action = Action & "Closes all documents currently loaded in "
                Action = Action & "SolidWorks and reloaded the current assembly."
         Case Else
  End Select
  TextBoxDescribe.Value = Action
  Me.Repaint
End Sub

Sub SimpleStatus(Stat)
  Me.Repaint
End Sub

' -----------------------------------------------------------------------------
' Enable/Disable controls on form
' -----------------------------------------------------------------------------
Sub ControlStatus(ControlMode As Boolean)
  ControlSet ControlMode, ListBoxAssemblyTools
  ControlSet ControlMode, TextBoxDescribe
  ControlSet ControlMode, CommandLaunch
  ControlSet ControlMode, CommandButtonClose
  Me.Repaint
End Sub

Private Sub UserForm_Initialize()
  ListBoxAssemblyTools.Clear
  ListBoxAssemblyTools.AddItem "Unsuppress All Components"
  ListBoxAssemblyTools.AddItem "Copy All Components Here"
  ListBoxAssemblyTools.AddItem "Reload Assembly Model"
  ListBoxAssemblyTools.ListIndex = -1
End Sub

Private Sub UserForm_Activate()
  If Me.ListBoxAssemblyTools.ListCount = 0 Then UserForm_Initialize
End Sub

Private Sub Start()
  main.main
End Sub
```

## SubCommon Code
```vbnet

' -----------------------------------------------------------------------------
' Common routines used by various macros/programs
' -----------------------------------------------------------------------------
' This function will count the occurance of "stringToFind" inside "theString".
' -----------------------------------------------------------------------------
Function ParseStringCount(theString As String, stringToFind As String)
  Dim CountPosition As Long
  CountPosition = 1
  CountNum = 0
  While CountPosition > 0
    CountPosition = InStr(1, theString, stringToFind)
    If CountPosition <> 0 Then    ' Which characters do we keep
      theString = Mid(theString, CountPosition + Len(stringToFind))
      CountNum = CountNum + 1
    End If
  Wend
  ParseStringCount = LTrim(CountNum)
End Function

' -----------------------------------------------------------------------------
' This function will locate the occurance of "stringToFind" inside "theString".
'   Occur  = 0       Get first occurrance
'            1       Get last occurrance
' From that position, it will return text based on the 'Retrieve' variable.
'   StrRet = 0       Get text before
'            1       Get text after
'            2       If no occurrance found, return blank
' -----------------------------------------------------------------------------
Function NewParseString(theString As String, stringToFind As String, _
                            Occur As Boolean, StrRet As Integer)
  Location = 0
  TempLoc = 0
  PosCount = 1
  While PosCount < Len(theString)
    TempLoc = InStr(PosCount, theString, stringToFind)
    If TempLoc <> 0 Then
      If Occur = 0 Then
        If Location = 0 Then
          Location = TempLoc
        End If
      Else
        Location = TempLoc
      End If
      PosCount = TempLoc + 1
    Else
      PosCount = Len(theString)
    End If
  Wend
  If Location <> 0 Then
    If StrRet = 0 Then
      theString = Mid(theString, 1, Location - 1)
    Else
      theString = Mid(theString, Location + Len(stringToFind))
    End If
  ElseIf StrRet = 2 Then
    theString = ""
  End If
  NewParseString = theString
End Function

' ---------------------------------------------------------------------------------
' Read file attributes
' ---------------------------------------------------------------------------------
Function FileAttrib(filespec, Bit, SetClear)
  Set FileSys = CreateObject("Scripting.FileSystemObject")
  Set f = FileSys.GetFile(filespec)
  If SetClear = 1 Then
    f.Attributes = f.Attributes + Bit
  Else
    f.Attributes = f.Attributes - Bit
  End If
  ' 0  Normal       4  System
  ' 1  ReadOnly     32 Archive
  ' 2  Hidden
End Function

' ---------------------------------------------------------------------------------
' Change attributes for file
' ---------------------------------------------------------------------------------
Function FileAttribRead(filespec, Bit)
  Set FileSys = CreateObject("Scripting.FileSystemObject")
  Set f = FileSys.GetFile(filespec)
  If f.Attributes And Bit Then
    FileRO = True
  Else
    FileRO = False
  End If
  ' 0  Normal       4  System
  ' 1  ReadOnly     32 Archive
  ' 2  Hidden
End Function

' -----------------------------------------------------------------------------
' Control enable/disable
' -----------------------------------------------------------------------------
Function ControlSet(ControlMode As Boolean, ControlObject As Object)
  If ControlMode = True Then
    ' Enable control
    ControlName = ControlObject.Name
    If UCase(Left$(ControlName, 7)) = UCase("ListBox") Or _
       UCase(Left$(ControlName, 7)) = UCase("TextBox") Then
      ControlObject.Locked = False
    ' ControlObject.BackColor = vbWindowBackground
    ' ControlObject.ForeColor = vbWindowText
    Else
      ControlObject.Enabled = True
    End If
  Else
    ' Disable control
    ControlName = ControlObject.Name
    If UCase(Left$(ControlName, 7)) = UCase("ListBox") Or _
       UCase(Left$(ControlName, 7)) = UCase("TextBox") Then
      ControlObject.Locked = True
    ' ControlObject.BackColor = vbButtonFace
    ' ControlObject.ForeColor = vbGrayText
    Else
      ControlObject.Enabled = False
    End If
  End If
End Function

' ---------------------------------------------------------------------------------
' Close all models currently loaded/opened in SolidWorks
' ---------------------------------------------------------------------------------
Function CloseAll()
  Set OpenDoc = swApp.ActiveDoc()
  While Not OpenDoc Is Nothing
    swApp.QuitDoc (OpenDoc.GetTitle)
    Set OpenDoc = swApp.ActiveDoc()
  Wend
End Function

Private Sub Start()
    main.main
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).