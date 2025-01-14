# Surface Area Calculation for Painting in SolidWorks

## Description
This macro calculates the total surface area of all components within a SolidWorks assembly or part. It's particularly useful for applications where painting or surface treatment is required.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> A SolidWorks part or assembly must be open.
> The macro must be run within the SolidWorks environment.

## Results
- Calculates and displays the total surface area.
- Adds a custom property to the original document that lists the total surface area for painting.

## Steps to Setup the Macro

1. **Create the VBA Modules**:
   - Open the SolidWorks VBA editor by pressing (`Alt + F11`).
   - Insert the modules as shown in your project (`aaaMainModule`, `xxxBodSelection`, `xxxFunctions`).

2. **Run the Macro**:
   - Save the macro file (e.g., `GetSurfaceAreaForPainting.swp`).
   - Run the macro from within SolidWorks by navigating to **Tools** > **Macro** > **Run**, then select your saved macro file.

3. **Using the Macro**:
   - The macro automatically processes the active document, calculating the surface area for each body.
   - Updates the document's properties to include this new data for reference or further processing.

## VBA Macro Code

### Main Module (`aaaMainModule`)

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

' Global variable for debug mode
Global DebugMacro As Integer

' Main subroutine to execute the sequence of operations
Sub Main()
    ' Initialize the debug flag (0 = Debugging off, 1 = Debugging on)
    DebugMacro = 0  ' Set debug flag off by default

    ' Step-by-step operations in the process
    ' Retrieve the original assembly file name and path
    Call GetOriginalAssemblyFileName
    
    ' Save a defeatured version of the file for further processing
    Call SaveDefeaturedFile
    
    ' Open the defeatured file for analysis
    Call OpenDefeaturedFile
    
    ' Retrieve the file name of the defeatured file
    Call GetDefeaturedFileName
    
    ' Select all bodies (solid and surface) in the defeatured part file
    Call SelectAllBodiesInPartFile
    
    ' Calculate the surface area of the selected bodies
    Call GetSurfaceArea
    
    ' Close and delete the defeatured file to clean up
    Call CloseDefeaturedFile
    
    ' Reactivate the original assembly file in SolidWorks
    Call MakeOriginalAssemblyFileActive
    
    ' Add the calculated surface area as a custom property to the original assembly file
    Call AddSurfaceAreaForPaintingPro
    
    ' Save the original assembly file to store the updated custom property
    Call SaveOriginalAssemblyFile
End Sub
```

### Body Selection Module (xxxBodSelection)

```vbnet
Option Explicit

' Function to select all solid and surface bodies in a part file
Function SelectAllBodiesInPartFile()
    ' Declare variables for SolidWorks application, models, and bodies
    Dim swApp As SldWorks.SldWorks               ' SolidWorks application object
    Dim swModel As SldWorks.ModelDoc2            ' Active document object
    Dim swPart As SldWorks.PartDoc               ' Part document object
    Dim vBody As Variant                         ' Array of body objects
    Dim i As Long                                ' Loop counter for iterating through bodies
    Dim bRet As Boolean                          ' Boolean for operation status

    ' Initialize the SolidWorks application object
    Set swApp = Application.SldWorks

    ' Get the active document
    Set swModel = swApp.ActiveDoc

    ' Clear any existing selections in the active document
    swModel.ClearSelection2 True

    ' Debugging: Print the file path of the active document to the Immediate window
    Debug.Print "File = " & swModel.GetPathName

    ' Check the document type of the active model
    Select Case swModel.GetType
        Case swDocPART
            ' If the document is a part file
            Set swPart = swModel

            ' Select all solid bodies in the part
            vBody = swPart.GetBodies2(swSolidBody, True)
            SelectBodies swApp, swModel, vBody, ""

            ' Select all surface bodies in the part
            vBody = swPart.GetBodies2(swSheetBody, True)
            SelectBodies swApp, swModel, vBody, ""
        
        Case swDocASSEMBLY
            ' If the document is an assembly, process it accordingly
            ProcessAssembly swApp, swModel
        
        Case Else
            ' If the document is neither a part nor an assembly, exit the function
            Exit Function
    End Select
End Function
```

### Functions Module (xxxFunctions)
```vbnet
Option Explicit

' Function to retrieve the original assembly file name
Function GetOriginalAssemblyFileName()
    ' Initialize the SolidWorks application object
    Set swApp = Application.SldWorks
    
    ' Get the active document
    Set swModel = swApp.ActiveDoc
    
    ' Extract the file name of the active document
    FileName = swModel.GetTitle
    
    ' Remove the file extension from the file name
    FilenameLessExt = Left(FileName, Len(FileName) - 7)
    
    ' Extract the file path without the file name
    FilePath = Replace(swModel.GetPathName, FileName, "")
    
    ' Debugging: Display the file name without extension if debugging is enabled
    If DebugMacro = 1 Then MsgBox FilenameLessExt
End Function

' Function to save a defeatured version of the file
Function SaveDefeaturedFile()
    ' Save a defeatured version of the file for processing
    Call DateAndTime ' Generate the current timestamp for file naming
    
    ' Construct the path and name for the defeatured file
    DefeaturedPathName = FilePath & "for-getting-surfacearea" & " " & CurrentDateAndTime & ".sldprt"
    
    ' Debugging: Display the defeatured file path if debugging is enabled
    If DebugMacro = 1 Then MsgBox DefeaturedPathName
    
    ' Save the defeatured file using SolidWorks API
    boolstatus = swModel.Extension.SaveDefeaturedFile(DefeaturedPathName)
End Function

' Function to generate a timestamp for file naming
Function DateAndTime()
    ' Format the current date
    sDay = Format(Day(Date), "00")
    sMonth = Format(Month(Date), "00")
    sYear = Format(Year(Date), "0000")
    CurrentDate = sYear & "-" & sMonth & "-" & sDay
    
    ' Format the current time
    sHour = Left(Time, 2)
    sMinute = Mid(Time, 4, 2)
    sSecond = Right(Time, 2)
    CurrentTime = sHour & sMinute & sSecond
    
    ' Combine date and time into a single timestamp
    CurrentDateAndTime = "(" & CurrentDate & " " & CurrentTime & ")"
    
    ' Debugging: Display the timestamp if debugging is enabled
    If DebugMacro = 1 Then MsgBox CurrentDateAndTime
End Function

' Function to open the defeatured file
Function OpenDefeaturedFile()
    ' Open the defeatured file using SolidWorks API
    Set swModel = swApp.OpenDoc6(DefeaturedPathName, 1, 0, "", longstatus, longwarnings)
    
    ' Set the active document to the opened defeatured file
    Set swModel = swApp.ActiveDoc
End Function

' Function to calculate the surface area of the defeatured model
Function GetSurfaceArea()
    ' Create a mass property object for the defeatured model
    Set swMass = swModel.Extension.CreateMassProperty
    
    ' Use system units for the calculation
    swMass.UseSystemUnits = False
    
    ' Retrieve the total surface area of the defeatured model
    TotalSurfaceArea = swMass.SurfaceArea
End Function

' Function to close and delete the defeatured file
Function CloseDefeaturedFile()
    ' Close the defeatured file in SolidWorks
    swApp.QuitDoc DefeaturedFileName
    
    ' Delete the defeatured file from the file system
    Kill DefeaturedPathName
End Function

' Function to add the calculated surface area to the original file's custom properties
Function AddSurfaceAreaForPaintingPro()
    ' Remove any existing custom property for "Total Painted Surface Area"
    retval = swModel.DeleteCustomInfo2("", "Total Painted Surface Area")
    
    ' Add the new surface area as a custom property to the original file
    retval = swModel.AddCustomInfo3("", "Total Painted Surface Area", 30, TotalSurfaceArea)
End Function

' Function to save the original assembly file
Function SaveOriginalAssemblyFile()
    ' Save the active document in SolidWorks
    swModel.Save
End Function
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).