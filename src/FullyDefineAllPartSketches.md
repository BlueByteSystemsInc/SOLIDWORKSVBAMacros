# Fully Define Under-Defined Sketches in Part Feature Tree

## Description
This macro traverses the part feature tree and fully defines any sketch that is under-defined. It is particularly useful for automating the process of constraining sketches to ensure all dimensions and relations are applied correctly. The macro checks each sketch within the part and applies the `FullyDefineSketch` method if it is found to be under-defined.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - The active document must be a part file.
> - The part file must contain sketches or features with sketches (e.g., holes, extrudes).
> - Ensure the part is open and active before running the macro.

## Results
> [!NOTE]
> - All under-defined sketches in the part will be fully defined with dimensions and relations.
> - A confirmation message or error message will be displayed based on the operation's success.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

' Declare global variables
Dim swApp As Object                          ' SolidWorks application object
Dim Part As Object                           ' Active document object (part)
Dim SelMgr As Object                         ' Selection manager object
Dim boolstatus As Boolean                    ' Boolean status variable to capture operation results
Dim longstatus As Long, longwarnings As Long ' Long status and warning variables for operations
Dim Feature, swSketch As Object              ' Feature and Sketch objects for iterating through features and accessing sketches
Dim SubFeatSketch As Object                  ' Sub-feature sketch object for handling sketches inside features like Hole Wizard
Dim SketchName, MsgStr, FeatType, SubFeatType As String ' Strings for storing feature names, types, and messages
Dim EmptyStr, SubFeatName As String          ' Empty strings for message formatting and sub-feature names
Dim longSketchStatus As Long                 ' Status variable for checking if the sketch is fully defined

' --------------------------------------------------------------------------
' Main subroutine to traverse the feature tree and fully define under-defined sketches
' --------------------------------------------------------------------------
Sub main()

    ' Initialize the SolidWorks application and get the active document
    Set swApp = Application.SldWorks
    Set Part = swApp.ActiveDoc
    Set SelMgr = Part.SelectionManager

    ' Check if the active document is a part file
    longstatus = Part.GetType
    If longstatus <> 1 Then End   ' Exit if the document type is not a part (1 = swDocPART)

    ' Get the first feature in the feature tree of the part
    Set Feature = Part.FirstFeature

    ' Loop through each feature in the feature tree until no more features are found
    Do While Not Feature Is Nothing
        
        ' Get the feature name and type
        FeatName = Feature.Name
        FeatType = Feature.GetTypeName
        
        ' Check if the feature is a sketch-based feature (e.g., "ProfileFeature" for extrudes, revolves, etc.)
        If FeatType = "ProfileFeature" Then
            ' Get the sketch associated with the feature
            Set swSketch = Feature.GetSpecificFeature2

            ' Check the constraint status of the sketch (e.g., fully defined, under-defined)
            longSketchStatus = swSketch.GetConstrainedStatus()
            ' If the sketch is under-defined (2 = swUnderDefinedSketch), fully define it
            If longSketchStatus = 2 Then
                ' Clear any existing selections in the document
                Part.ClearSelection2 True

                ' Select the under-defined sketch by its name
                boolstatus = Part.Extension.SelectByID2(FeatName, "SKETCH", 0, 0, 0, False, 0, Nothing, 0)

                ' Enter the sketch edit mode
                Part.EditSketch
                Part.ClearSelection2 True

                ' Select the origin point of the sketch to help define constraints
                boolstatus = Part.Extension.SelectByID2("Point1@Origin", "EXTSKETCHPOINT", 0, 0, 0, True, 6, Nothing, 0)

                ' Fully define the sketch using the `FullyDefineSketch` method
                longstatus = Part.SketchManager.FullyDefineSketch(1, 1, 1023, 1, 1, Nothing, -1, Nothing, 1, -1)

                ' Clear selection and exit sketch edit mode
                Part.ClearSelection2 True
                Part.SketchManager.InsertSketch True
                Part.ClearSelection2 True
            End If
        End If

        ' Check if the feature is a Hole Wizard feature (contains a sub-feature sketch)
        If FeatType = "HoleWzd" Then
            ' Get the first sub-feature within the Hole Wizard feature (usually a sketch)
            Set SubFeatSketch = Feature.GetFirstSubFeature
            SubFeatName = SubFeatSketch.Name
            SubFeatType = SubFeatSketch.GetTypeName
            
            ' Get the sketch associated with the sub-feature
            Set swSketch = SubFeatSketch.GetSpecificFeature2
            
            ' Check the constraint status of the sub-feature sketch
            longSketchStatus = swSketch.GetConstrainedStatus()
            ' If the sketch is under-defined (2 = swUnderDefinedSketch), fully define it
            If longSketchStatus = 2 Then
                ' Clear any existing selections in the document
                Part.ClearSelection2 True

                ' Select the under-defined sub-feature sketch by its name
                boolstatus = Part.Extension.SelectByID2(SubFeatName, "SKETCH", 0, 0, 0, False, 0, Nothing, 0)

                ' Enter the sub-feature sketch edit mode
                Part.EditSketch
                Part.ClearSelection2 True

                ' Select the origin point of the sketch to help define constraints
                boolstatus = Part.Extension.SelectByID2("Point1@Origin", "EXTSKETCHPOINT", 0, 0, 0, True, 6, Nothing, 0)

                ' Fully define the sub-feature sketch using the `FullyDefineSketch` method
                longstatus = Part.SketchManager.FullyDefineSketch(1, 1, 1023, 1, 1, Nothing, -1, Nothing, 1, -1)

                ' Clear selection and exit sketch edit mode
                Part.ClearSelection2 True
                Part.SketchManager.InsertSketch True
                Part.ClearSelection2 True
            End If
        End If

        ' Move to the next feature in the feature tree
        Set Feature = Feature.GetNextFeature

    Loop

End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).


