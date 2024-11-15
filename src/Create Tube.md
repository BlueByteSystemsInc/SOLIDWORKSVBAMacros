# Create Straight Tube in SolidWorks

## Description
This macro allows users to create a straight tube based on user-defined parameters for diameter, wall thickness, and length (in inches). The macro leverages a UserForm to gather input values from the user, simplifying the creation process and providing precise dimensions.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - The active document must be a SolidWorks part document.
> - The macro requires the UserForm (UserForm1) to collect diameter, wall thickness, and length values.

## Results
> [!NOTE]
> - A new tube is created in the active SolidWorks document with the specified diameter, wall thickness, and length.
> - The tube is centered on the origin and dimensioned accurately based on user inputs.

## STEPS to Setup the Macro

1. **Create Macro File**:
   - In SolidWorks, go to **Tools > Macro > New** to create a new macro.
   - Save the file with a `.swp` extension.

2. **Add Macro Code**:
   - Open the macro in the VBA editor.
   - Copy the **Macro Code** provided below and paste it into the VBA editor module.

3. **Add UserForm**:
   - In the VBA editor, add a new UserForm by selecting **Insert > UserForm**.
   - Name the UserForm as `UserForm1`.
   - Add input fields for **Diameter**, **Wall Thickness**, and **Length** on the form.
   - Copy the **UserForm Code** provided below and paste it into the code section of `UserForm1`.

4. **Run the Macro**:
   - In SolidWorks, open or create a new part document.
   - Run the macro to display the UserForm.
   - Enter the diameter, wall thickness, and length in the UserForm, then click **OK** to create the tube.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

Dim Part As Object
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long

Sub main()
    ' Main program area for the Main event trigger
    Set swApp = Application.SldWorks
    Set swmodel = swApp.ActiveDoc
    Load UserForm1 ' Load the UserForm for tube creation inputs
    UserForm1.Show ' Display the UserForm
End Sub
```

## VBA UserForm1 Code

```vbnet
Private Sub CommandButton1_Click()
    ' Variables for tube dimensions, converted from inches to meters
    Dim od As Double, Wall As Double, Length As Double
    od = TextBox1.Value / 39.37
    Wall = TextBox2.Value / 39.37
    Length = TextBox3.Value / 39.37

    ' Initialize SolidWorks application object
    Set swApp = Application.SldWorks
    Set Part = swApp.ActiveDoc

    ' Select Right Plane to start sketching the tube profile
    boolstatus = Part.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
    Part.SketchManager.InsertSketch True
    Part.ClearSelection2 True

    ' Create a circle for the outer diameter of the tube
    Dim skSegment As Object
    Set skSegment = Part.SketchManager.CreateCircle(-0#, 0#, 0#, od / 2, 0#, 0#)
    Part.ShowNamedView2 "*Trimetric", 8
    Part.ClearSelection2 True

    ' Extrude the circle to create the tube with specified wall thickness and length
    boolstatus = Part.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
    Dim myFeature As Object
    Set myFeature = Part.FeatureManager.FeatureExtrusionThin(True, False, False, 0, 0, Length, 0.00254, False, False, False, False, 0.01745329251994, 0.01745329251994, False, False, False, False, True, Wall, 0.00254, 0.00254, 1, 0, False, 0.005, True, True)
    Part.SelectionManager.EnableContourSelection = False

    ' Close the UserForm after creation
    Unload UserForm1
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).
