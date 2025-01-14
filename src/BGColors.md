# Background Colors Macro for SolidWorks

## Description
This macro facilitates customization of the background settings within SolidWorks. Users can enable or disable gradient backgrounds for models, set top and bottom gradient colors, and choose the color for the viewport background in drawings.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
- The macro should be executed within the SolidWorks environment with sufficient user privileges to change system settings.

## Results
- Users can customize the gradient background settings for models and the background color for drawings.
- Provides visual feedback by updating the UI with the chosen colors.

## Steps to Setup the Macro

1. **Create the UserForm**:
   - Open the VBA editor in SolidWorks by pressing (`Alt + F11`).
   - In the Project Explorer, locate the `BColors` project.
   - Right-click on `Forms` and select **Insert** > **UserForm**.
     - Rename the newly created form to `Form_BGColors`.
     - Design the form to include:
       - A Checkbox to enable or disable the gradient background (`CheckBoxEnableGradient`).
       - Buttons to open color dialogues for top and bottom gradient colors (`CommandTopColor`, `CommandBottomColor`).
       - A button to set the viewport background color (`CommandDrawColor`).
       - A `Close` button to exit the macro.

2. **Implement the Module**:
   - Right-click on `Modules` within the `BColors` project.
   - Select **Insert** > **Module**.
   - Add VBA code to this module (`BGColors1`) to handle the color selection logic and interface with the SolidWorks system settings.

3. **Configure Event Handlers and Functions**:
   - In `Form_BGColors`, implement event handlers for color selection and checkbox interactions.
   - Use `ShowColor` function to integrate with the common dialog color picker.

4. **Save and Run the Macro**:
   - Save the macro file (e.g., `BackgroundColors.swp`).
   - Run the macro by navigating to **Tools** > **Macro** > **Run** in SolidWorks, then select your saved macro.

5. **Using the Macro**:
   - The macro will open the `Form_BGColors`.
   - Toggle the gradient background option and select colors as needed.
   - Use the `Close` button to save settings and exit the macro.

## VBA Macro Code

```vbnet
'------------------------------------------------------------------------------
' BackgroundColors.swp
'------------------------------------------------------------------------------
Global swApp As Object
Global Document As Object
Global Const swSystemColorsViewportBackground = 99
Global Const swSystemColorsTopGradientColor = 100
Global Const swSystemColorsBottomGradientColor = 101
Global Const swColorsGradientPartBackground = 68

Sub Main()

  Set swApp = CreateObject("SldWorks.Application")
  Set Document = swApp.ActiveDoc
  Form_BGColors.Show

End Sub
```

## VBA UserForm Code

```vbnet
'------------------------------------------------------------------------------
' BackgroundColors.swp
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
' Type definition used to select color.
'------------------------------------------------------------------------------
Private Type CHOOSECOLOR
' Function to allow change part color
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  rgbResult As Long
  lpCustColors As String
  flags As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" _
  Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long

'------------------------------------------------------------------------------
' Function used to select color.
'------------------------------------------------------------------------------
Function ShowColor(CustomColor) As Long
  Dim cc As CHOOSECOLOR
  Dim CustColor(16) As Long
  cc.lStructSize = Len(cc)
' cc.hwndOwner = Form1.hWnd
' cc.hInstance = App.hInstance
  cc.flags = 0
' NOTE:  We can't pass an array of pointers so
' we fake this passing a string of chars:  In this
' example we set all custom colors to 0, or black.
  cc.lpCustColors = String$(16 * 4, 0)
  cc.rgbResult = CustomColor
  NewColor = CHOOSECOLOR(cc)
  If NewColor <> 0 Then
    ShowColor = cc.rgbResult
    CustomColors = StrConv(cc.lpCustColors, vbFromUnicode)
  Else
    ShowColor = -1
  End If
End Function

'------------------------------------------------------------------------------
' Allows user to enable/disable gradient background and set top/bottom colors.
'------------------------------------------------------------------------------
Private Sub CheckBoxEnableGradient_Click()
  swApp.SetUserPreferenceToggle swColorsGradientPartBackground, _
        CheckBoxEnableGradient
  If CheckBoxEnableGradient = True Then
    LabelTopColor.Enabled = True
    CommandTopColor.Enabled = True
    LabelBottomColor.Enabled = True
    CommandBottomColor.Enabled = True
    FrameModels.Caption = "Gradient Background - SolidWorks Models"
    FrameDrawing.Caption = "Background - SolidWorks Drawings"
    DisplayColors
  Else
    LabelTopColor.Enabled = False
    CommandTopColor.Enabled = False
    CommandTopColor.BackColor = vbButtonFace
    LabelBottomColor.Enabled = False
    CommandBottomColor.Enabled = False
    CommandBottomColor.BackColor = vbButtonFace
    FrameModels.Caption = "Gradient Background"
    FrameDrawing.Caption = "Background - SolidWorks Models and Drawings"
  End If
  If Not Document Is Nothing Then
    Document.GraphicsRedraw
  End If
End Sub

'------------------------------------------------------------------------------
' Set top gradient background color.
'------------------------------------------------------------------------------
Private Sub CommandTopColor_Click()
  SC = ShowColor(swApp.GetUserPreferenceIntegerValue(swSystemColorsTopGradientColor))
  If SC >= 0 Then
    HSC = "000000" + Hex(SC)
    HSC = Mid(HSC, Len(HSC) - 5, 6)
    swApp.SetUserPreferenceIntegerValue swSystemColorsTopGradientColor, SC
    DisplayColors
  End If
End Sub

'------------------------------------------------------------------------------
' Set bottom gradient background color.
'------------------------------------------------------------------------------
Private Sub CommandBottomColor_Click()
  SC = ShowColor(swApp.GetUserPreferenceIntegerValue(swSystemColorsBottomGradientColor))
  If SC >= 0 Then
    HSC = "000000" + Hex(SC)
    HSC = Mid(HSC, Len(HSC) - 5, 6)
    swApp.SetUserPreferenceIntegerValue swSystemColorsBottomGradientColor, SC
    DisplayColors
  End If
End Sub

'------------------------------------------------------------------------------
' Set viewport background color.
'------------------------------------------------------------------------------
Private Sub CommandDrawColor_Click()
  SC = ShowColor(swApp.GetUserPreferenceIntegerValue(swSystemColorsViewportBackground))
  If SC >= 0 Then
    HSC = "000000" + Hex(SC)
    HSC = Mid(HSC, Len(HSC) - 5, 6)
    swApp.SetUserPreferenceIntegerValue swSystemColorsViewportBackground, SC
    DisplayColors
  End If
End Sub

'------------------------------------------------------------------------------
' Initialize user form.
'------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
  Form_BGColors.Caption = Form_BGColors.Caption + " " + Version
  DisplayColors
  CheckBoxEnableGradient = _
        swApp.GetUserPreferenceToggle(swColorsGradientPartBackground)
  CheckBoxEnableGradient_Click
End Sub

'------------------------------------------------------------------------------
' Display current color settings.
'------------------------------------------------------------------------------
Private Sub DisplayColors()
  CommandTopColor.BackColor = _
        swApp.GetUserPreferenceIntegerValue(swSystemColorsTopGradientColor)
  CommandBottomColor.BackColor = _
        swApp.GetUserPreferenceIntegerValue(swSystemColorsBottomGradientColor)
  CommandDrawColor.BackColor = _
        swApp.GetUserPreferenceIntegerValue(swSystemColorsViewportBackground)
  If Not Document Is Nothing Then
    Document.GraphicsRedraw
  End If
End Sub

'------------------------------------------------------------------------------
' Close and end macro.
'------------------------------------------------------------------------------
Private Sub CommandClose_Click()
  End
End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).