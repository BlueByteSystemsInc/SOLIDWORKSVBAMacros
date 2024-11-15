# Hide / Show All Notes in Drawing Document

## Description
This macro automates the process of hiding or showing annotations in a SolidWorks drawing. It begins by checking if a document is open and if the active document is a drawing. If not, it prompts the user to open a drawing. Once a valid drawing is open, the macro presents a message box asking the user whether they want to hide or show annotations. Based on the user's choice, the macro loops through all views in the drawing, processing each one to either hide or display annotations of the "Note" type. After processing all views, the drawing is redrawn to reflect the changes. The macro consists of two subroutines: one for hiding and one for showing annotations, which it calls depending on the user's input.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 10 or later

## Pre-Conditions
> [!NOTE]
> - SolidWorks must be installed and running on the machine.
> - An active drawing is open.

## Post-Conditions
> [!NOTE]
> - The macro will hide or show all notes in the drawing based on the user selection
> 

 
## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Option Explicit

Sub main()

    ' Declare variables for SolidWorks application, model, drawing, view, and a boolean return value
    Dim swApp                       As SldWorks.SldWorks
    Dim swModel                     As SldWorks.ModelDoc2
    Dim swDraw                      As SldWorks.DrawingDoc
    Dim swView                      As SldWorks.View
    Dim bRet                        As Boolean

    ' Get the SolidWorks application instance
    Set swApp = Application.SldWorks
    
    ' Get the active document (drawing/model) from SolidWorks
    Set swModel = swApp.ActiveDoc
   
    ' Check if there is an active document loaded
    If swModel Is Nothing Then
    
        ' Display a message if no document is open
        swApp.SendMsgToUser ("No document loaded, please open a drawing")

        ' Exit the macro if no document is found
        Exit Sub

    End If
    
    ' Check if the active document is not a drawing
    If (swModel.GetType <> swDocDRAWING) Then

        ' Display a message if the document is not a drawing
        swApp.SendMsgToUser ("This is not a drawing, please open a drawing")
    
    Else
        ' If the document is a drawing, set it as the drawing document object
        Set swDraw = swModel
          
        ' Get the first view in the drawing
        Set swView = swDraw.GetFirstView
        
        ' Prompt the user with a message box to hide or show annotations
        nResponse = MsgBox("Hide Annotations (Yes = Hide; No = Show)?", vbYesNo)

        ' Loop through each view in the drawing
        Do While Not Nothing Is swView
           
            ' If user selects 'Yes', hide annotations
            If nResponse = vbYes Then
                ProcessDrawing1 swApp, swDraw, swView
            Else
                ' If user selects 'No', show annotations
                ProcessDrawing swApp, swDraw, swView
            End If

            ' Move to the next view in the drawing
            Set swView = swView.GetNextView

        Loop

        ' Redraw the document to reflect the changes
        swModel.GraphicsRedraw2
    
    End If

End Sub

' Subroutine to process and show annotations in a drawing view
Sub ProcessDrawing(swApp As SldWorks.SldWorks, swModel As SldWorks.ModelDoc2, swView As SldWorks.View)
    
    ' Declare variable for an annotation in the drawing view
    Dim swAnn As SldWorks.Annotation
    
    ' Get the first annotation in the current view
    Set swAnn = swView.GetFirstAnnotation2

    ' Loop through each annotation in the view
    Do While Not Nothing Is swAnn

        ' Check if the annotation is of the Note type
        If swNote = swAnn.GetType Then
            ' Make the annotation visible
            swAnn.Visible = swAnnotationVisible
        End If

        ' Move to the next annotation
        Set swAnn = swAnn.GetNext2

    Loop

End Sub

' Subroutine to process and hide annotations in a drawing view
Sub ProcessDrawing1(swApp As SldWorks.SldWorks, swModel As SldWorks.ModelDoc2, swView As SldWorks.View)
    
    ' Declare variable for an annotation in the drawing view
    Dim swAnn As SldWorks.Annotation
    
    ' Get the first annotation in the current view
    Set swAnn = swView.GetFirstAnnotation2

    ' Loop through each annotation in the view
    Do While Not Nothing Is swAnn

        ' Check if the annotation is of the Note type
        If swNote = swAnn.GetType Then
            ' Hide the annotation
            swAnn.Visible = swAnnotationHidden
        End If

        ' Move to the next annotation
        Set swAnn = swAnn.GetNext2

    Loop

End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).