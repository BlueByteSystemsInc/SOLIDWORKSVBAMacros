# Suppress New Features and Mates in All Configurations Except the Active Configuration

## Description
This macro suppresses all newly added features and mates in all configurations of the active model, except for the currently active configuration. This allows new features or mates to be added only to the active configuration, ensuring that they are suppressed in other configurations by default.

## System Requirements
- **SolidWorks Version**: SolidWorks 2014 or newer
- **Operating System**: Windows 7 or later

## Pre-Conditions
> [!NOTE]
> - The active document must be a part or assembly.
> - Ensure that you have multiple configurations created in the active document.
> - This macro does not work for drawing files.

## Results
> [!NOTE]
> - New features and mates will only be unsuppressed in the active configuration.
> - All other configurations will suppress the newly added features and mates.
> - A message will be displayed upon completion, confirming the configuration is ready for new features.

## VBA Macro Code

```vbnet
' Disclaimer:
' The code provided should be used at your own risk.  
' Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code.  
' For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

' ****************************************************************************** 
'             Set New Features And Mates For This Configuration Only             
' ****************************************************************************** 

Dim swApp           As SldWorks.SldWorks
Dim swModel         As SldWorks.ModelDoc2
Dim oConfigs        As Variant
Dim sCurrentConfig  As String
Dim sConfigComment  As String
Dim sConfigAltName  As String
Dim swConfig        As SldWorks.Configuration
Dim i               As Integer
Dim bRet            As Boolean

Sub main()

    ' Initialize the SolidWorks application
    Set swApp = Application.SldWorks
    
    ' Exit if no documents are open
    If swApp.GetDocumentCount() = 0 Then Exit Sub

    ' Get the active document (part or assembly)
    Set swModel = swApp.ActiveDoc

    ' Exit if the document is a drawing
    If swModel.GetType() = swDocumentTypes_e.swDocDRAWING Then Exit Sub

    ' Get the active configuration
    Set swConfig = swModel.GetActiveConfiguration
    sCurrentConfig = swConfig.Name

    ' Retrieve all configuration names
    oConfigs = swModel.GetConfigurationNames

    ' Loop through each configuration
    For i = 0 To UBound(oConfigs)

        Set swConfig = swModel.GetConfigurationByName(oConfigs(i))

        ' Check if configuration exists
        If Not swConfig Is Nothing Then
            sConfigComment = swConfig.Comment
            sConfigAltName = swConfig.AlternateName

            ' If it's the current active configuration, set the new features to be unsuppressed
            If swConfig.Name = sCurrentConfig Then
                bRet = swModel.EditConfiguration3(swConfig.Name, swConfig.Name, sConfigComment, sConfigAltName, 32)
            Else
                ' For all other configurations, suppress new features and mates by default
                bRet = swModel.EditConfiguration3(swConfig.Name, swConfig.Name, sConfigComment, sConfigAltName, swConfigurationOptions2_e.swConfigOption_SuppressByDefault)
            End If
        End If
    Next i

    ' Rebuild the model to apply changes
    swModel.ForceRebuild3 (False)

    ' Notify the user that the operation is complete
    MsgBox ("This Configuration Is Now Ready For New Features")

End Sub
```

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).