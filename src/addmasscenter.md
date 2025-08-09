# Add Center of Mass to a Part in SOLIDWORKS


<img src="../images/AddMassCenter.png" alt="Description of image" width="600" style="display: block; margin: 0 auto;">

## Macro Description

This VBA macro is designed to automatically add a center of mass (COM) point to a part in SOLIDWORKS. The macro calculates the center of mass of the part and inserts a point at that location, making it easier to analyze the balance and mass distribution of the part. This functionality is especially useful for engineers and designers who frequently work with parts requiring balance and stability assessments.

## VBA Macro Code


```vbnet
'The code provided is for educational purposes only and should be used at your own risk. 
'Blue Byte Systems Inc. assumes no responsibility for any issues or damages that may arise from using or modifying this code. 
'For more information, visit [Blue Byte Systems Inc.](https://bluebyte.biz).

Dim swApp As SldWorks.SldWorks
Dim swModelDoc As SldWorks.ModelDoc2
Dim swCenterMass As SldWorks.Feature
Dim swCenterMassReferencePoint As SldWorks.Feature

Option Explicit

Sub main()

    Set swApp = Application.SldWorks
    Set swModelDoc = swApp.ActiveDoc

    Set swCenterMass = swModelDoc.FeatureManager.InsertCenterOfMass
    Set swCenterMassReferencePoint = swModelDoc.FeatureManager.InsertCenterOfMassReferencePoint
   

End Sub

```

## System Requirements
To run this VBA macro, ensure that your system meets the following requirements:

- **SOLIDWORKS Version**: SOLIDWORKS 2017 or later
- **VBA Environment**: Pre-installed with SOLIDWORKS (Access via Tools > Macro > New or Edit)
- **Operating System**: Windows 7, 8, 10, or later
- **Additional Libraries**: None required (uses standard SOLIDWORKS API references)

> [!NOTE]
> Pre-conditions 
>- The active document must be a part (`.sldprt`) in SOLIDWORKS.
>- The part should have a valid material assigned to ensure the correct center of mass is calculated.
>- The part must not be empty (should contain geometry).

> [!NOTE]
> Post-conditions
>- A center of mass point will be inserted into the part.


## Macro
You can download the macro from [here](../images/AddMassCenter.swp)

## Customization
Need to modify the macro to meet specific requirements or integrate it with other processes? We provide custom macro development tailored to your needs. [Contact us](https://bluebyte.biz/contact).