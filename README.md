# MSI Template PS

Do you need to add over and over again the same properties or summary information in your MSIs/MSTs?

With this tool you can achieve that. And it's all done in Powershell.

## How to configure the template?

During the first run, an XML template is created in %appdata% of the user called ModifyMsiXML.XML

This is automatically populated with some summary information and properties.

You can change the properties in the left pane of the app, under **Configuration**.

Click the **+** button to expand the tree.

**SumInfo** represents what you can customize in the summary information stream.

**Property** represents what properties you can add.

After you make the desired changes click **Save Template**.

## How to apply it to MSI or MST?

There are 3 scenarios:

1. Change only an MSI file. In this case, in the first textbox, select you desired MSI and press **Modify**.
1. Create a transform (MST) for a desired MSI. In this case, in the first textbox, select the desired MSI and check the checkbox bellow **Create MST**. After that click **Modify**
1. Change a transform file. In this case, select your MSI in the first texbox, and select the MSI in the secont textbox. At the end press **Modify**

## Build

In the **compiled** directory you can find the MSITemplatePS.exe

The ps1 has been compiled to exe using PS2EXE-GUI

## Known Issues

* When you change an MST or create an MST, the copy MSI needed in order to create the transform and MSTs are in use with the program until is closed. Because of this, those files need to be deleted manually afterwards.


## Disclaimer

Use this program at your own risk. I will not be held responsable for any issues that this will do to you MSIs.