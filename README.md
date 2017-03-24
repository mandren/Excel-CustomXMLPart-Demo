# Excel-CustomXMLPart-Demo 
This sample add-in demonstrates how you can leverage the new Custom XML Part JavaScript API for Microsoft Excel 2016 to create a compelling Excel add-in. With this add-in, you will learn how to read and write data from / to a custom XML part contained in an Excel workbook, and use that information within Excel as part of your application.  This demo also illustrates how to leverage bindings to expose selection changed events within Excel to your add-in, and further extend the capabilities of your user interface.
## Table of Contents
* [Prerequisites](#prerequisites)
* [Run the project](#run-the-project)

## Prerequisites
You'll need the following:
* [Visual Studio 2017](https://www.visualstudio.com/downloads/download-visual-studio-vs.aspx)
* [Office Developer Tools for Visual Studio](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)
* Excel 2016, version 1703 or later
## Run the project
1.	Copy the project to a local folder. Ensure that the file path is not too long, otherwise you might run into an error in Visual Studio when it tries to install the NuGet packages necessary for the project.
2.	Then open the Excel-CustomXMLPart-Demo.sln in Visual Studio.
3.	Press F5 to build and deploy the sample add-in. Excel launches and will open the test workbook included as part of the solution (TestBook.xlsx).
4.	Click the Show Taskpane button on the ribbon (It should be the last button on the Home Tab).  This will display the add-in’s taskpane.
5.	Click the Load XML button to extract the data from the custom XML part in the workbook.  When that task completes, you will see a banner at the bottom of the task pane indicating the number on binding objects that were found in the XML.
6.	Click the Load Workbook button to hydrate the workbook with the bindings from the loaded XML
7.	Selecting different cells on Sheet1 at this point will update the taskpane UI appropriately.
8.	Select a blank cell and add a binding by selecting the binding type from the drop down and clicking the Add Binding button.
9.	Repeat steps 7 & 8 to add additional bindings as desired.
10.	 Click the Serialize Data button to write the data back into the XML Part, and clean up the content in the workbook.
11.	Repeat steps 5 – 7 to see the updated XML reflected in the Workbook. (You can also save the workbook at this point, and the changes to the CustomXML part will be persisted to disk for the next time you run the application.


