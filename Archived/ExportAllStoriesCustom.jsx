//ExportAllStories.jsx
//An InDesign JavaScript
/*  
@@@BUILDINFO@@@ "ExportAllStories.jsx" 3.0.0 15 December 2009
*/
//Exports all stories in an InDesign document in a specified text format.
//
//For more on InDesign/InCopy scripting see the documentation included in the Scripting SDK 
//available at http://www.adobe.com/devnet/indesign/sdk.html
//or visit the InDesign Scripting User to User forum at http://www.adobeforums.com
//
main();
function main(){
	//Make certain that user interaction (display of dialogs, etc.) is turned on.
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
	if(app.documents.length != 0){
		if (app.activeDocument.stories.length != 0){
			myDisplayDialog();
		}
		else{
			alert("The document does not contain any text. Please open a document containing text and try again.");
		}
	}
	else{
		alert("No documents are open. Please open a document and try again.");
	}
}
function myDisplayDialog(){
	with(myDialog = app.dialogs.add({name:"ExportAllStories"})){
		//Add a dialog column.
		myDialogColumn = dialogColumns.add()	
		with(myDialogColumn){
			with(borderPanels.add()){
				staticTexts.add({staticLabel:"Export as:"});
				with(myExportFormatButtons = radiobuttonGroups.add()){
					radiobuttonControls.add({staticLabel:"Text Only", checkedState:true});
					radiobuttonControls.add({staticLabel:"RTF"});
					radiobuttonControls.add({staticLabel:"InDesign Tagged Text"});
                      radiobuttonControls.add({staticLabel:"CSV"});
				}
			}
		}
		myReturn = myDialog.show();
		if (myReturn == true){
			//Get the values from the dialog box.
			myExportFormat = myExportFormatButtons.selectedButton;
			myDialog.destroy;
			myFolder= Folder.selectDialog ("Choose a Folder");
			if((myFolder != null)&&(app.activeDocument.stories.length !=0)){
				myExportAllStories(myExportFormat, myFolder);
			}
		}
		else{
			myDialog.destroy();
		}
	}
}
//myExportStories function takes care of exporting the stories.
//myExportFormat is a number from 0-2, where 0 = text only, 1 = rtf, and 3 = tagged text.
//myFolder is a reference to the folder in which you want to save your files.
function myExportAllStories(myExportFormat, myFolder){
    var myStoryText = "";
    //myStory = app.activeDocument.stories.firstItem();
	for(myCounter = 0; myCounter < app.activeDocument.stories.length; myCounter++){
        
         myStory = app.activeDocument.stories.item(myCounter).contents.replace(/(\r\n|\n|\r)/gm,"@^").replace("\,","#%");
         
         //var sideBar = myStory.split("@^@^");
         //if(sideBar.length != 0){
             //myStory = sideBar(0) + "\"\,";
             //for (sbCounter = 0; sbCounter < sideBar.length - 1; cbCounter++){
             //myStory = myStory + "\"" + sideBar(sbCounter) + "\"\,";
             //;
         //myStory = myStory.replace(("Size@^"),"").replace("Services@^","\,\"").replace("Client@^","\","").replace("Awards@^","\",)
		myStoryText =  myStoryText + "\"" + myStory + "\"\,";
		//myID = app.activeDocument.stories.firstItem().id;
        
        //
		switch(myExportFormat){
			case 0:
				myFormat = ExportFormat.textType;
				myExtension = ".txt"
				break;
			case 1:
				myFormat = ExportFormat.RTF;
				myExtension = ".rtf"
				break;
			case 2:
				myFormat = ExportFormat.taggedText;
				myExtension = ".txt"
				break;
              case 3:
                  myFormat = ExportFormat.textType;
                  myExtension = ".csv"
                  break;
		}
		
	}
         myFileName = "StoryID"+ myExtension;
         //myStoryText = myStoryText + "," + myStory.toSource();
		myFilePath = myFolder + "/" + myFileName;
		var myFile = new File(myFilePath);
		writeFile(myFile, myStoryText);
}

function toCSV(text) {
    return text.replace(/(\r\n|\n|\r)/gm,"@^")
    }

function writeFile(fileObj, fileContent, encoding) {  
    encoding = encoding || "ASCII";  
    fileObj = (fileObj instanceof File) ? fileObj : new File(fileObj);  
  
  
    var parentFolder = fileObj.parent;  
    if (!parentFolder.exists && !parentFolder.create())  
        throw new Error("Cannot create file in path " + fileObj.fsName);  
  
  
    fileObj.encoding = encoding;  
    fileObj.open("w");  
    fileObj.write(fileContent);  
    fileObj.close();  
  
  
    return fileObj;  
}  