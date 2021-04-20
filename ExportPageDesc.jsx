//scrapePages: worked on during August - September 2019
//Scrapes data from Adobe InDesign Document in one of various standard firm layouts for output to loadable CSV

main();
function main(){
	
	//Make certain that user interaction (display of dialogs, etc.) is turned on. Turn on to see broken links
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
    //app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
	if(app.documents.length != 0){
		if (app.activeDocument.stories.length != 0){
			//myDisplayDialog();
              myFolder = ("M:/Enterprise_Stds/ProjectPages")
              expFormat = ".txt"          
              myExportDesc(expFormat, myFolder)
              $.gc()
		}
    
		else{
			alert("The document does not contain any text. Please open a document containing text and try again.");
		}
	}
	else{
		alert("No documents are open. Please open a document and try again.");
	}
}

function myExportDesc(myExportFormat, myFolder){
    var curDate = new Date();
    myFileName = "ak.DataMine_" + curDate.toDateString()+ myExportFormat;
    myFilePath = myFolder + "/" + myFileName;
    var myFile = new File(myFilePath);
    
//resolves the lack of indexOf function   
if (typeof Array.prototype.indexOf != "function") {  
    Array.prototype.indexOf = function (el) {  
        for(var i = 0; i < this.length; i++) if(el === this[i]) return i;  
        return -1;  
        }  
}  

    
    var linkBucket = [{}];
    linkBucket.pop();
    
    var pageRow = [];
    //fill row with null values at first
    for(pR = 0; pR < 16 ; pR++) pageRow[pR] = ("\"" + "\"");
    
    largest = "";

    //file information 
    pageRow[0] = csvQuotes(app.activeDocument.fullName.toString().replace("/m/","//d-peapcny.net/enterprise/M_Marketing/"));
    pageRow[1] = csvQuotes(File(app.activeDocument.fullName).modified);

    var fnArr = app.activeDocument.name.split("_");
    pageRow[2] = ("\"" + app.activeDocument.filePath + "\"")
    pageRow[3] = ("\"" + fnArr[0] + "\"");
    pageRow[4] = ("\"" + fnArr[1] + "\"");
    
    mainLayer = app.activeDocument.activeLayer;
    
    myLinks = app.activeDocument.links;

    for (w=0;w<myLinks.length;w++) if (myLinks[w].parent.parentPage !== null && myLinks[w].parent.itemLayer == mainLayer) linkBucket.push(myLinks[w]);
    for (z = 0;z<linkBucket.length;z++) pageRow[11+z] = csvQuotes(linkBucket[z].name);   
    
    for(myCounter = 0; myCounter < app.activeDocument.stories.length; myCounter++){
         
         myStory = app.activeDocument.stories.item(myCounter);
                  
         myStoryText = csvQuotes(csvFriendly(myStory.contents));
         
         if (myStory.paragraphs.firstItem().isValid == true) myStoryStyle = myStory.paragraphs.firstItem().appliedParagraphStyle.name;
                       
         if (myStoryStyle == "Project Name"){
             var nameArr = new String();
             nameBox = csvFriendly(myStory.contents);
             nameArr = nameBox.split("</p><p>");      
             pageRow[6] = "\"" + nameArr[nameArr.length-1] + "\"";
             nameArr.pop();
             pageRow[5] = "\"" + nameArr.toString().replace("\,"," ") + "\"";
             }
         
//~          //parses Sidebar for size
       if (myStoryStyle == "Project Details"){
             pageRow[7] = parseSidebar(myStory,"Size","Services");
             pageRow[8] = parseSidebar(myStory,"Services","Client");
            }
         
//~          //parses for Description text & compares the length of the story to see if it is the longest string represented so far    
         if (myStoryStyle == "Body Text" && myStoryText.length > largest.length){
             largest = myStoryText;
             pageRow[9] = myStoryText;
             }
         
         if (myStoryStyle == "Project Page Callout") pageRow[10] = myStoryText;
         
        

    }
    //myPageText = pageRow.toString() + "\," +   linkArr.toString() + "\n";
    myPageText = pageRow.toString() + "\n";
    //alert(pageRow.toString());
    //myStoryText = myStoryText +"\"" + app.activeDocument.name + "\"\," + "\"" + myBounds + "\"\," + "\"" + largest + "\"\, \n";
    writeFile(myFile, myPageText);
}

function writeFile(fileObj, fileContent, encoding) {  
    encoding = encoding || "UTF-8";  
    var titleRow = [        csvQuotes("Link"),csvQuotes("Modified"),
                            csvQuotes("Path"),csvQuotes("FileName"),csvQuotes("ProjectNo"),
                            csvQuotes("Title"),csvQuotes("Location"),csvQuotes("SquareFootage"),
                            csvQuotes("Services"),csvQuotes("ProjectDescription"),csvQuotes("CallOut"),
                            csvQuotes("Link1"),csvQuotes("Link2"),csvQuotes("Link3"),
                            csvQuotes("Link4"),csvQuotes("Link5"),csvQuotes("Link6")];
                            
    if (!fileObj.exists) fileContent2 = titleRow.toString() + "\n" + fileContent;
    else fileContent2 = fileContent;
    
   //alert(fileContent2)
    
    fileObj = (fileObj instanceof File) ? fileObj : new File(fileObj);
  
    var parentFolder = fileObj.parent;  
    if (!parentFolder.exists && !parentFolder.create()) throw new Error("Cannot create file in path " + fileObj.fsName);  
  
    fileObj.encoding = encoding;  
    fileObj.open("a");  
    fileObj.write(fileContent2);  
    fileObj.close();  
  
    return fileObj;  
}  

//function to search to
function parseSidebar(myStory,Start,Finish){
    var sizeArr = [""];
    refString = csvFriendly(myStory.contents);
    var refArr = refString.split("</p><p>");
    sizeReturn = "";
    
    //loads paragraphs into string reference array
//~     for(myCounter = 0; myCounter < myStory.paragraphs.length ; myCounter++){
//~         sbLine = myStory.paragraphs.item(myCounter).contents.replace(/(\r\n|\n|\r)/gm,"").replace("\,","#%");
//~         refArr.push(sbLine);
//~         }
//~     
    //loads sizeArr with relevant data after size line
    sizeArr = refArr.slice(refArr.indexOf(Start)+1, refArr.indexOf(Finish));
    sizeReturn = "\"" + sizeArr.toString().replace(/,/g,";")     + "\"";
    return sizeReturn;    
    }
    
//3 functions to rework text into quotes/HTML format for the purposes of being loaded into a workable CSV
function csvQuotes(myText){
    myText = ("\"" + trim(myText) + "\"");
    return myText
}

//replaces story paragraph breaks with </p><p> and commas with #% --DO NOT FORGET TO REPLACE THEM
function csvFriendly(myText){
    myText = trim(myText.toString().replace(/(\r\n|\n|\r)/gm,"</p><p>").replace(/,/g,"#%"));
    return myText;
 }
 
 function trim(str) {
    return str.toString().replace(/^\s\s*/, '').replace(/\s\s*$/, '');
}

function logMe(input){
     var now = new Date();
     var output = now.toTimeString() + ": " + input;
     $.writeln(output);
     var logFile = File("/path/to/logfile.txt");
     logFile.open("e");
     logFile.writeln(output);
     logFile.close();

}