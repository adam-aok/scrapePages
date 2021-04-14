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

    largest = "";    
    var pageRow = [];
    
    //fill row with null values at first
    for(pR = 0; pR < 7 ; pR++){
        pageRow[pR] = ("\"" + "\"");
        }
    
    var fnArr = app.activeDocument.name.split("_");
    pageRow[0] = ("\"" + app.activeDocument.filePath + "\"")
    pageRow[1] = ("\"" + fnArr[0] + "\"");
    pageRow[2] = ("\"" + fnArr[1] + "\"");
    
    for(myCounter = 0; myCounter < app.activeDocument.stories.length; myCounter++){
         
         myStory = app.activeDocument.stories.item(myCounter);
         
         //replaces story paragraph breaks with @^ and commas with </p><p> --DO NOT FORGET TO REPLACE THEM
         myStoryText = "\"" + csvFriendly(myStory.contents) + "\"";
         
         if (myStory.paragraphs.firstItem().isValid == true){
         myStoryStyle = myStory.paragraphs.firstItem().appliedParagraphStyle.name;
         }
     
         //myBounds = app.activeDocument.textFrames.item(myCounter).geometricBounds.toString();
              
         //parse page title, commit to row         
         if (myStoryStyle == "Project Name"){
             var nameArr = new String();
             nameBox = myStory.contents.replace(/(\r\n|\n|\r)/gm,"</p><p>").replace("\,","#%");
             nameArr = nameBox.split("</p><p>");             
             pageRow[4] = "\"" + nameArr[nameArr.length-1] + "\"";
             nameArr.pop();
             pageRow[3] = "\"" + nameArr.toString().replace("\,"," ") + "\"";
             }
         
         //parses Sidebar for size
       if (myStoryStyle == "Project Details"){
             pageRow[5] = parseSidebar(myStory,"Size","Services");
             pageRow[6] = parseSidebar(myStory,"Services","Client");
            }
         
         //parses for Description text & compares the length of the story to see if it is the longest string represented so far    
         if (myStoryStyle == "Body Text" && myStoryText.length > largest.length){
             largest = myStoryText;
             pageRow[7] = myStoryText;
             }
         
        }

//link adding functions seems to have created an issue
//var linkArr = [];

//parses for link list
//for(myCounter = 0; myCounter < app.activeDocument.links.length; myCounter++){
   // eachLink = app.activeDocument.links.item(myCounter).filePath;
    //if (linkArr.indexOf("\"" + eachLink + "\"") == -1){
    //    linkArr.push("\"" + eachLink + "\"")
    //}
    //}
    
    //myPageText = pageRow.toString() + "\," +   linkArr.toString() + "\n";
    myPageText = pageRow.toString() + "\n";    
    //myStoryText = myStoryText +"\"" + app.activeDocument.name + "\"\," + "\"" + myBounds + "\"\," + "\"" + largest + "\"\, \n";
    

         writeFile(myFile, myPageText);
}

//test 

function writeFile(fileObj, fileContent, encoding) {  
    encoding = encoding || "UTF-8";  
    var titleRow = [csvQuotes("Path"),csvQuotes("FileName"),csvQuotes("ProjectNo"),
                            csvQuotes("Title"),csvQuotes("Location"),csvQuotes("SquareFootage"),
                            csvQuotes("Services"),csvQuotes("ProjectDescription")];
                            
    if (!fileObj.exists) fileContent2 = titleRow.toString() + "\n" + fileContent;
    else fileContent2 = fileContent;
    fileObj = (fileObj instanceof File) ? fileObj : new File(fileObj);  
  
  
    var parentFolder = fileObj.parent;  
    if (!parentFolder.exists && !parentFolder.create())  
        throw new Error("Cannot create file in path " + fileObj.fsName);  
  
  
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