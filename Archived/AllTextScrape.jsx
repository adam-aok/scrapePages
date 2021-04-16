﻿main();

//basic declaration of 
function main(){
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
    //app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
    if(app.documents.length != 0){
		if (app.activeDocument.stories.length != 0){
              myFolder = ("U:/")
              expFormat = ".txt"          
              myExportPages(expFormat, myFolder)
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

function myExportPages(myExportFormat, myFolder){
    
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
    //var basicMasters = ["A-Master","B.1-Project Team (Headshots)","B.2-Project Team (Group Photo)","B.3-Project Details","C.1-P
      
    var pageRow = [];
    
    //fill row with null values at first
    for(pR = 0; pR < 3 ; pR++) pageRow[pR] = ("\"" + "\"");
        
            
    //input filepath into pageRow
    
    
    //q4 2019 locations array
//~     var locArr1 = [1.638,2.3902,3.1502,3.8978,4.651,5.4052,6.1586,6.9531]
//~     //option 2 locations array
//~     var locArr2 = [1.638,2.2874,2.9445,3.5894,4.2398,4.8912,5.5442,6.2071,2.04]
//~     //option 3 locations array
//~     var locArr3 = [1.635,2.2744,2.9215,3.5564,4.1968,4.8382,5.4793,5.5951]
//~     var phArr = ["•	Predicted Energy Use Intensity- pEUI (kBTU/sf/yr):  Enter Value (if applicable)</P><P>•	Energy Savings (% reduction from the baseline):  Enter Value</P><P>","Type your Body Text here", "Type your Body Text here</P><P>Type your Body Text here</P><P>Type your Body Text here</P><P>Type your Body Text here</P><P>","Briefly describe sustainability goals and strategies (i.e.  Sustainability Strategies Checklist)</P><P>Type your text here"]
    
             
    var titlesIndex = [];
    var fileArr = ["\"" + "\"","\"" + "\"","\"" + "\""]
    //for(myCounter = 0; myCounter < app.activeDocument.pages.length; myCounter++)
    var posList = [];
    
    posList = locArr2;
    mSpread = app.activeDocument.masterSpreads.item(0);
    
//~     for(a=0;a<mSpread.textFrames.length;a++){
//~         if (mSpread.textFrames.item(a).parentStory.contents == "DESIGN ON THE BOARDS      Q-4 2019") posList= locArr1;
//~         if (mSpread.textFrames.item(a).parentStory.contents == "DESIGN ON THE BOARDS      Q-4 2017") posList = locArr3;
//~         }
        
    var projIndex = [[]];
    projIndex.pop();
    
    //alert(File(app.activeDocument.fullName).modified);
    fileArr[0] = csvQuotes(app.activeDocument.fullName.toString().replace("/g/","G:/"));
    fileArr[1] = csvQuotes(File(app.activeDocument.fullName).modified);
         
        
    //iterating through spreads of document (spreads is chosen instead of pages to accommodate spreads containing common information. open to modification to spreads)
    for(myCounter = 0; myCounter < app.activeDocument.pages.length; myCounter++){
         
         //get current page
         myPage = app.activeDocument.pages.item(myCounter);         
        
        if (myPage.appliedMaster !== null) masterCheck = myPage.appliedMaster;
        else var masterCheck = app.activeDocument.masterSpreads.item(0);
        
         modify = false;
         pageRow = [];
         for(pR = 0; pR < 11 ; pR++) pageRow[pR] = ("\"" + "\"");
         margin = 0;
         largest = "";
         masterModified = false;
                  
         for(z=0;z<myPage.groups.length;z++){
                checkGroup = myPage.groups.item(z);
                checkGroup.ungroup(); 
            }


//~          for(z=0;z<myPage.textFrames.length;z++){
//~                 checkTextFrame = myPage.textFrames.item(z);
//~                 checkTitle = csvQuotes(csvFriendly(checkTextFrame.parentStory.contents.toUpperCase()));
//~                 if (checkTextFrame.parentStory.contents == "Location" && checkTextFrame.geometricBounds[0] !== 1.44) margin = checkTextFrame.geometricBounds[0] - 1.44;
//~                 if (checkTextFrame.parentStory.contents == "Practice Area" && checkTextFrame.geometricBounds[0] !== 1.44) margin = checkTextFrame.geometricBounds[0] - 1.44;
//~                 if (titlesIndex.length !== 0 && titlesIndex[titlesIndex.length-1] == checkTitle && approx(checkTextFrame.geometricBounds[0],.609) == true){                
//~                     pageRow = projIndex[projIndex.length - 1];
//~                     modify = true;
//~                 }
//~             }
//~         
//~         for(n = 0; n < masterCheck.textFrames.length;n++){
//~             mTextFrame = masterCheck.textFrames.item(n)
//~                 if (approx(mTextFrame.geometricBounds[0],.609) == true && mTextFrame.parentStory.isValid == true){
//~                             //&& approx(mTextFrame.geometricBounds[1],6.3766) == true 
//~                              pageTitle = csvQuotes(mTextFrame.parentStory.contents.toUpperCase());
//~                              if (pageTitle !== csvQuotes("CONFIDENTIAL PROJECT NAME") && pageTitle !== csvQuotes("PROJECT NAME")){
//~                              masterModified = true;
//~                              largest = "";
//~                              //checks if the title of this page is already in the titlesIndex. if so, it sets the modify value to false.
//~                              if (titlesIndex.indexOf(pageTitle) == -1){
//~                                 titlesIndex.push(pageTitle);
//~                                 modify = false;
//~                                  }
//~                              else if (titlesIndex.indexOf(pageTitle) !== -1){
//~                                  //pageRow = projIndex[titlesIndex.indexOf(pageTitle)];
//~                                  pageRow = projIndex[projIndex.length-1];
//~                                  modify = true;
//~                                  }                             
//~                              pageRow[0] = pageTitle;
//~                              }
//~                          }
//~                      }     

//~          if (masterCheck.name == "D-Divider"){
//~              for(x = 0; x < masterCheck.textFrames.length;x++){
//~                  mcFrame = masterCheck.textFrames.item(x);
//~                  if (mcFrame.parentStory.isValid == true){
//~                      mcStory = csvFriendly(mcFrame.parentStory.contents.toUpperCase());
//~                      if (approx(mcFrame.geometricBounds[1],13.8646) == true && mcStory !== "OFFICE LOCATION") fileArr[2] = csvQuotes(mcStory);
//~                      }
//~                  }
//~              }
                         
                                
         for(myCount = 0; myCount < myPage.textFrames.length; myCount++){

                //get textFrame on page
                 myTextFrame = myPage.textFrames.item(myCount);
                 //alert(myTextFrame.rotationAngle);
                                                   
                 myStory = myTextFrame.parentStory;                          
                 myStoryText = csvFriendly(myStory.contents);
                                  
                 var myStoryStyle = "undefined";
                 
                 //get paragraph style for that text box
                 if (myStory.paragraphs.firstItem().isValid == true) myStoryStyle = myStory.paragraphs.firstItem().appliedParagraphStyle.name;
                 
                 //get geometric bounds of this text frame
                 myPosition = myTextFrame.geometricBounds;
          
                 if (approx(myPosition[0],.609) == true && myPosition[1] < 3){
                             pageTitle = csvQuotes(myStoryText.toUpperCase());
                             if (titlesIndex.indexOf(pageTitle) == -1) titlesIndex.push(pageTitle);
                             pageRow[0] = pageTitle;
                             }
                 //if (masterCheck.name == "B.1-Project Team (Headshots)" || masterCheck.name == "B.2-Project Team (Group Photo)"){
                                 
                 //if (masterCheck.name == "B.3-Project Details"){
                                      
                    if (approx(myPosition[0],posList[0]+margin) == true && pageRow[2] == "\"" + "\"") pageRow[2] = csvQuotes(myStoryText);
                    if (approx(myPosition[0],posList[1]+margin) == true && pageRow[3] == "\"" + "\"") pageRow[3] = csvQuotes(myStoryText);                        
                    if (approx(myPosition[0],posList[2]+margin) == true && pageRow[4] == "\"" + "\"") pageRow[4] = csvQuotes(myStoryText);
                    if (approx(myPosition[0],posList[3]+margin) == true && pageRow[5] == "\"" + "\"") pageRow[5] = csvQuotes(myStoryText);
                    if (approx(myPosition[0],posList[4]+margin) == true && pageRow[6] == "\"" + "\"") pageRow[6] = csvQuotes(myStoryText);
                    if (approx(myPosition[0],posList[5]+margin) == true && pageRow[7] == "\"" + "\"") pageRow[7] = csvQuotes(myStoryText);
                    if (approx(myPosition[0],posList[6]+margin) == true && pageRow[1] == "\"" + "\"") pageRow[1] = csvQuotes(myStoryText);
                    if (approx(myPosition[0],posList[7]+margin) == true && pageRow[8] == "\"" + "\"") pageRow[8] = csvQuotes(myStoryText);
                    if (approx(myPosition[1],8.0872) == true && approx(myPosition[0],1.6814) == true && pageRow[10] !== "\"" + "\"" && pageRow[8] == "\"" + "\"") pageRow[8] = csvQuotes(myStoryText);
                     if (myPosition[2] - myPosition[0] > 1 && myStoryText.length > 100 && myStoryText.length > largest.length && pageRow[9] == "\"" + "\""){
                        largest = myStoryText;
                        pageRow[9] = csvQuotes(myStoryText);
                        }                    
                     if (approx(myPosition[1],0.3686) == true && myStoryText.length > 40  && myStoryStyle == "Body Text" && pageRow[10] == "\"" + "\"") pageRow[10] = csvQuotes(myStoryText);
                     //}
               }
           
    
    //rerun for the masterSpread      
    //alert
    
    //checks if a title was found on this page, if not, initializes the "masterModified" variable as false, to be checked in the next step
    // if (pageRow[0] == "\"" + "\""){
         //masterCheck.textFrames.item(x).isValid == true
        //checks the master page to see if the master page has been modified, with the indication being whether the master title page is equal to the default values.
//~             if (masterModified = true){
//~                 //alert("Checking master");
//~                 for(z=0;z<masterCheck.groups.length;z++){                    
//~                 checkGroup = myPage.groups.item(z);
//~                 checkGroup.ungroup(); 
//~                 }
//~                 for(masterCount = 0; masterCount < masterCheck.textFrames.length; masterCount++){
//~                     
//~                      myTextFrame = masterCheck.textFrames.item(masterCount);
//~                      myPosition = myTextFrame.geometricBounds;
//~                      myStory = myTextFrame.parentStory;                          
//~                      myStoryText = csvFriendly(myStory.contents);
//~                      //alert(myStoryText);
//~                      
//~                      var myStoryStyle = "undefined";
//~                      if (myStory.paragraphs.firstItem().isValid == true) myStoryStyle = myStory.paragraphs.firstItem().appliedParagraphStyle.name;       
//~                     //alert(myStoryText);
//~                      if ((myStoryStyle == "Body Text" || myStoryStyle == "Body Text with Bullets") && phArr.indexOf(myStoryText) == -1) {
//~                          //alert(myStoryText);
//~                         if (approx(myPosition[0],posList[0]+margin) == true && pageRow[2] == "\"" + "\"") pageRow[2] = csvQuotes(myStoryText);
//~                         if (approx(myPosition[0],posList[1]+margin) == true && pageRow[3] == "\"" + "\"") pageRow[3] = csvQuotes(myStoryText);                        
//~                         if (approx(myPosition[0],posList[2]+margin) == true && pageRow[4] == "\"" + "\"") pageRow[4] = csvQuotes(myStoryText);
//~                         if (approx(myPosition[0],posList[3]+margin) == true && pageRow[5] == "\"" + "\"") pageRow[5] = csvQuotes(myStoryText);
//~                         if (approx(myPosition[0],posList[4]+margin) == true && pageRow[6] == "\"" + "\"") pageRow[6] = csvQuotes(myStoryText);
//~                         if (approx(myPosition[0],posList[5]+margin) == true && pageRow[7] == "\"" + "\"") pageRow[7] = csvQuotes(myStoryText);
//~                         if (approx(myPosition[0],posList[6]+margin) == true && pageRow[1] == "\"" + "\"") pageRow[1] = csvQuotes(myStoryText);
//~                         if (approx(myPosition[0],posList[7]+margin) == true && pageRow[8] == "\"" + "\"") pageRow[8] = csvQuotes(myStoryText);
//~                         if (approx(myPosition[1],8.0872) == true && pageRow[10] !== "\"" + "\"" && pageRow[8] == "\"" + "\"") pageRow[8] = csvQuotes(myStoryText);
//~                         if (myPosition[2] - myPosition[0] > 1 && myStoryText.length > 100 && myStoryText.length > largest.length && pageRow[9] == "\"" + "\""){
//~                         largest = myStoryText;
//~                         pageRow[9] = csvQuotes(myStoryText);
//~                         } 
//~                          if (approx(myPosition[1],0.3686) == true && myStoryText.length > 40  && myStoryStyle == "Body Text" && pageRow[10] == "\"" + "\"") pageRow[10] = csvQuotes(myStoryText);
//~                          }   
//~                    
//~                }
//~            }
       //}
   //alert(pageRow[10]);
   //alert(posList + margin);
   
    if (modify == true) projIndex[projIndex.length-1] = pageRow;
    if (pageRow[0] !== "\""+ "\"" && projIndex[projIndex.length-1] !== pageRow && modify == false) projIndex.push(pageRow);
         
    
    
    
            
//~     if (titlesIndex.indexOf(pageRow[0]) == -1){
//~                      if (pageRow[0] !== "\"" + "\"" && projIndex.indexOf(pageRow) == -1){
//~                          projIndex.push(pageRow);
//~                          alert(pageRow);
//~                          }    
//~                      }
    
}
//link adding functions seems to have created an issue
//var linkArr = [];

//parses for link list
//for(myCounter = 0; myCounter < app.activeDocument.links.length; myCounter++){
   // eachLink = app.activeDocument.links.item(myCounter).filePath;
    //if (linkArr.indexOf("\"" + eachLink + "\"") == -1){
    //    linkArr.push("\"" + eachLink + "\"")
    //}
    
    for (j = 0; j< projIndex.length;j++){
        //alert(projIndex[j].toString())
        myPageText = fileArr.toString() +"\," + projIndex[j].toString() + "\n";
        //alert(myPageText);
        //alert(titleRow.toString());
        writeFile(myFile,myPageText);
        }
       
   // if (pageRow[0] !=="\"" + "\"")    writeFile(myFile, myPageText);
}


//test 

function writeFile(fileObj, fileContent, encoding) {  
    encoding = encoding || "UTF-8";
    var titleRow = [csvQuotes("Path"),csvQuotes("Modified"),csvQuotes("Office"),csvQuotes("Project Name"),csvQuotes("Project Number"),csvQuotes("Location"),csvQuotes("Practice Area"),csvQuotes("Construction Type"),csvQuotes("Size"),csvQuotes("Estimated Cost"),csvQuotes("Estimated Completion"),csvQuotes("Sustainability"),csvQuotes("Project Description"),csvQuotes("Team")];
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

function parseSidebar(myStory,Start,Finish){
    var sizeArr = [""];
    refString = csvFriendly(myStory.contents);
    var refArr = refString.split("</P><P>");
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
//.replace("\,",";")    
    return sizeReturn;
    }


//convert text into compatible csv format--REPLACE </P><P> WITH PARAGRAPH BREAK AND #% WITH COMMA
function csvFriendly(myText){
    myText = trim(myText.toString().replace(/(\r\n|\n|\r)/gm,"</P><P>").replace(/,/g,"#%"));
    //myText = myText.replace(/(\r\n|\n|\r)/gm,"</P><P>").replace(/,/g,"#%");
    return myText;
 }
 
function approx(number,reference,delta){
    //if delta <> undefined;
    var delta = .02;
    if (Math.abs((number - reference)) <= delta || Math.abs((reference - number)) <= delta){
    return true
    }
    else return false
    }

function csvQuotes(myText){
    myText = ("\"" + trim(myText) + "\"");
    return myText
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


