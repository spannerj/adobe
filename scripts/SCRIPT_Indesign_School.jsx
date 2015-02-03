#include Check_Indesign.jsx
#include utils.jsx

$.evalFile(new File("c:/Adobe/scripts/readXML.jsx"));
$.evalFile(new File("c:/Adobe/scripts/utils.jsx"));

var count = 1;
var pdfName = '';
var newjpgfolder ='';

//initialize processor (including parms)
var gP = new Processor();
var strTitle = "Process School";

//read in parameters from xml file
g_script_XMLFunctions.ReadXMLFile(gP.params);

//set up all required folders
//output folder
//internet folder



function Processor() {

    this.InitParams = function() {
		var params = new Object();
		params["version"] = 1;
		params["useopen"] = false;
		params["includesub"] = false;
		params["source"] = "";
		params["open"] = false;
		params["saveinsame"] = true;
		params["dest"] = "";
		params["jpeg"] = true;
		params["psd"] = false;
		params["tiff"] = false;
		params["lzw"] = false;
		params["converticc"] = false;
		params["q"] = 5;
         params["imagecount"] = 6;
		params["max"] = true;
		params["jpegresize"] = false;
		params["jpegw"] = "";
		params["jpegh"] = "";
		params["psdresize"] = false;
		params["psdw"] = "";
		params["psdh"] = "";
		params["tiffresize"] = false;
		params["tiffw"] = "";
		params["tiffh"] = "";
		params["runaction"] = false;
		params["actionset"] = "";
		params["action"] = "";
		params["info"] = "";
		params["icc"] = true;
		params["keepstructure"] = false;
         params["watermark"] = "Return by";
		return params;
	}
    
    this.params = this.InitParams();
}


main();

function main(){

    storedSettings = getStoredSettings();
    
    var settings = getSettings(storedSettings.watermark, storedSettings.startFolder, storedSettings.outputFolder);
   
    if ( folder_check( storedSettings.startFolder ) )
    {
        return;
    }
   
    //create new folder for jpg folders to be moved to
    var f = new Folder(settings.startFolder + '_TIF_PEFN');
    var jpgFolder = new Folder( settings.outputFolder + settings.startFolder.substring(settings.startFolder.lastIndexOf('\\') ) +  '_TIF_PEFN' );
    if (!jpgFolder.exists) jpgFolder.create();
    if (!f.exists) f.create();
    
    //create new folder for pdfs to be moved to
    outputFolder = settings.outputFolder + settings.startFolder.substring(settings.startFolder.lastIndexOf('\\'));
    var of = new Folder(outputFolder);
    if (!of.exists) of.create();   
    
    loopFolder(settings.startFolder, settings.watermark, outputFolder);

    return true
};


function loopFolder(path, watermark, output){
     // Create new folder object based on path string  
     var folder = new Folder(path);  
     
      // Get everything in the top folder  
     var folderArray = folder.getFiles(onlyFolders); 

     //progress bar
     // Create a palette-type window (a modeless or floating dialog),
	var win = new Window("palette", "Folder Progress", [150, 150, 600, 260]); 
	this.windowRef = win;
    
	// Add a panel to contain the components
	win.pnl = win.add("panel", [10, 10, 440, 100], "");

	// Add a progress bar with a label and initial value of 0, max value of 200.
	win.pnl.progBarLabel = win.pnl.add("statictext", [20, 20, 320, 35], "Processing folder 1 of " + folderArray.length);
	win.pnl.progBar = win.pnl.add("progressbar", [20, 35, 410, 60], 1, folderArray.length);
   
   // Display the window
	win.show();

     // Loop over the folders and files 
     for (var i = 0; i < folderArray.length; i++)  
     { 
         win.pnl.progBarLabel.text = "Processing folder " + (i + 1) + " of " + folderArray.length;
          //  ignore files and loop through folders recursively with the current folder as an argument  
          if (folderArray[i] instanceof Folder)  
          {  
              var folderName = folderArray[i].name;
              
              //pics folder found look for JPEG folder inside
              var picFolder = new Folder(folderArray[i]);
              var jpgFolder = picFolder.getFiles ("JPEG");         

              if (count == 1)
              {
                app.open("C:\\A4Template.indd")
              }
              //create random password
              var rand = randomString(3);
              
              loopFiles (jpgFolder, folderName, watermark, output, rand, path);
          }  
      
        win.pnl.progBar.value++;      
     }     
 
    //close progress bar
    win.close();
};
 
 
 function loopFiles(path, folderName, watermark, output, rand, startPath){
      // Create new folder object based on path string  
     var folder = new Folder(path);  
     var filesArray = new Array();
  
     // Get all files in the current folder  
     var files = folder.getFiles();  
     var ref = folderName.replace('%20',' ');
    
     // Loop over the files in the files object  
     for (var i = 0; i < files.length; i++)  
     {  
          // Check if the file is an instance of a file   
          if (files[i] instanceof File)  
          {   
               // Convert the file object to a string for matching purposes (match only works on String objects)  
               var fileString = String(files[i]);  
  
               // Check if the file contains the right extension  
               if ( fileString.match(/.(jpg)$/i) || fileString.match(/.(CR2)$/i) )  
               {  
                    // Place the image in the template  
                    filesArray.push(fileString);
               }  
          }  
     }      
    
    //pass files array and place files
    placeFiles (filesArray, watermark, ref, output, rand, startPath)

    pw = ref + rand;
    writeToFile (ref + ' - ' + pw, startPath);
    //deleteFiles(filesArray); 
    //folder.remove();
};

function writeToFile(text, logPath) {
    path = logPath + "_TIF_PEFN" + "\\" + "paswords.txt" 
	file = new File(path);
	file.encoding = "UTF-8";
	if (file.exists) {
		file.open("e");
		file.seek(0, 2);
	}
	else {
		file.open("w");
	}
	file.write(text + "\r\n"); 
	file.close();
}

//remove all files in the jpg folder
function deleteFiles(filesArray){
     for (j=0; j<6;j+=1)
    {            
        var f = new File(filesArray[j]);    
        f.remove();
    }
}

function placeFiles(filesArray, watermark, ref, output, rand, startPath){

    var doc = app.activeDocument;      

     for (j=0; j<6; j+=1)
     {            
        var f = new File(filesArray[j]);  
        doc.pages[0].rectangles[j].place(f, false);  
        doc.pages[0].rectangles[j].images[0].select();
        sel = doc.selection[0];
        var h = sel.geometricBounds[2] - sel.geometricBounds[0];
        var w = sel.geometricBounds[3] - sel.geometricBounds[1] ;
        if (w > h)
        {
                sel.rotationAngle = 90;
        }
        doc.pages[0].rectangles[j].images[0].fit(FitOptions.CONTENT_TO_FRAME); 

      //move files into a renamed folder in parent _TIF_PEFN folder
      var newFolderName = ref + rand;
      var destinationFolder = new Folder(startPath + "_TIF_PEFN" + "\\" + newFolderName);
      var jpgFolder = new Folder( output + output.substring(output.lastIndexOf('\\') ) +  '_TIF_PEFN'  + "\\" + newFolderName);
      
      if (!destinationFolder.exists) destinationFolder.create();
    
        var vbScript = 'Set fs = CreateObject("Scripting.FileSystemObject")\r';
        vbScript +=  'fs.CopyFile "' + f.fsName.replace("\\", "\\\\") + '", "' + destinationFolder.fsName.replace("\\", "\\\\") + "\\" + f.name + '"';

        app.doScript(vbScript, ScriptLanguage.visualBasic); 
    }

     var allTextFrames = doc.textFrames;
     for (j=0; j<6;j+=1)
    {        
        var tf = allTextFrames[j];          
        tf.contents = watermark
    }
    pw = ref.replace('%20',' ') + rand;
    var tf = allTextFrames[6];
    tf.contents = ref.replace('%20',' ');
    
    var refFrame = allTextFrames[7];
    refFrame.contents = pw;
    
    //store ref for filename
    pdfName = ref;
    
    exportPDF(output);
}

function exportPDF(outputPath){
    //export finished doc
    app.activeDocument.exportFile (ExportFormat.PDF_TYPE, File(outputPath + "\\"+pdfName +".pdf"), false, "Press Quality");
    app.activeDocument.close(SaveOptions.no); 
}