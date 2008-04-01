// Configure HTML Dialog path/filename
var app = Application;
var baseUrl  = getGablibDir(app) + "TOOLS\\Homesite Addons\\smartSense\\"
//Application.AppPath + "UserData\\smartSense\\";
var htmlDialog  = baseUrl + "smartSenseDlg.html";
var objScript = new ActiveXObject("ScriptX.Factory");
var orgCaretPosX, orgCaretPosY

//Define a class to store dialog arguments in
function dlgArg(){}

function Main(){
	
	//"Include" actions from action.js
	eval(include(baseUrl + "actions.js"));
	
	//app.ActiveDocument.BeginUpdate()
	
	//Get orginal position of the cart
	orgCaretPosX=app.ActiveDocument.CaretPosX
	orgCaretPosY=app.ActiveDocument.CaretPosY
	
	//Toggle the word-wrap to get the actuall row numbre

	//with (Application){with (ActiveDocument){ExecCommandByName("cmdEditWordWrap", 0)}}
	//wrapDiff=orgCaretPosY-app.ActiveDocument.CaretPosY
	//with (Application){with (ActiveDocument){ExecCommandByName("cmdEditWordWrap", 0)}}
	
	//orgCaretPosY=orgCaretPosY-wrapDiff
	
	//app.StatusWarning(wrapDiff)
	
	text_s=app.ActiveDocument.Text
	text_a=text_s.split("\n")
	
	//For some reason SmartSense doesn't like when the document is empty.
	//So if thats the case we add an empty space in it before we continue.
	if (text_s=="")
		app.ActiveDocument.Text=" "
		
	
	//If current line contains tabs we have to recalculate orgCaretPosX to get an accurate number
	currentLine_s=text_a[orgCaretPosY-1]
	if (currentLine_s){
		tabs_a=currentLine_s.split("	")
		numberOfTabs=tabs_a.length-1	
	}else{
		numberOfTabs=0
	}
	
	orgCaretPosX=orgCaretPosX-numberOfTabs*3
	
	//app.StatusWarning("x: "+orgCaretPosX+ "  y: "+ orgCaretPosY)
				
	//Select the text to the left
	app.ActiveDocument.CursorWordLeft(1)
	selWord=app.ActiveDocument.SelText
	app.ActiveDocument.SetCaretPos(orgCaretPosX, orgCaretPosY)
	
	//Get all code in front of the cart
	codeBefor=""
	for (i = 0; i <= orgCaretPosY-2;i++)
		codeBefor+=text_a[i]
	
	if(currentLine_s)
		currentLine_s=currentLine_s.substring(0,orgCaretPosX-1)
	codeBefor+=currentLine_s
			
	//Define arguments to send to dialog
	oDA = new dlgArg()
	oDA.selWord=selWord
	oDA.codeBefor=codeBefor
	oDA.app=app
	
	//app.ActiveDocument.EndUpdate()
	app.ActiveDocument.SetCaretPos(orgCaretPosX, orgCaretPosY)
	
	rv=objScript.ShowHtmlDialog(htmlDialog,oDA,"dialogHeight: 270px; dialogWidth: 320px; edge: Raised; center: Yes; help: No; resizable: No; status: No; unadorned: yes;")
	if (rv) {
		if (rv.insert) app.ActiveDocument.InsertText(rv.insert,false);
		if (rv.actions) eval(rv.actions)
	}else
		app.ActiveDocument.SetCaretPos(orgCaretPosX, orgCaretPosY)
			
	objScript=null
}

function include(filnam) {
//Function to "simulate" including of a js-file
  var fso = new ActiveXObject('Scripting.FileSystemObject');
  var fil = fso.OpenTextFile(filnam);
  var s = fil.ReadAll();
  fil.Close();
  return s;
}

function getGablibDir(app) {
	var configFilePath = app.AppPath + "\gabLibrarySettings.txt";
	fso = new ActiveXObject("Scripting.FileSystemObject");
	var configFile = fso.OpenTextFile(configFilePath, 1, true);
	var found = false;
	while (!configFile.AtEndOfStream)  {
		line = configFile.ReadLine();
		parts = line.split("|");
		if (parts[0] == "GablibDir" && parts[1] != "") {
			dir = parts[1];
			found = true;
			break;
		}
	}
	configFile.Close();
	
	if (!found) {
		var configFile = fso.OpenTextFile(configFilePath, 8, true);
		dir = "";
		while (dir.substring(dir.length -1, dir.length) != "\\") {
			dir = app.InputBox("Path", "Path of the GabLibrary on your Development-machine: (U:\\gab_Library\\)", "");
		}
		configFile.WriteLine("GablibDir|" + dir);
		configFile.Close();
	}
	
	configFile = null;
	fso = null;
	
	return dir;
}
