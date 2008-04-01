/***************************************************************************************************************
* klassen-template
* Michal Gabrukiewicz
***************************************************************************************************************/
function Main() {	
	var app = Application;
	
	//aktuelles defaulttemplate holen
	SET_DEFAULT_TEMPLATE = 5;
	var originalTemplate = app.GetApplicationSetting(SET_DEFAULT_TEMPLATE);
	
	//classtemplate stattdessen
	app.SetApplicationSetting(SET_DEFAULT_TEMPLATE, getGablibDir(app) + "TOOLS\\Homesite Addons\\defaultClassTemplate.asp");
	app.NewDocument(true);
	var active = app.ActiveDocument;
	active.BeginUpdate();
	
	//einiholen
	var txt = active.Text;
	
	txt = txt.replace(/<<< CREATOR >>>/, getCreator(app));
	
	//replace CREATEDON
	var dat = new Date();
	txt = txt.replace(/<<< CREATEDON >>>/, dat.getYear() + "-" + format(dat.getMonth() + 1) + "-" + format(dat.getDate()) + " " + format(dat.getHours()) + ":" + format(dat.getMinutes()));
	
	//check if user wants to add a classtitle
	var classTitle = app.InputBox("Classtitle", "Tell me the name of your new class:", "-");
	//capitalize the first letter
	classTitle = classTitle.substring(0, 1).toUpperCase() + classTitle.substring(1, classTitle.length);
	txt = txt.replace(/<<< CLASSTITLE >>>/g, classTitle);
	
	//aussaschreiben
	active.Text = txt;
	
	//altes template wieder rein
	app.SetApplicationSetting(SET_DEFAULT_TEMPLATE, originalTemplate);
	
	//karee setzen
	active.SetCaretPos(4, 22);
	active.EndUpdate();
}

function format(val) {
	val = "00" + val;
	return val.substring(val.length - 2, val.length);
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

function getCreator(app) {
	var configFilePath = app.AppPath + "\gabLibrarySettings.txt";
	fso = new ActiveXObject("Scripting.FileSystemObject");
	var configFile = fso.OpenTextFile(configFilePath, 1, true);
	var found = false;
	while (!configFile.AtEndOfStream)  {
		line = configFile.ReadLine();
		parts = line.split("|");
		if (parts[0] == "Name" && parts[1] != "") {
			creator = parts[1];
			found = true;
			break;
		}
	}
	configFile.Close();
	
	if (!found) {
		var configFile = fso.OpenTextFile(configFilePath, 8, true);
		creator = app.InputBox("Whats your name?", "Your name:", "Pro Grammer");
		configFile.WriteLine("Name|" + creator);
		configFile.Close();
	}
	
	configFile = null;
	fso = null;
	
	return creator;
}