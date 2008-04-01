/***************************************************************************************************************
* neues dokument wird erstellt und mit datum und creator bef&uuml;llt
* Michal Gabrukiewicz
***************************************************************************************************************/
function Main() {	
	var app = Application;	
	app.NewDocument(true);
	var active = app.ActiveDocument;
	active.BeginUpdate();
	
	//einiholen
	var txt = active.Text;
	
	//replace CREATOR
	txt = txt.replace(/<<< CREATOR >>>/, getCreator(app));
	
	//replace CREATEDON
	var dat = new Date();
	txt = txt.replace(/<<< CREATEDON >>>/, dat.getYear() + "-" + format(dat.getMonth() + 1) + "-" + format(dat.getDate()) + " " + format(dat.getHours()) + ":" + format(dat.getMinutes()));
	
	//check if user wants to add an description
	txt = txt.replace(/<<< DESCRIPTION >>>/, app.InputBox("File Description", "Describe the content of your file now:", "-"));
	
	//aussaschreiben
	active.Text = txt;
	
	//karee setzen
	active.SetCaretPos(4, 28);
	active.EndUpdate();
}

function format(val) {
	val = "00" + val;
	return val.substring(val.length - 2, val.length);
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