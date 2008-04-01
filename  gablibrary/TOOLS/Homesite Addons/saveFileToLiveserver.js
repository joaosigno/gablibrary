/***************************************************************************************************************
* speichert die Datei auf dem Live-server
***************************************************************************************************************/
function Main() {
	var app = Application;
	var active = app.ActiveDocument;
	
	var configFilePath = app.AppPath + "\gabLibrarySettings.txt";
	fso = new ActiveXObject("Scripting.FileSystemObject");
	var configFile = fso.OpenTextFile(configFilePath, 1, true);
	var found = false;
	while (!configFile.AtEndOfStream)  {
		line = configFile.ReadLine();
		parts = line.split("|");
		if (parts[0] == "LiveServer" && parts[1] != "") {
			liveServerDir = parts[1];
			found = true;
			break;
		}
	}
	configFile.Close();
	
	if (!found) {
		var configFile = fso.OpenTextFile(configFilePath, 8, true);
		liveServerDir = "";
		while (liveServerDir.substring(liveServerDir.length -1, liveServerDir.length) != "\\") {
			liveServerDir = app.InputBox("Liveserver Path?", "Path of your liveserver? (e.g. L:\\intranet\\):", "");
		}
		configFile.WriteLine("LiveServer|" + liveServerDir);
		configFile.Close();
	}
	
	configFile = null;
	fso = null;
	
	var activeFile = active.Filename;
	var kurz = activeFile.substring(3, activeFile.length);
	var liveServerDir = liveServerDir + kurz;
	
	var meldung = "Are you sure you want to save the file " + activeFile + "\non the Live-Server at " + liveServerDir + "?";
	
	//wenn der user auf yes klickt, wird datei gespeichert
	if (app.MessageBox (meldung, "Save file to LIVE-Server?", 4) == 6) {
		if (!active.SaveAs(liveServerDir)) {
			app.MessageBox("Saving failed.", "Error while saving!", 48)
		} else {
			// wir m&uuml;ssen die datei wieder auf dem urspr&uuml;nglichen platz speichern,
			// damit die datei auch wieder am urspr&uuml;nglichen ort gespeichert wird,
			// wenn der user auf save klick ohne das skript zu benutzen.
			active.SaveAs(activeFile);
		}
	}
}