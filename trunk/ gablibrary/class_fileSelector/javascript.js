FSSelectedFolder = null;
FSSelectedFileID = null;

/******************************************************************************************
* click on a folder
******************************************************************************************/
function FSToggleFolder(folderID, sourceObj) {
	folder = document.getElementById(folderID);
	folder.style.display = (folder.style.display == 'none') ? 'inline' : 'none';
	
	if (FSSelectedFolder) FSSelectedFolder.className = sourceObj.className;
	sourceObj.className += " selectedFolder";
	
	//change the folder-image
	folderIMG = document.getElementById("img" + folderID);
	if (FSEndsWith(folderIMG.src, "folder.gif")) {
		folderIMG.src = folderIMG.src.replace(/folder\.gif/, "folder_open.gif");
	} else {
		folderIMG.src = folderIMG.src.replace(/folder_open\.gif/, "folder.gif");
	}
	
	FSSelectedFolder = sourceObj;
}

/******************************************************************************************
* endswith
******************************************************************************************/
function FSEndsWith(str, compare) {
	return (str.substr(str.length - compare.length, str.length) == compare);
}

/******************************************************************************************
* select a file
******************************************************************************************/
function FSselect(labelID, sender) {
	lbl = document.getElementById(labelID);
	lbl.className = (sender.checked) ? 's' : '';
	if (sender.type == "radio") {
		if (FSSelectedFileID != null && labelID != FSSelectedFileID) document.getElementById(FSSelectedFileID).className = '';
		FSSelectedFileID = labelID;
	}
}

/******************************************************************************************
* uncheck radiobutton
******************************************************************************************/
function FSuc(sender) {
	sender.checked = false;
	document.getElementById(FSSelectedFileID).className = '';
}