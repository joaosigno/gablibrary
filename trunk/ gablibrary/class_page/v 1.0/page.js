function gablib() {}

//toggles the visibility of an element (eID)
//displayType is optional and by default block (inline possible)
//returns true if it is visible
gablib.toggle = function(eID, displayType) {
	el = byID(eID);
	if (!el) return false;	
	if (el.style.display.toLowerCase() == 'none') {
		el.style.display = (arguments.length > 1) ? displayType : 'block';
		return true;
	} else {
		el.style.display = 'none';
	}
}

//opens a centered modaldialog with the wanted parameters. 
function openCenteredModal(url, width, height, scrolling) {
	strScroll = (scrolling) ? "Yes" : "No";
	return window.showModalDialog(url, window, 'dialogHeight: ' + height + 'px; dialogWidth: ' + width + 'px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No; scroll: ' + strScroll);
}

//opens a modaldialog right-bottom to the mouse. if no place on the bottom then to the top.
function openModal(url, width, height, scrolling) {
	strScroll = (scrolling) ? "Yes" : "No";
	var mouseX 	= window.event.screenX;
	var mouseY = window.event.screenY;
	return window.showModalDialog(url, window, 'center:no; dialogHeight: ' + height + 'px; dialogTop:' + mouseY + '; dialogLeft:' + mouseX + 'px; dialogWidth: ' + width + 'px; edge: Raised; help: No; resizable: No; status: No; scroll: ' + strScroll);
}

function handleModalEscape() {
	currentKeyCode = event.keyCode;
	if (typeof(window) != undefined) {
	
		if (currentKeyCode == 27) window.close();
	}
}

function viewSource() {
	window.location='view-source:' + window.location.href;
}

function byID(ID) {
	return document.getElementById(ID);
}