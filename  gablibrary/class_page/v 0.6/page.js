/*****************************************************************************************************************
'' @SDESCRIPTION:	opens a centered modaldialog with the wanted parameters. 
'' @PARAM:			- url [string]: url of the page you want to display in the dialog
'' @PARAM:			- width [int]: height of the window
'' @PARAM:			- height [int]: width of the window
'' @PARAM:			- scrolling [bool]: should scrollbars be enabled or not
******************************************************************************************************************/
function openCenteredModal(url, width, height, scrolling) {
	strScroll = (scrolling) ? "Yes" : "No";
	return window.showModalDialog(url, window, 'dialogHeight: ' + height + 'px; dialogWidth: ' + width + 'px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No; scroll: ' + strScroll);
}

/*****************************************************************************************************************
'' @SDESCRIPTION:	handles the escape key in modal dialog's
******************************************************************************************************************/
function handleModalEscape() {
	currentKeyCode = event.keyCode;
	
	if (typeof(window) != "undefined") {
		switch(currentKeyCode) {
			//escape key
			case 27 :
				window.close();
				break;
		}
	}
}

/*****************************************************************************************************************
'' @SDESCRIPTION:	use this to view the source code of a page, e.g. onclick="viewSource()"
******************************************************************************************************************/
function viewSource() {
	window.location='view-source:' + window.location.href;
}