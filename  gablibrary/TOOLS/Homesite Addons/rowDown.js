/***************************************************************************************************************
* moving the selected line one line down
* written by Michal Gabrukiewicz
* recommended shortcut ALT+arrow down
***************************************************************************************************************/
function Main() {	
	var app = Application;
	var a = app.ActiveDocument;
	a.BeginUpdate();
	var l = a.CaretPosY;
	var currentLine = a.Lines(l - 1);
	var nextLine = a.Lines(l);
	a.Lines(l - 1) = nextLine;
	a.Lines(l) = currentLine;
	a.EndUpdate();
}