/***************************************************************************************************************
* moving the selected line one line up
* written by Michal Gabrukiewicz
* recommended shortcut ALT+arrow up
***************************************************************************************************************/
function Main() {	
	var app = Application;
	var a = app.ActiveDocument;
	//var strText = new String();
	
	a.BeginUpdate();
	var l = a.CaretPosY;
	var currentLine = a.Lines(l - 1);
	var beforeLine = a.Lines(l - 2);
	a.Lines(l - 1) = beforeLine;
	a.Lines(l - 2) = currentLine;
	a.EndUpdate();
}