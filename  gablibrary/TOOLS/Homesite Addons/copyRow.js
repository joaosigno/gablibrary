/***************************************************************************************************************
* copies the current row one row down.
* written by Michal Gabrukiewicz
* recommended shortcut ALT + page down
***************************************************************************************************************/
function Main() {	
	var app = Application;
	var a = app.ActiveDocument;
	a.BeginUpdate();
	var toCopy = a.Lines(a.CaretPosY - 1);
	a.CursorLineEnd(false);
	a.InsertText("\n", false);
	a.Lines(a.CaretPosY - 1) = toCopy;
	a.EndUpdate();
}