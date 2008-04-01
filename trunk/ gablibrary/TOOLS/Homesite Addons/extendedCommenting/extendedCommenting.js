/***************************************************************************************************************
* commenting & uncommenting ASP in Homesite like in Visual Studio (in use with gabLibrary)
* - it comments lines which are not commented and ucomments commented lines.
* - single & multiline support
* written by Michal Gabrukiewicz (gabru@grafix.at) & Michael Rebec
* recommended shortcut CTRL+SHIFT+#
***************************************************************************************************************/
function Main() {	
	var COMMENT = "'";
	var app = Application;
	var active = app.ActiveDocument;
	var strText = new String();
	
	active.BeginUpdate();
	strText = active.SelText;

	var lines = strText.split("\r");
	strText = "";
	for (var i = 0; i < lines.length; i++) {
		if (i > 0) {
			strText += "\r";
		}
		
		var found = false;
		for (var j = 0; j < lines[i].length; j++) {
			if (lines[i].charCodeAt(j) == 39) {
				strText += removeAt(lines[i], j);
				found = true;
				break;
			} else if ((lines[i].charCodeAt(j) != 9) && (lines[i].charCodeAt(j) != 32) && (lines[i].charCodeAt(j) != 10) && (lines[i].charCodeAt(j) != 13)) {
				strText += insertAt(lines[i], j, COMMENT);
				found = true;
				break;
			}
		}
		
		if (!found) {
			strText += lines[i];
		}
	}
		
	active.EndUpdate();
	active.SelText = strText;
}

function insertAt(str, pos, newStr) {
	return str.substring(0, pos) + newStr + str.substring(pos);
}

function removeAt(str, pos) {
	return str.substring(0, pos) + str.substring(pos + 1);
}