<SCRIPT LANGUAGE="JavaScript">

var lastRowColor;

function rowHoverIn(row) {
	lastRowColor = row.bgColor;
	row.bgColor = '<%= rowColorHover %>';
}

function rowHoverOut(row) {
	row.bgColor = lastRowColor;
}

function dtToggleVisibility(elemID) {
	var elem = document.getElementById(elemID);
	if (elem.style.display == "none") {
		elem.style.display = "block";
	} else {
		elem.style.display = "none";
	}
}

//*************************************************************************************
// selects all radiobuttons which has the wanted value 
//*************************************************************************************
function selectAllRadioButtons(neededValue) {
	if (confirm("<%= TXTSELECTALLLASTSLONG %>")) {
		for (i=0; i < document.rbfrm.length; i++) {
			curObj = rbfrm[i];
			
			if (curObj.name.indexOf("rb_") != -1) {
				if (curObj.checked == false) {
					if (curObj.value == neededValue) {
						curObj.checked = true;
						nameParts = curObj.name.split("_");
						changeBG(curObj.value, nameParts[1]);
					}
				}
			}
		}
	}
}

//*************************************************************************************
// THIS FUNCTION IS NEEDED TO CHECK IF THE USER PRESSED RETURN IN THE FULLSEARCH
//*************************************************************************************
function checkChar(e) {
	var characterCode;
	if(e && e.which){ //if which property of event object is supported (NN4)
		e = e
		characterCode = e.which //character code is contained in NN4's which property
	} else {							
		e = event						
		characterCode = e.keyCode //character code is contained in IE's keyCode property
	}
	if(characterCode == 13) {
		restoreMyFilter();
	}
}

//*************************************************************************************
// FAST DELETE CONFIRMATION BOX
//*************************************************************************************
function yesNoFastDelete(myID,msg) {
	Check = confirm(msg);
	if(Check == true) {
		document.myVeryHiddenForm.fastDeleteID.value = myID;
		document.myVeryHiddenForm.submit();
	}
}

//*************************************************************************************
// WE SET ALL FILTERFIELDS TO THE DEFAULT POSITION
//*************************************************************************************
function restoreMyFilter() {
	for (i=0; i < document.myVeryHiddenForm.length; i++)
		if (myVeryHiddenForm[i].name.indexOf("fltrField_") != -1)
			if (myVeryHiddenForm[i].type == "text")
				myVeryHiddenForm[i].value = "";
			else
				myVeryHiddenForm[i].value = "XXXYYYZZZXXX";
}

//*************************************************************************************
// THIS FUNCTION WRITES VALUES IN A HIDDEN FIELD SO WE DONT NEED TO UPDATE ALL RECORDS
// WE ONLY UPDATE THE RECORDS WHICH ARE CHANGED
//*************************************************************************************
var updateUserStr = ",";
function changeBG(rbValue, userID) {
	//change the backgroundcolor of the row so the user know what changed
	tr = document.getElementById("tr_" + userID);
	tr.style.backgroundColor = "<%= SELECTEDBGCOLOR %>";
	for (i = 0; i < tr.childNodes.length; i++) {
		tr.childNodes[i].style.color = "<%= SELECTEDCOLOR %>";
	}
	
	//now we cut a part of the updateUserStr if it is already in it
	//its important to find a string with a comma in front so we prevent the following situation:
	//we want find 2:201 and find 1992:201 tooo .. that would be bad. so the solution is to
	//include a comma at the beginning. think about it!
	
	var reg = new RegExp("," + userID + ":(\\d+),");
	updateUserStr = updateUserStr.replace(reg, ",");			
	updateUserStr = updateUserStr + userID + ":" + rbValue + ",";
	
	//put the new string into the strMatrix field.
	document.rbfrm.strMatrix.value = updateUserStr;
}

var tableSubmitionFlag = true;
//*************************************************************************************
// This function submits the table if a filter has been activated and makes sure that
// only one submit will be executed
//*************************************************************************************
function submitTableData() {
	if(tableSubmitionFlag) {
		tableSubmitionFlag = false;
		document.myVeryHiddenForm.submit();
	}
}

</script>
