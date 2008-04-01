<SCRIPT LANGUAGE="JavaScript">

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
var updateUserStr = "";
function changeBG(user_id,ui) {
	document.getElementById('tr_'+ui).style.backgroundColor='<%= selectedColor %>'; //change the backgroundcolor of the row so the user know what changed
	
	//-------- now we cut a part of the updateUserStr if it is already in it
	var reg = new RegExp(ui + ":(\\d+),");
	updateUserStr = updateUserStr.replace(reg, "");
	//-------- cut end
	
	//now check all the values of the radiobutton array and add the VALUE
	for(i=0; i < user_id.length; ++i)
		if(user_id[i].checked)
			var myVal = user_id[i].value;
			
	updateUserStr = updateUserStr + ui + ":" + myVal + ","
	document.rbfrm.strMatrix.value = updateUserStr;
}

</script>	