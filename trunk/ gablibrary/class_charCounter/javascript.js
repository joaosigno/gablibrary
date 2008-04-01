<SCRIPT language="JavaScript">
<!--
//*****************************************************************
//* FIRST WE SET THE TEXTS FOR UNUSED AND USED BAR
//*****************************************************************
var charsUsed = document.charCounter_used_<%= controlName %>.alt;
var charsUnused = document.charCounter_unused_<%= controlName %>.alt;

//*****************************************************************
//* DYNAMICCHARCOUNTERJSFUNCTION
//*****************************************************************
function dynamicCharCounterJSFunction(targetObj,formName) {

	//get the max-allowed value
	eval("var maximum = " + formName + ".dynamicCharCounterMaxValue_" + targetObj.name + ".value;");
	
	//get the Barlength
	eval("var barLength = document.getElementById('charCounterTable_" + targetObj.name + "').width;");
	
	availableChars = maximum - targetObj.value.length;
	
	//cut the string if we already have too much chars.
	if (availableChars < 0) { targetObj.value = targetObj.value.substring(0,maximum); availableChars = 0; }

	//set the width of the "used"-bar
	eval("document.charCounter_used_" + targetObj.name + ".width = barLength * (maximum - availableChars) / maximum;");
	
	//set the "alternative" text to the images (bars)
	eval("document.charCounter_used_" + targetObj.name + ".alt = targetObj.value.length + ' ' + charsUsed;");
	eval("document.charCounter_unused_" + targetObj.name + ".alt = availableChars + ' ' + charsUnused;");
}
//-->
</SCRIPT>
