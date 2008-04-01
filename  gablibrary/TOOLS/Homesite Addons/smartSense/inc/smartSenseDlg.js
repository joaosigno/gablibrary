selWord_s=dialogArguments.selWord //The code to the left of the cursor
codeBefor_s=dialogArguments.codeBefor.toLowerCase() //All the code in the document in front of the cursor.
app=dialogArguments.app //All the code in the document in front of the cursor.

var items=new Array();


function Item(display,insert,help,actions,id,fontWeight,cName) {
	this.display=display;
	this.insert=insert;
	this.help=help;
	this.actions=actions;
	this.id=id;
	this.fontWeight=fontWeight;
	this.cName=cName;
}

var xml
var filepath_s=app.ActiveDocument.Filename //The path to current file		

//Define a class to store dialog arguments in
//This object is returned to parent on send.
function dlgArg(){}

//Look for patterns in the code to decide what language definition file to use
var highestIndex=-1
var langDefFile="javascript.xml" //Default file to use
var actions=""

var xml = new ActiveXObject("Msxml2.DOMDocument");

if (xml.load("patterndef.xml")){  //If loading xml was successful go one
	var root = xml.documentElement;
			
	for(var i =0;i<root.childNodes.length;i++) {
		currNode=root.childNodes[i]
		currPattern=currNode.selectSingleNode("pattern").text
		
		if(currNode.selectSingleNode("langDefFile")){
			currlangDefFile=currNode.selectSingleNode("langDefFile").text
			currActions=""
		}
			
		if(currNode.selectSingleNode("actions")){
			currActions=currNode.selectSingleNode("actions").text
			currlangDefFile=""
		}
		
		//If this the pattern is closest to the cursor, set it to current
		if(codeBefor_s.lastIndexOf(currPattern)>highestIndex) {
			highestIndex=codeBefor_s.lastIndexOf(currPattern)
			langDefFile=currlangDefFile
			actions=currActions
		}
	}
}else{ 	
	app.MessageBox("Couldn't load patternDef.xml!","Error!",0)
}

//If we still haven't found any language definition file to use we take a look at the extensions of the file.
//Else us default.
if(highestIndex<0){
	extension=filepath_s.substring(filepath_s.lastIndexOf(".")+1,filepath_s.length)
	
	if (xml.load("extensionDef.xml")){  //If loading xml was successfull go one
		var root = xml.documentElement;
				
		for(var i =0;i<root.childNodes.length;i++) {
			currNode=root.childNodes[i]
			currExtension=currNode.selectSingleNode("extension").text
			currlangDefFile=currNode.selectSingleNode("langDefFile").text
			
			if(currExtension==extension) {
				langDefFile=currlangDefFile
			}
		}
	}else{ 	
		app.MessageBox("Couldn't load extensionDef.xml!","Error!",0)
	}
}		

if(langDefFile!=""){
	//Now load the language def file
	if (xml.load("langDef//"+langDefFile)){  //If loading xml was successfull go one
		var list_s=""
		var root = xml.documentElement;
		var oN=null;
		
		//Type handler
		//If selected word contains an underscore we should look for types
		if (selWord_s.indexOf("_")>-1){
			type=getType(selWord_s)
			types=root.selectSingleNode("types")
			for(var i =0;i<types.childNodes.length;i++) {
				currTypeNode=types.childNodes[i]
							
				name=currTypeNode.attributes.getNamedItem("name").text
				if (name+"."==type+"."){
					typeString=currTypeNode.selectSingleNode("type")
					selWord_s=typeString.text
				}
			}
		}
	
		//Remove .
		if (endC(selWord_s)==".") selWord_s=selWord_s.substring(0,selWord_s.length-1)
				
		var elements=getNode(selWord_s) //Return the node which match this selected word
		if(!elements) elements=root.selectSingleNode("elements") //If false use the root elements

		//"Precompile" all regular expresion elements, RexEl
		for(var y =0;y<elements.childNodes.length;y++) {
			currNode=elements.childNodes[y]
			if(currNode.nodeName=="RexEl"){
				var pattern_s=getNodeText("pattern",currNode)
				if(pattern_s){	//Only do this operation if the pattern is defined
							
					var trimStart=getNodeText("trimStart",currNode)
					if(!trimStart) trimStart=0
					
					var trimEnd=getNodeText("trimEnd",currNode)
					if(!trimEnd) trimEnd=0
					
					var type=getNodeText("type",currNode)
					if(!type) type=""
					
					var scanAllDocs=getNodeText("scanAllDocs",currNode)
					if(!scanAllDocs) scanAllDocs=0
					
					var extensions=getNodeText("extensions",currNode)
					if(!extensions) extensions=""
					else  scanAllDocs=1
					
					var displayCode=getNodeText("displayCode",currNode)
					if(!displayCode) displayCode="%match%"
					
					var insertCode=getNodeText("insertCode",currNode)
					if(!insertCode) insertCode="%match%"
					
					arr_a=getRegExpArray(pattern_s, scanAllDocs, extensions)
					
					oRegExpXML=getRegExpXMLObject(arr_a, trimStart, trimEnd, type, displayCode, insertCode)
					
					currNode.parentNode.replaceChild(oRegExpXML, currNode);
				}
			}
		}
		
		list_s=childLoop(elements) 
	
	}else{ 	
		alert("Couldn't load xml file!")
		window.close()
	}
} else if(actions!=""){
	eval(actions)
}else{
	alert("Error")
}

function childLoop(oNodes){
//Return html code for a list of all child nodes
	var s=""
	for(var i =0;i<oNodes.childNodes.length;i++) {
		currNode=oNodes.childNodes[i]
		var currDiplay=""
		var currInsert=""
		var currHelp=""
		var currActions=""
		
		//Collect info about current node
		currDiplay=getNodeText("display",currNode)
		if(!currDiplay){
			alert("Err: Languages file "+langDefFile + " are not properly declared.\nDisplay node is required.")			
			window.close()
		}
		
		currInsert=currDiplay
		//If the inset node is present over ride the display node info
		if (getNodeText("insert",currNode)) currInsert=getNodeText("insert",currNode)
		
		currHelp=getNodeText("help",currNode)
		if(!currHelp) help_s=""
		else help_s=currHelp
		
		currActions=getNodeText("actions",currNode)		
		
		var endChar=""
		var cName="default"	//Class name
		var fontWeight="normal"
		
		//See if there should be a dot in the end of the string
		currChildElementsNode=currNode.selectSingleNode("childElements")
		if (currChildElementsNode){
			if (currChildElementsNode.childNodes.length>0){
				endChar="."
				cName="objectNode"
				fontWeight="bold" //All nodes with cild nodes are in bold.
			}
		}
		
		//Don't set the end dot in this cases
		if (endC(currInsert)=="=") endChar=""
		if (endC(currInsert)==":") endChar=""
		//MG: when its an include file then we dont need a dot too...
		if (endC(currInsert)==">") endChar=""
			
		//If current string ends with a dot it should look lika an object with child nodes.		
		if((endC(currDiplay)==".")||(endC(currInsert)==".")){
			cName="objectNode"
			fontWeight="bold"
		}
		
		//Check if there's any type attribute defined
		if(currNode.attributes.getNamedItem("type")){
			switch(currNode.attributes.getNamedItem("type").text){ 
				case "p": 
					cName="propertie"
					endChar=""
					break
				case "m": 
					cName="method"
					break
				case "style": 
					cName="style"
					break
				case "resWord": 
					cName="reservedWord"
					break
				case "o": 
					cName="objectNode"
					break
			}
		}
		
		//MG: check if obsolete
		if (currNode.attributes.getNamedItem("obsolete")) {
			if (currNode.attributes.getNamedItem("obsolete").text == "1") cName += " obsolete";
		}
		
		//Add the end char at the end of the insert string
		currInsert+=endChar
		currentId=items.length
		
		//Add current node to the array			
		items[currentId]=new Item(currDiplay,currInsert,currHelp,currActions,currentId,fontWeight,cName);
	}
	
	//items=sortBy('display', items)
	
	for (j=0;j<items.length;j++){
			//Put together the html code for current node			
			
			//MG: replace the quotes fixed. no "false" displayed as help anymore if no help available.
			help = items[j].help;
			if (help != "") help = help.replace(/\"/g, "&quot;");
			s+="<div id=\"node_"+ items[j].id +"\" title=\""+ ((help != "") ? help : "") + "\" style=\"font-weight:"+items[j].fontWeight+";\""
			s+=" onDblClick=\"send()\" onClick=\"clickSelect(this)\" class=\""+items[j].cName+"\">"
			s+=items[j].display + "</div>"
	}
	
	return s
}		

//Here's the object sort functions
function sortBy(prop, oArray, mode) {
	sortProp=prop;
	return oArray.sort(sortInc);
}

function sortInc(prop1,prop2){
	if (prop1[sortProp]<prop2[sortProp]) retVal=-1;
	else if (prop1[sortProp]>prop2[sortProp]) retVal=1;
	else retVal=0;
	return retVal;
}
	
function getNode(name){
	findParent(root.selectSingleNode("elements"), name)
	
	//If no parent node found in elements look among to common nodes
	if(!oN) findParent(root.selectSingleNode("common"), name)
	return oN
}	

function getNodeText(nodeName, oNode){
	if (!oNode.selectSingleNode(nodeName))
		return false
	else
		return oNode.selectSingleNode(nodeName).text
}	

function findParent(oNodes, name){
	if(!oNodes) return false
	
	var end = false;
	var currentNode=null;
	for(var i =0;i<oNodes.childNodes.length;i++) {
		currentNode=oNodes.childNodes[i]
		
		current_s=getNodeText("display",currentNode)
		//If the inset node is present over ride the display node info
		if (getNodeText("insert",currentNode)) current_s=getNodeText("insert",currentNode)
		
		currChildElementsNode=currentNode.selectSingleNode("childElements")
		
		if ((current_s==name)&&((currentNode.childNodes.length>0))){
			oN = currChildElementsNode;
			break;
		}else{
			findParent(currChildElementsNode, name)
		}
	}
}

function regExpList(pattern_s, trimStart, trimEnd, type, scanAllDocs, thisExtension, displayCode, insertCode){
//This function create lists of matches of patterns from the documents open in the editor.
	
	arr_a=getRegExpArray(pattern_s, scanAllDocs, thisExtension)
	
	var xml_s="<childElements>\n"
	
	xml_s+=getRegExpXML(arr_a, trimStart, trimEnd, type, displayCode, insertCode)
	
	xml_s+="</childElements>"
	xml.loadXML(xml_s);
	var root = xml.documentElement;
	list_s=childLoop(root) 
}

function getRegExpXML(arr_a, trimStart, trimEnd, type, displayCode, insertCode){
//This function render xml code based on the content of an array
//	arr_a= the array
//	trimStart = chars to trim in the beginning of the strings in the array 
//	trimEnd = chars to trim in the end of the strings in the array
//	Type what type to display the list as. EG m=method, style=CSS
//	displayCode the code to put in the display node, where %match% are replaced with current trimmed string
//	Same as displayCode but for the insert node
	var typeCode_s=""
	var xml_s=""
	if(type) typeCode_s=" type=\""+type+"\""
	
	for (i = 0; i <= arr_a.length-1; i++){ 
		current_s=arr_a[i]
		current_s=current_s.substring(trimStart,current_s.length-1*trimEnd)
		
		xml_s+="	<el"+typeCode_s+">\n"
		
		if(displayCode) xml_s+="		<display><![CDATA["+replace(displayCode,"%match%",current_s)+"]]></display>\n"
		else xml_s+="		<display>"+current_s+"</display>\n"
		
		if(insertCode) xml_s+="		<insert><![CDATA["+replace(insertCode,"%match%",current_s)+"]]></insert>\n"
		else xml_s+="		<insert>"+current_s+"</insert>\n"
		
		xml_s+="	</el>\n"
	}
	
	return xml_s;
}

function getRegExpXMLObject(arr_a, trimStart, trimEnd, type, displayCode, insertCode){
//	This function render an xml document fragment object based on the content of an array
//	arr_a= the array
//	trimStart = chars to trim in the beginning of the strings in the array 
//	trimEnd = chars to trim in the end of the strings in the array
//	Type what type to display the list as. EG m=method, style=CSS
//	displayCode the code to put in the display node, where %match% are replaced with current trimmed string
//	Same as displayCode but for the insert node

	var docFragment = xml.createDocumentFragment();
	for (i = 0; i <= arr_a.length-1; i++){ 
		current_s=arr_a[i]
		current_s=current_s.substring(trimStart,current_s.length-1*trimEnd)
		
		var elem = xml.createElement("el");
		if(type) elem.setAttribute("type", type)
		
		var displayEl = xml.createElement("display");		
		if(displayCode) displayEl.text=replace(displayCode,"%match%",current_s)
		else  displayEl.text=current_s
		elem.appendChild(displayEl);
		
		var insertEl = xml.createElement("insert");		
		if(insertCode) insertEl.text=replace(insertCode,"%match%",current_s)
		else  insertEl.text=current_s
		elem.appendChild(insertEl);

		docFragment.appendChild(elem);
	}
	return docFragment
}

function getRegExpArray(pattern_s, scanAllDocs, thisExtension){
	//Create the regExp object
	var re = new RegExp(pattern_s,"g")
	
	//Add the code for current page
	var scanCode=app.ActiveDocument.Text
	
	//If scanAllDocs is thru look in the other documents as well.
	if(scanAllDocs){
		
		var currentIndex=app.DocumentIndex	//Store the index of current page
		for (idx = 0; idx <= app.DocumentCount-1; idx++){ 
			if(thisExtension){ //Scan only docs of this type
				filepath_s=app.DocumentCache(idx).Filename
				extension=filepath_s.substring(filepath_s.lastIndexOf(".")+1,filepath_s.length)
				if(extension==thisExtension){
					app.DocumentIndex=idx
					scanCode+=app.ActiveDocument.Text
				}
			}else{
				app.DocumentIndex=idx
				scanCode+=app.ActiveDocument.Text
			}
		}
		app.DocumentIndex=currentIndex	//Activate the current page again
	}
	
	arr_a=scanCode.match(re)

	if(arr_a){
		arr_a.sort()
		arr_a=removeDuplicates(arr_a)
		return arr_a
	}else
		return false;
}

function removeDuplicates(oArr){
	if (oArr.length==1) return oArr
	
	for (i = 1; i <= oArr.length-1; i++){ 
		if(oArr[i]==oArr[i-1]){
		    oArr.splice(i,1)  
			oArr=removeDuplicates(oArr)
		}
	}
	return oArr
}

function replace(inStr, repStr, newStr){
	var re = new RegExp(repStr,"g")
	return inStr.replace(re, newStr)
}

function send(){
	arrIndex=currNode.id.substring(5,currNode.id.length)
	arrIndex=parseInt(arrIndex)
	
	//Define argumente to send to dialog
	oDA = new dlgArg()
	oDA.insert=items[arrIndex].insert
	oDA.actions=items[arrIndex].actions
	
	window.returnValue=oDA
	window.close()
}

function getType(s){
	if (s.indexOf("_")==-1)	return false;
	
	type=s.substring(s.lastIndexOf("_")+1,s.length)
	type=type.substring(0,type.length-1)			
	return type
}

function endC(in_s){
	return in_s.charAt(in_s.length-1)
}