var searchStr=""
var intId=null
var currListObj=null
var currNode=null
var currClassName=""
var oList=null, oListNodes=null

function getListObjects(){
	oList=document.getElementById("list")
	oListNodes=oList.document.getElementsByTagName("DIV")
}

function clearSearchStr(){
	searchStr=""
	if (intId) clearInterval(intId)
}

function doUnSelecte(o){
	if (o) o.className=currClassName
}

function oneUp(mode){
	if (!currNode.previousSibling) return false
	if(currNode){
		doUnSelecte(currNode)
		doSelect(currNode.previousSibling)
	}
	return true
}

function oneDown(mode){
	if (!currNode.nextSibling) return false
	if(currNode){
		doUnSelecte(currNode)
		doSelect(currNode.nextSibling)
	}
	return true
}

function doSelect(o){
	if (!o) return false
	currNode=o
	currNode.scrollIntoView()
	currClassName=currNode.className
	currNode.className=currClassName + " selected"
	return true
}

function check(){
	alert(currNode.innerText)
	alert(currNode.previousSibling.innerText)
}

function clickSelect(o){
	doUnSelecte(currNode)
	doSelect(o)
}

function doSearch(){
	currSS=searchStr
	for (i = 0; i <= oListNodes.length-1; i++){
		if(oListNodes[i].innerText.substring(0,currSS.length).toUpperCase()==currSS.toUpperCase()){
			doUnSelecte(currNode)
			doSelect(oListNodes[i])
			break
		}
	}i
}

function updateDelta(){
	if (event.wheelDelta >= 120) oneUp(1)
	else if (event.wheelDelta <= -120) oneDown(1)
}

function goHome(){
	doUnSelecte(currNode)
	if (oListNodes[0]) doSelect(oListNodes[0])
}

function goEnd(){
	doUnSelecte(currNode)
	doSelect(oListNodes[oListNodes.length-1])
}

