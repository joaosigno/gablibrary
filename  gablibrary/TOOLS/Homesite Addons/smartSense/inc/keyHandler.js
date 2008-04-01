//Capture key strokes
Key = window.event.keyCode
cKey = window.event.ctrlKey
//alert(Key)
switch(Key){ 
	case 27: //Esc
		window.close()
		break
	case 13: //Enter
		send()
		window.close()
		break
	case 38: //Key up
		oneUp()
		break
	case 40: //Key down
		oneDown()
		break
	case 36: //Home
		goHome()
		break
	case 35: //End
		goEnd()
		break
	default :
		chars="0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ"
		searchStr+=chars.charAt(Key-48)
		if (intId) clearInterval(intId)
		intId=self.setInterval('clearSearchStr()', 500)
		doSearch()
		break
}