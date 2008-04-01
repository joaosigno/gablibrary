var showMenuAgain = false;

function menuHoverIn(target, isMonthView) {
	target.style.textDecoration = 'underline';
}

function menuHoverOut(target, isMonthView) {
	target.style.textDecoration = 'none';
}

function goToUrl(url) {
	window.location.href = url;
}

function updateClock() {
	setTimeout("document.getElementById(\"todayClock\").innerText = now();updateClock()", 1000);
}

function now() {
	nw = new Date();
	hr = nw.getHours();
	hr = (hr < 10) ? "0" + hr : hr;
	mn = nw.getMinutes();
	mn = (mn < 10) ? "0" + mn : mn;
	sc = nw.getSeconds();
	sc = (sc < 10) ? "0" + sc : sc;
	return hr + ":" + mn + ":" + sc;
}

function goToDate(urlToLoad, selectedDate, calendarUrl) {
	//first set the hiddenvalue to zero
	calendarForm.dummyGoToDateField.value = '0';
	
	//setInactive();
	
	//now show the calendar
	openCenteredModal(calendarUrl + "&selectedDate=" + selectedDate, 300, 300, false);
	
	//setActive();
	
	//get the value of the hidden-field
	dateVal = calendarForm.dummyGoToDateField.value;
	
	//if the value changed we reload
	if (dateVal != "0") {
		newUrl = urlToLoad.replace("dateToChange", dateVal);
		goToUrl(newUrl);
	}
}

function setInactive() {
	ht = document.getElementsByTagName("body");
	ht[0].style.filter = "Gray()";
}

function setActive() {
	ht = document.getElementsByTagName("body");
	ht[0].style.filter = "";
}

function dayMenuGoToUrl(url) {
	//we have to catch the clicked date.
	clickedDate = document.getElementById("dayMenuClickedDate");
	goToUrl(url.replace("dateToChange", clickedDate.value));
}

function mouse(x, y) {
	this.X = x;
	this.Y = y;
}

var aMouse = new mouse(0, 0);

document.onmousemove = getMousePosition;

function getMousePosition(e) {
	var evt = (e) ? e : window.event;
	if (evt.clientX) {
		aMouse.X = evt.clientX + (document.documentElement.scrollLeft ? document.documentElement.scrollLeft : document.body.scrollLeft);
	    aMouse.Y = evt.clientY + (document.documentElement.scrollTop ?  document.documentElement.scrollTop :  document.body.scrollTop);
	}
}

function showDayMenu(dayviewUrl, currentDate, weekDayName) {
	var menu = document.getElementById("dayMenu");
	var menuHeader = document.getElementById("dayMenuItemHeadline");
	
	if (arguments.length > 1) {
		document.getElementById("dayMenuClickedDate").value = currentDate;
		menuHeader.innerHTML = weekDayName + ", " + currentDate;
	}
	//if the menu was already here then the user maybe clicked on another
	//day to view the menu. so we need to use this trick.
	if (menu.style.visibility == "visible") showMenuAgain = true;

	var mouseX = aMouse.X;
	var mouseY = aMouse.Y;
	var bottomSpace = document.body.clientHeight - mouseY - 10;
	var rightSpace = document.body.clientWidth - mouseX - 10;
	menu.style.left = mouseX + "px";
	menu.style.top = mouseY + "px";
	menu.style.height = "70px"; // for firefox
	menu.style.visibility = "visible";
}


function hideDayMenu() {
	var menu = document.getElementById("dayMenu");
	menu.style.visibility = "hidden";
	document.onclick = "";
	if (showMenuAgain) {
		showDayMenu();
		showMenuAgain = false;
	}
}

function changeCssClass(cssClass, target) {
	target.className = cssClass;
}
