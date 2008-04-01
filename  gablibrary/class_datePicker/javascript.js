var lastCssClass;

function submitCalendarForm() {
	datePickerFrm.submit();
}

function changeChosenDate(value) {
	datePickerFrm.chosenDate.value = value;
	submitCalendarForm();
}

function hoverIn(obj, cssClass) {
	lastCssClass = obj.className
	obj.className = lastCssClass + " " + cssClass;
}

function hoverOut(obj) {
	obj.className = lastCssClass;
}

function init() {
	cancelButton.focus();
}

function clearDisplay() {
	dateDisplay.innerHTML = "";
}

function setDisplay(value) {
	dateDisplay.innerHTML = value;
}

function syncHeight() {
	var height = window.document.body.scrollHeight + 25;
	var availHeight = screen.height;
	
	external.dialogHeight = height + "px";
	external.dialogTop = ((availHeight / 2) - ((height) / 2)) - 100 + "px";
}

function closeMe() {
	window.close();
}

function handleKeys() {
	currentKeyCode = event.keyCode;
	if (typeof(datePickerFrm) != "undefined") {
		switch(currentKeyCode) {
			//escape key
			case 27 :
				closeMe();
				break;
			//right arrow key
			case 39 :
				if (!datePickerFrm.nextMonth.disabled) changeChosenDate(datePickerFrm.nextMonthValue.value);
				break;
			//left arrow key
			case 37 :
				if (!datePickerFrm.prevMonth.disabled) changeChosenDate(datePickerFrm.prevMonthValue.value);
				break;
			//up arrow key
			case 38 :
				if (!datePickerFrm.nextYear.disabled) changeChosenDate(datePickerFrm.nextYearValue.value);
				break;
			//down arrow key
			case 40 :
				if (!datePickerFrm.prevYear.disabled) changeChosenDate(datePickerFrm.prevYearValue.value);
				break;
		}
	}
}

function checkWheel() {
	upDown = event.wheelDelta / 120;
	if (upDown == 1) {
		if (!datePickerFrm.nextMonth.disabled) changeChosenDate(datePickerFrm.nextMonthValue.value);
	} else {
		if (!datePickerFrm.prevMonth.disabled) changeChosenDate(datePickerFrm.prevMonthValue.value);
	}
}

function sendDate(target, value) {
	eval("dialogArguments." + target + ".value = value");
	closeMe();
}