var maximumDisplayDigits = 13;
var operand1 = 0;
var operand2 = 0;
var operator = "";
var clickedButton;
var buttonClass;
var lastAction = "";

//timeouts are here to prevent problems with settimeout-function
var timeOut = false;
var buttonTimeOut = false;

var lastCssClass;

function handleKeys() {
	currentKeyCode = event.keyCode;
	//alert(currentKeyCode);
	switch(currentKeyCode) {
		//escape key
		case 27 :
			closeMe();
			break;
		//0 key
		case 96 :
			imitateButtonClick(btnZero);
			setNumber('0');
			break;
		//1 key
		case 97 :
			imitateButtonClick(btnOne);
			setNumber('1');
			break;
		//2 key
		case 98 :
			imitateButtonClick(btnTwo);
			setNumber('2');
			break;
		//3 key
		case 99 :
			imitateButtonClick(btnThree);
			setNumber('3');
			break;
		//4 key
		case 100 :
			imitateButtonClick(btnFour);
			setNumber('4');
			break;
		//5 key
		case 101 :
			imitateButtonClick(btnFive);
			setNumber('5');
			break;
		//6 key
		case 102 :
			imitateButtonClick(btnSix);
			setNumber('6');
			break;
		//7 key
		case 103 :
			imitateButtonClick(btnSeven);
			setNumber('7');
			break;
		//8 key
		case 104 :
			imitateButtonClick(btnEight);
			setNumber('8');
			break;
		//9 key
		case 105 :
			imitateButtonClick(btnNine);
			setNumber('9');
			break;
		//, key
		case 110 :
			imitateButtonClick(btnComma);
			setComma();
			break;
		//+ key
		case 107 :
			imitateButtonClick(btnPlus);
			setOperator('+');
			break;
		//- key
		case 109 :
			imitateButtonClick(btnMinus);
			setOperator('-');
			break;
		// / key
		case 111 :
			imitateButtonClick(btnDivide);
			setOperator('/');
			break;
		//* key
		case 106 :
			imitateButtonClick(btnMulti);
			setOperator('*');
			break;
		//enter key
		case 13 :
			imitateButtonClick(btnEqual);
			doOperation();
			break;
		//backspace key
		case 8 :
			imitateButtonClick(btnBack);
			clearLastChar();
			break;
		//del key
		case 46 :
			imitateButtonClick(btnClearAll);
			clearAll();
			break;
	}
}

function imitateButtonClick(button) {
	if (!buttonTimeOut) {
		clickedButton = button;
		buttonClass = clickedButton.className;
		buttonTimeOut = true;
		clickedButton.className = buttonClass + " calculatorButtonClicked";
		setTimeout("clickedButton.className = buttonClass; buttonTimeOut = false;", 40)
	}
}

function closeMe() {
	window.close();
}

function hoverIn(obj, cssClass) {
	lastCssClass = obj.className
	obj.className = lastCssClass + " " + cssClass;
}

function hoverOut(obj) {
	obj.className = lastCssClass;
}

function sendValue(target, value) {
	eval("dialogArguments." + target + ".value = value");
	closeMe();
}

function updateDisplay() {
	display = document.getElementById("calculatorDisplay");
	calcValue = calcValue + "";
	display.value = calcValue.replace(".", commaStyle);
	display.focus();
}

function displayBlink() {
	if (!timeOut) {
		cssClass = "";
		display = document.getElementById("calculatorDisplay");
		cssClass = display.style.color;
		display.style.color = "white";
		timeOut = true;
		setTimeout("display.style.color = cssClass; timeOut = false;", 100);
	}
}

function calcValueHasComma() {
	if (calcValue.indexOf(".") > -1) {
		return true;
	} else {
		return false;
	}
}

function roundValue(myVal) {
	digits = maximumDisplayDigits;
	beforeComma = (myVal + "").indexOf(".");
	if (beforeComma > -1) {
		digits = digits - beforeComma - 1;
	}
	return Math.round(myVal * Math.pow(10, digits)) / Math.pow(10, digits);
}

function trimToAllowedLength(val) {
	if (val.length > maximumDisplayDigits) {
		return val.substr(0, maximumDisplayDigits);
	} else {
		return val;
	}
}

/***********************************************************************************
* button actions
***********************************************************************************/

function setNumber(number) {
	if (calcValue == '0') {
		calcValue = number;
	} else {
		calcValue = "" + calcValue + number;
	}
	calcValue = trimToAllowedLength(calcValue);
	updateDisplay();
	lastAction = number;
}

function setComma() {
	if (calcValue == '0') {
		calcValue = "0.";
	} else {
		if (!calcValueHasComma()) {
			calcValue = "" + calcValue + ".";
		}
	}
	calcValue = trimToAllowedLength(calcValue);
	updateDisplay();
	lastAction = "setComma";
}

function doOperation() {
	//we check if the users wants to do one operation more than once.
	//you know like in a calculator. 5 + 2 -> enter -> +2 -> enter +2, etc.
	if (lastAction != "enter") {
		operand2 = calcValue;
	} else {
		operand1 = calcValue;
	}
	
	if ((operator != "" && operator != "/") || (operator == "/" && operand2 != "0")) {
		eval("calcValue = roundValue(" + operand1 + " " + operator + " " + operand2 + ");");
		updateDisplay();
	}
	
	lastAction = "enter";
}
var sqrCalc = 0;
function setOperator(op) {
	//if you have pressed any operator instead of enter - calculate value
	if(operand1 != 0 && !isNaN(parseInt(lastAction))) {
		doOperation();
	}

	//we check the last keystroke because maybe the user changes the operator more than once
	if (lastAction != "+" && lastAction != "-" && lastAction != "/" && lastAction != "*") {
		operand1 = calcValue;
	}
	
	operator = op;
	if(calcValue != 0)
		sqrCalc = calcValue;
	calcValue = "0";
	displayBlink();
	
	lastAction = op;

}

function clearCurrent() {
	calcValue = '0';
	displayBlink();
	updateDisplay();
	lastAction = "clearCurrent";
}

function clearAll() {
	operator = "";
	operand1 = 0;
	operand2 = 0;
	clearCurrent();
	lastAction = "clearAll";
}

function clearLastChar() {
	calcValue = calcValue + "";
	if (calcValue.length > 1) {
		calcValue = (calcValue).substr(0, calcValue.length - 1);
	} else {
		calcValue = "0";
	}
	updateDisplay();
	lastAction = "clearLastChar";
}

function setPercent() {
	calcValue = roundValue(operand1 * (calcValue / 100));
	updateDisplay();
	lastAction = "setPercent";
}

function switchPlusMinus() {
	checkCalcValue();
	
	calcValue = calcValue * -1
	updateDisplay();
	lastAction = "switchPlusMinus";
}

function calcSquareroot() {
	checkCalcValue();

	//its not possible to calculate negative squareroot
	if (0 + calcValue >= 0) {
		calcValue = roundValue(Math.sqrt(calcValue));
		updateDisplay();
	}
	lastAction = "calcSquareroot";
}

function calcSquare() {
	checkCalcValue();
		
	calcValue = roundValue(calcValue * calcValue);
	updateDisplay();
	lastAction = "calcSquare";
}

function setCustomValue(val) {
	calcValue = roundValue(val);
	updateDisplay();
	lastAction = "setCustomValue";
}

function checkCalcValue() {
	if(calcValue == 0) {
		calcValue = sqrCalc;
		sqrCalc = 0;
	}
}