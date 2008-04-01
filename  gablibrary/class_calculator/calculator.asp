<!--#include file="config.asp"-->
<!--#include file="lang/en.asp"-->
<!--#include file="class_calculatorButton.asp"-->
<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		Calculator
'' @CREATOR:		Michal Gabrukiewicz - gabru @ grafix.at
'' @CREATEDON:		29.07.2004
'' @CDESCRIPTION:	Makes a cool calculator which allows the user to calculate numbers and insert them
''					into a form afterward. it also supports custom button, which can hold custom values
''					e.g. currency values from the database in order to be available for calculations
'' @VERSION:		1.0

'**************************************************************************************************************
class Calculator

	public JSTarget					''[string] the target of the field you want to input the calculated value. e.g: frm.commonField
	public defaultStylesheet		''[bool] use the default styles for this control. default = true
	public displayedValue			''[string] the value which should be displayed when the calculator is being loaded
	public commaStyle				''[string] whats the style of the comma. default is ","
	public cssLocation				''[string] virtual path to the stylesheet which is used by the calculator. defualt taken from config.asp
	
	private classLocation			
	private DEFAULT_DISPLAY_VALUE	
	private JAVASCRIPT_COMMA		
	private customButtons			
	
	'***********************************************************************************************************
	'* constructor
	'***********************************************************************************************************
	private sub class_Initialize()
		defaultStylesheet = true
		JSTarget = empty
		DEFAULT_DISPLAY_VALUE = "0"
		JAVASCRIPT_COMMA = "."
		displayedValue = empty
		commaStyle = ","
		set customButtons = Server.createObject("Scripting.Dictionary")
		cssLocation = CALC_CSS_LOCATION
		classLocation = CALC_CLASSLOCATION
	end sub
	
	'***********************************************************************************************************
	'' @SDESCRIPTION:	Draws the calculator-control. page must be a modal!
	'' @DESCRIPTION:	you need to register the following event onkeyup="handleKeys();" in the body-tag
	''					in the page you are using the calculator.
	'***********************************************************************************************************
	public sub draw()
		initJavascript()
		initStyles()
		printHeader()
		printCalculator()
		printFooter()
	end sub
	
	'***********************************************************************************************************
	'' @SDESCRIPTION:	Adds an custombutton to the calculator.
	'' @PARAM:			buttonObj [calculatorButton]: the button you want to add
	'***********************************************************************************************************
	public sub addCustomButton(buttonObj)
		customButtons.add lib.getUniqueID(), buttonObj
	end sub
	
	'***********************************************************************************************************
	'' printHeader 
	'***********************************************************************************************************
	private sub printHeader()
		with str
			.writeln("<div id=headline>")
			.writeln("	<input type=Text id=calculatorDisplay value=""" & getDisplayedValue() & """ readonly>")
			.writeln("</div>")
		end with
	end sub
	
	'***********************************************************************************************************
	'* getDisplayedValue 
	'***********************************************************************************************************
	private function getDisplayedValue()
		getDisplayedValue = convertToCustomComma(DEFAULT_DISPLAY_VALUE)
		if not displayedValue = empty then
			converted = convertToCustomComma(displayedValue)
			if isNumeric(converted) then
				getDisplayedValue = converted
			end if
		end if
	end function
	
	'***********************************************************************************************************
	'* getCalcValue 
	'***********************************************************************************************************
	private function getCalcValue()
		getCalcValue = DEFAULT_DISPLAY_VALUE
		if not displayedValue = empty then
			converted = convertToJavascriptComma(displayedValue)
			if isNumeric(converted) then
				getCalcValue = converted
			end if
		end if
	end function
	
	'***********************************************************************************************************
	'* convertToJavascriptComma 
	'***********************************************************************************************************
	private function convertToJavascriptComma(value)
		tmp = replace(value, ",", JAVASCRIPT_COMMA)
		tmp = replace(tmp, commaStyle, JAVASCRIPT_COMMA)
		convertToJavascriptComma = tmp
	end function
	
	'***********************************************************************************************************
	'* convertToCustomComma 
	'***********************************************************************************************************
	private function convertToCustomComma(value)
		tmp = replace(value, ".", commaStyle)
		tmp = replace(tmp, ",", commaStyle)
		convertToCustomComma = tmp
	end function
	
	'***********************************************************************************************************
	'* printCalculator 
	'***********************************************************************************************************
	private sub printCalculator()
		with str
			
			.writeln("<div class=calcButtons>")
			drawButton LANG_BACK, LANG_BACK_HELP, empty, "clearLastChar()", "btnBack"
			drawButton "CE", LANG_CLEAR_CURRENT_HELP, "calculatorButtonDoubleWidth", "clearCurrent()", "btnClearCurrent"
			drawButton "C", LANG_CLEAR_ALL_HELP, "calculatorButtonDoubleWidth", "clearAll()", "btnClearAll"
			.writeln("<br>")
			drawButton "7", empty, "calculatorButtonNumber", "setNumber('7')", "btnSeven"
			drawButton "8", empty, "calculatorButtonNumber", "setNumber('8')", "btnEight"
			drawButton "9", empty, "calculatorButtonNumber", "setNumber('9')", "btnNine"
			drawButton "/", empty, empty, "setOperator('/')", "btnDivide"
			drawButton "sqrt", LANG_SQUAREROOT_HELP, empty, "calcSquareroot()", "btnSqrt"
			.writeln("<br>")
			drawButton "4", empty, "calculatorButtonNumber", "setNumber('4')", "btnFour"
			drawButton "5", empty, "calculatorButtonNumber", "setNumber('5')", "btnFive"
			drawButton "6", empty, "calculatorButtonNumber", "setNumber('6')", "btnSix"
			drawButton "*", empty, empty, "setOperator('*')", "btnMulti"
			drawButton "sqr", LANG_SQUARE_HELP, empty, "calcSquare()", "btnSqr"
			.writeln("<br>")
			drawButton "1", empty, "calculatorButtonNumber", "setNumber('1')", "btnOne"
			drawButton "2", empty, "calculatorButtonNumber", "setNumber('2')", "btnTwo"
			drawButton "3", empty, "calculatorButtonNumber", "setNumber('3')", "btnThree"
			drawButton "-", empty, empty, "setOperator('-')", "btnMinus"
			drawButton "%", empty, empty, "setPercent()", "btnPercent"
			.writeln("<br>")
			drawButton "+/-", empty, empty, "switchPlusMinus()", "btnPlusMinus"
			drawButton "0", empty, "calculatorButtonNumber", "setNumber('0')", "btnZero"
			drawButton ",", empty, empty, "setComma()", "btnComma"
			drawButton "+", empty, empty, "setOperator('+')", "btnPlus"
			drawButton "=", empty, empty, "doOperation()", "btnEqual"
			.writeln("</div>")
			
			if customButtons.count > 0 then
				.writeln("<div class=customButtons>")
				for each btn in customButtons.items
					btn.value = convertToJavascriptComma(btn.value)
					btn.draw()
					.writeln("<br>")
				next
				.writeln("</div>")
			end if
			
		end with
	end sub
	
	'***********************************************************************************************************
	'* drawButton 
	'***********************************************************************************************************
	private sub drawButton(val, toolTip, cssClasses, onClick, id)
		set btn = new calculatorButton
		with btn
			.toolTip = toolTip
			.cssClass = cssClasses
			.onClick = onClick
			.buttonID = id
			.caption = val
			.customButton = false
			.draw()
		end with
	end sub
	
	'***********************************************************************************************************
	'* printFooter 
	'***********************************************************************************************************
	private sub printFooter()
		with str
			.writeln("<div id=""endline"">")
			.writeln("	<button class=button accesskey=i onclick=""sendValue('" & JSTarget & "', calculatorDisplay.value);"" title=""" & LANG_SELECT_VALUE_HELP & """>" & LANG_SELECT_VALUE & "</button>&nbsp;")
			.writeln("	<button class=button onclick=""closeMe();"" name=cancelButton title=""" & LANG_CANCEL_HELP & """>" & LANG_CANCEL & "</button>")
			.writeln("</div>")
		end with
	end sub
	
	'***********************************************************************************************************
	'* initStyles
	'***********************************************************************************************************
	private sub initStyles()
		if defaultStylesheet then lib.page.loadStylesheetFile cssLocation, empty
	end sub
	
	'***********************************************************************************************************
	'* initJavascript
	'***********************************************************************************************************
	private sub initJavascript()
		with str
			.writeln("<script language=JavaScript>")
			.writeln("	var calcValue = parseFloat(" & getCalcValue() & ")")
			.writeln("	var commaStyle = '" & commaStyle & "';")
			.writeln("</script>")
		end with
		lib.page.loadJavascriptFile classLocation & "javascript.js"
	end sub

end class
lib.registerClass("Calculator")
%>