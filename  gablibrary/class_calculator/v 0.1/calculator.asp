<!--#include file="language.asp"-->
<!--#include file="class_calculatorButton.asp"-->
<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		calculator
'' @CREATOR:		Michal Gabrukiewicz - gabru@gmx.at
'' @CREATEDON:		29.07.2004
'' @CDESCRIPTION:	Draws a cool calculator
'' @VERSION:		0.1

'**************************************************************************************************************
class calculator

	public JSTarget					''[string] the target of the field you want to input the calculated value. e.g: frm.commonField
	public defaultStylesheet		''[bool] use the default styles for this control. default = true
	public displayedValue			''[string] the value which should be displayed when the calculator is being loaded
	public commaStyle				''[string] whats the style of the comma. default is ","
	
	private classLocation			
	private DEFAULT_DISPLAY_VALUE	
	private JAVASCRIPT_COMMA		
	private customButtons			
	
	'Konstruktor => set the default values
	private sub Class_Initialize()
		defaultStylesheet			= true
		JSTarget					= empty
		DEFAULT_DISPLAY_VALUE		= "0"
		JAVASCRIPT_COMMA			= "."
		displayedValue				= empty
		commaStyle					= ","
		classLocation				= "/gab_Library/class_calculator/"
		set customButtons			= Server.createObject("Scripting.Dictionary")
	end sub
	
	'***********************************************************************************************************
	'' @SDESCRIPTION:	Draws the calculator-control
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
	'' @PARAM:			- buttonObj [calculatorButton]
	'***********************************************************************************************************
	public sub addCustomButton(buttonObj)
		customButtons.add lib.getUniqueID(), buttonObj
	end sub
	
	'***********************************************************************************************************
	'' printHeader 
	'***********************************************************************************************************
	private sub printHeader()
		with str
			.writeln("<div class=headline>")
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
			call drawButton(LANG_BACK, LANG_BACK_HELP, empty, "clearLastChar()", "btnBack")
			call drawButton("CE", LANG_CLEAR_CURRENT_HELP, "calculatorButtonDoubleWidth", "clearCurrent()", "btnClearCurrent")
			call drawButton("C", LANG_CLEAR_ALL_HELP, "calculatorButtonDoubleWidth", "clearAll()", "btnClearAll")
			.writeln("<br>")
			call drawButton("7", empty, "calculatorButtonNumber", "setNumber('7')", "btnSeven")
			call drawButton("8", empty, "calculatorButtonNumber", "setNumber('8')", "btnEight")
			call drawButton("9", empty, "calculatorButtonNumber", "setNumber('9')", "btnNine")
			call drawButton("/", empty, empty, "setOperator('/')", "btnDivide")
			call drawButton("sqrt", LANG_SQUAREROOT_HELP, empty, "calcSquareroot()", "btnSqrt")
			.writeln("<br>")
			call drawButton("4", empty, "calculatorButtonNumber", "setNumber('4')", "btnFour")
			call drawButton("5", empty, "calculatorButtonNumber", "setNumber('5')", "btnFive")
			call drawButton("6", empty, "calculatorButtonNumber", "setNumber('6')", "btnSix")
			call drawButton("*", empty, empty, "setOperator('*')", "btnMulti")
			call drawButton("sqr", LANG_SQUARE_HELP, empty, "calcSquare()", "btnSqr")
			.writeln("<br>")
			call drawButton("1", empty, "calculatorButtonNumber", "setNumber('1')", "btnOne")
			call drawButton("2", empty, "calculatorButtonNumber", "setNumber('2')", "btnTwo")
			call drawButton("3", empty, "calculatorButtonNumber", "setNumber('3')", "btnThree")
			call drawButton("-", empty, empty, "setOperator('-')", "btnMinus")
			call drawButton("%", empty, empty, "setPercent()", "btnPercent")
			.writeln("<br>")
			call drawButton("+/-", empty, empty, "switchPlusMinus()", "btnPlusMinus")
			call drawButton("0", empty, "calculatorButtonNumber", "setNumber('0')", "btnZero")
			call drawButton(",", empty, empty, "setComma()", "btnComma")
			call drawButton("+", empty, empty, "setOperator('+')", "btnPlus")
			call drawButton("=", empty, empty, "doOperation()", "btnEqual")
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
			.writeln("<div class=""endline"">")
			.writeln("	<button class=button onclick=""sendValue('" & JSTarget & "', calculatorDisplay.value);"" title=""" & LANG_SELECT_VALUE_HELP & """>" & LANG_SELECT_VALUE & "</button>&nbsp;")
			.writeln("	<button class=button onclick=""closeMe();"" name=cancelButton title=""" & LANG_CANCEL_HELP & """>" & LANG_CANCEL & "</button>")
			.writeln("</div>")
		end with
	end sub
	
	'***********************************************************************************************************
	'* initStyles
	'***********************************************************************************************************
	private sub initStyles()
		if defaultStylesheet then
			str.writeln("<link rel=stylesheet type=text/css href=""" & classLocation & "standard.css"">")
		end if
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
		str.writeln("<script language=JavaScript src=""" & classLocation & "javascript.js""></script>")
	end sub

end class
%>