<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		calculatorButton
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		30.07.2004
'' @CDESCRIPTION:	A custom Button for the calculator-Control
'' @VERSION:		0.1

'**************************************************************************************************************
class calculatorButton

	public caption				''[string] caption which will be displayed on the button
	public toolTip				''[string] tooltip for the button
	public value				''[string] what value should the button stand for
	public onClick				''[string] JS function which should be called onClick. Leave empty if custom-button
	public cssClass				''[string] CssClass for the button
	public buttonID				''[string] id of the button. leave empty if custom-button
	public customButton			''[bool] is this a custom-button? default=true
	
	'Konstruktor => set the default values
	private sub Class_Initialize()
		caption					= empty
		tooltip					= empty
		value 					= empty
		onClick					= empty
		cssClass				= empty
		buttonID				= empty
		customButton			= true
	end sub
	
	private property get id()
		if customButton then
			id = "btnCustom_" & lib.getUniqueID()
		else
			id = buttonID
		end if
	end property
	
	private property get cssClasses()
		cssClasses = "calculatorButton " & cssClass
		if customButton then
			cssClasses = cssClasses & " calculatorButtonDoubleWidth calculatorButtonCustom"
		end if
	end property
	
	'***********************************************************************************************************
	'' @SDESCRIPTION:	Draws the button
	'***********************************************************************************************************
	public sub draw()
		if customButton then
			onClick = "setCustomValue('" & value & "')"
		end if
		
		str.writeln("<button id=" & id & " class=""" & cssClasses & """ title=""" & toolTip & """ " &_
						"onmouseover=""hoverIn(this, 'calculatorButtonHover');"" " &_
						"onmouseout=""hoverOut(this);"" onclick=""" & onClick & ";"" tabindex=""-1"">" & caption & "</button>")
	end sub

end class
%>