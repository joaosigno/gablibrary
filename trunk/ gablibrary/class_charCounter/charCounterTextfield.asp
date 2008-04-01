<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		charCounterTextfield
'' @CREATOR:		Michal Gabrukieiwcz - gabru@gmx.at
'' @CREATEDON:		03.10.2003
'' @CDESCRIPTION:	Thats a textfield for the charcounter
'' @VERSION:		0.1

'**************************************************************************************************************
class charCounterTextfield

	public name				''Name of the textfield
	public attributes		''If you want to add some Attributes yourself. e.g. style="...."
	public value			''You have a value for it?
	public disabled			''true/false
	public formName			''Name of the form
	
	'Constructor => set the default values
	private sub Class_Initialize()
		name		= empty
		attributes	= empty
		value		= empty
		disabled	= false
	end sub
	
	'************************************************************************************************************
	''@SDESCRIPTION:	Draws the textfield
	'************************************************************************************************************
	public sub draw()
		if disabled then
			myDisabled = " disabled"
		else
			myDisabled = empty
		end if
		response.write _
			"<INPUT " &_
				"name=""" & name  & """ value=""" & value & """ " & attributes & " " &_
				"onKeyUp=""dynamicCharCounterJSFunction(this,'" & formName & "');"" " &_
				"onKeyDown=""dynamicCharCounterJSFunction(this,'" & formName & "');"" " &_
				"onKeyFocus=""dynamicCharCounterJSFunction(this,'" & formName & "');"" " &_
				"onChange=""dynamicCharCounterJSFunction(this,'" & formName & "');"" " &_
				myDisabled & ">" & vbcrlf
	end sub

end class
%>