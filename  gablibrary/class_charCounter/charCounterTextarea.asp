<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		charCounterTextarea
'' @CREATOR:		Michal Gabrukieiwcz - gabru@gmx.at
'' @CREATEDON:		26.09.2003
'' @CDESCRIPTION:	Thats the textarea for a charcounter
'' @VERSION:		0.1

'**************************************************************************************************************
class charCounterTextarea

	public name				''Name of the textarea
	public cols				''How nany Columns
	public rows				''How many Rows
	public attributes		''If you want to add some Attributes yourself. e.g. style="...."
	public value			''You have a value for it?
	public disabled			''true/false
	public formName			''Name of the form
	
	'Constructor => set the default values
	private sub Class_Initialize()
		name		= empty
		cols		= 0
		rows		= 0
		attributes	= empty
		value		= empty
		disabled	= false
	end sub
	
	'************************************************************************************************************
	''@SDESCRIPTION:	Draws the textarea
	'************************************************************************************************************
	public sub draw()
		if disabled then
			myDisabled = " disabled"
		else
			myDisabled = empty
		end if
		response.write _
			"<TEXTAREA " &_
				"cols=""" & cols & """ " &_
				"rows=""" & rows & """ " &_
				"name=""" & name  & """ " & attributes & " " &_
				"onKeyUp=""dynamicCharCounterJSFunction(this,'" & formName & "');"" " &_
				"onKeyDown=""dynamicCharCounterJSFunction(this,'" & formName & "');"" " &_
				"onKeyFocus=""dynamicCharCounterJSFunction(this,'" & formName & "');"" " &_
				"onChange=""dynamicCharCounterJSFunction(this,'" & formName & "');"" " &_
				myDisabled &_
			">" & value & "</TEXTAREA>" & vbcrlf
	end sub

end class
%>