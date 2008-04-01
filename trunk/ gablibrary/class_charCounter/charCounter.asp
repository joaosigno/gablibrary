<!--#include file="const.asp"-->
<!--#include file="charCounterTextarea.asp"-->
<!--#include file="charCounterTextfield.asp"-->
<%
'**************************************************************************************************************

'' @CLASSTITLE:		CharCounter
'' @CREATOR:		Michal Gabrukieiwcz - gabru@gmx.at
'' @CREATEDON:		26.09.2003
'' @CDESCRIPTION:	Easily create a counter bar to visualize how many chars left. The user can see how much
''					he/she is allowed to write.
'' @VERSION:		0.1

'**************************************************************************************************************
class charCounter

	private ControlToUse
	private barColorCssClass	'Stylesheet-class for the BAR-color.
	private barBgColorCssClass	'Stylesheet-class for the BAR-Backgroundcolor
	
	public maxChars				''The amount of maximum allowed chars.
	public barLength			''The length of you counter-bar.
	public iAmNotAlone			''true/false - if you use more than one charCounter in your Site then you should
								''set only the first on false and all other on true. Just one Javascript function
								''will be placed in your document.
	
	'Constructor => set the default values
	private sub Class_Initialize()
		set ControlToUse	= Server.createObject("Scripting.Dictionary")
		maxChars			= 0
		barLength			= 0
		iAmNotAlone			= false
		barColorCssClass	= BARCOLORCSSCLASS
		barBgColorCssClass	= BARBGCOLORCSSCLASS
	end sub
	
	'************************************************************************************************************
	''@SDESCRIPTION:	Use this function to allocate an control-object to the charCounter.
	''@PARAM:			- controlObj: Control-Object e.g. a Textarea, textfield, etc.
	'************************************************************************************************************
	public sub allocateControl(controlObj)
		ControlToUse.removeAll
		ControlToUse.Add 1, controlObj
	end sub
	
	'************************************************************************************************************
	''@SDESCRIPTION:	Draws the control you have allocated to the counter
	'************************************************************************************************************
	public sub drawControl()
		ControlToUse.item(1).draw
		drawHiddenFields()
	end sub
	
	'************************************************************************************************************
	''@SDESCRIPTION:	Draws hidden field to store the maximum value
	'************************************************************************************************************
	private sub drawHiddenFields()
		response.write "<INPUT type=Hidden value=""" & maxChars & """ name=""dynamicCharCounterMaxValue_" & ControlToUse.item(1).name & """>"
	end sub
	
	'************************************************************************************************************
	''@SDESCRIPTION:	Draws the counter Bar of your dreams :)
	'************************************************************************************************************
	public sub drawCounter()
		controlName = ControlToUse.item(1).name
		with response
			.write "<TABLE cellpadding=0 border=0 cellspacing=0>" & vbcrlf
			.write "<TR>"
			.write "	<TD valign=middle>" & vbcrlf
		    .write "		<TABLE cellpadding=0 cellspacing=0 border=0 width=""" & barLength & """ id=""charCounterTable_" & controlName & """>" & vbcrlf
		    .write "		<TR>" & vbcrlf
		    .write "			<TD class=" & barColorCssClass & " width=""0""><IMG src=""" & CHARCOUNTERCLASSLOCATION & "images/trans.gif"" name=""charCounter_used_" & controlName & """ height=5 width=0 alt=""" & TXTCHARSAVAILABLE & """></TD>" & vbcrlf
		    .write "			<TD class=" & barBgColorCssClass & " width=""100%""><IMG src=""" & CHARCOUNTERCLASSLOCATION & "images/trans.gif"" name=""charCounter_unused_" & controlName & """ height=5 alt=""" & TXTCHARSLEFT & """ style=""width:100%;""></TD>" & vbcrlf
		    .write "		</TR>" & vbcrlf
		    .write "		</TABLE>" & vbcrlf
		    .write "	</TD>"
			.write "</TR>" & vbcrlf
			.write "</TABLE>" & vbcrlf
		end with
		if not iAmNotAlone then
			%><!--#include file="javascript.js"--><%
		end if
	end sub
	
	'************************************************************************************************************
	''@SDESCRIPTION:	Executes the Javascript on pageload. So the counter will be updated on first pageload
	'************************************************************************************************************
	public sub executeJS()
		response.write "<SCRIPT language=""JavaScript"">"
		response.write "	dynamicCharCounterJSFunction(document." & ControlToUse.item(1).formName & "." & ControlToUse.item(1).name & ",'" & ControlToUse.item(1).formName & "');"
		response.write "</SCRIPT>"
	end sub

end class
lib.registerClass("CharCounter")
%>