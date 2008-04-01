<!--#include virtual="/gab_LibraryConfig/_form.asp"-->
<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003 - This file is part of GAB_LIBRARY		
'* For license refer to the license.txt in the root    						
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		Form
'' @CREATOR:		David Rankin
'' @CREATEDON:		2006-10-23 09:41
'' @CDESCRIPTION:	A few tools, such as the drawError, print button, cancel buttons (More to come), that we 
''					commonly use with forms are found in here
'' @REQUIRES:		Validateable
'' @VERSION:		0.2
'' @POSTFIX:  		frm

'**************************************************************************************************************
class Form

	'public members
	public printIconPath	''[string] the virtual path to the print icon. default is taken from the gablib icons
	public errorsTitle		''[String] the title for the Errors that are drawed using drawError()
	public validator		''[validateble] The validator you use in your page.
	public message			''[variant] some message the form should store. usefull for error-, succes-messages, etc. 
							''hint: use an array if you want to store the cssclass also ;)
	
	public property get action ''[string] gets the action needed for the form to submit. full qualified filename incl. the querystring
		action = request.serverVariables("SCRIPT_NAME") & "?" & lib.QS("")
	end property
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		lib.require("Validateable")
		printIconPath = lib.init(GL_FORM_PRINTICON, consts.STDAPP("icons/print.gif"))
		errorsTitle = TXT_FRM_ERROR
		set validator = new Validateable
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	If your form is not valid, use this to display neat error messages
	'' @DESCRIPTION:	This displays errors similar to the error handler in .Net. We loop through all the
	''					errors in the validator, if a match found, we display the error
	'' @PARAM:			name [string], [array]: The name of your field(s) to draw if and error is found
	'**********************************************************************************************************
	public sub drawError(name)
		arr = name
		if not isArray(name) then arr = array(name)
		for i = 0 to ubound(arr)
			errorMessage = validator.getInvalidDescriptionFor(arr(i))
			if errorMessage <> "" then exit for
		next
		tip = " title=""" & str.HTMLEncode(errorsTitle & " " & errorMessage) & """"
		'are we using tooltips
		if lib.page.loadTooltips then tip = lib.tooltip(errorsTitle, errorMessage)
		if errorMessage <> "" then str.writeln("<img src=""" & consts.STDAPP("icons/form_error.gif") & """ align=absmiddle" & tip & ">")
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	Checks if an error occured on a specific control
	'' @PARAM:			name [string], [array]: The name of your field(s) to draw if and error is found
	'' @RETURN:			[bool] true if an error occured
	'**********************************************************************************************************
	public function errorOccuredOn(name)
		if not isArray(name) then arr = array(name)
		for i = 0 to ubound(arr)
			errorOccuredOn = (validator.getInvalidDescriptionFor(arr(i)) <> "")
			if errorMessage <> "" then exit for
		next
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	Draws the Print button with the print icon
	'**********************************************************************************************************
	public sub drawPrintButton()
		str.write("<button title=""" & TXT_FRM_PRINT_TT & """ class=""button"" onclick=""window.print()"">")
		str.write("<img src=""" & printIconPath & """ align=""absmiddle"">")
		str.write(" " & TXT_FRM_PRINT & "</button>")
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	Draws a cancel button for the form.
	'' @PARAM:			url [string]: The url you want the cancel button to send the user to!
	''					this parameter is ignored when the form is in a modalDialog. because then it will
	''					automatically make a window.close() action
	'**********************************************************************************************************
	public sub drawCancelButton(url)
		if lib.page.isModalDialog then
			onclick = "window.close()"
		else
			onclick = "window.location.href='" & url & "'"
		end if
		str.write("<button class=""button"" onclick=""" & onclick & """>" & TXT_FRM_CANCEL & "</button>")
	end sub

end class
lib.registerClass("Form")
%>