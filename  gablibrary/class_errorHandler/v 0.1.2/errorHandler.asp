<!--#include file="const.asp"-->
<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		errorHandler
'' @CREATOR:		Michal Gabrukiewicz - gabru@gmx.at
'' @CREATEDON:		20.11.2003
'' @CDESCRIPTION:	This class handles all the errors which happen in the classes. You can use this class
''					also by yourself to catch the errors. It will allow you to store the errors in your
''					database and/or send error by email to webadmins. 
'' @VERSION:		0.1.2

'**************************************************************************************************************
class errorHandler

	private emailContent		'Stores output for email
	private output				'Stores output for user
	private errorDesc			'Error description
	private errorSource			'Error Source
	private p_errorObject		'My Error Object (err)
	private p_debuggingVar		'debuggingvar
	
	public errorDuring			''[string] During what did the error happen. e.g. executing Sql-query
	public notifyViaMail		''[bool] Send Email to all WebAdmins with error details?
	public HTMLemail			''[bool] Should the Notify Email be HTML formatted?
	public logging				''[bool] Log this error to database?
	public additionalInfo		''[string] Add additional-information about the error if you want.
								''Its just for the user. Wont be logged or sent or anything else. Just displayed.
	
	'Constructor => set the default values
	private sub Class_Initialize()
		emailContent		= empty
		output				= empty
		errorDuring			= empty
		p_debuggingVar		= empty
		set p_errorObject	= nothing
		notifyViaMail		= consts.email_send_on_error
		logging				= true
		HTMLemail			= true
		additionalInfo		= empty
	end sub
	
	public property let debuggingVar(value) ''[variable] Sets a variable you want to display with the error. So it is faster for debugging. e.g. an Sql-query. So will see the variable when the error occurs.
		p_debuggingVar = server.HTMLEncode(value)
	end property
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Sets the error-object. 
	'' @PARAM:			- [err-object] If you use "on error resume" you will get an error-object which is needed here
	'******************************************************************************************************************
	public sub errorObject(obj)
		set p_errorObject = obj
	end sub
	
	'******************************************************************************************************************
	'* printFormCollection 
	'******************************************************************************************************************
	private function printFormCollection()
		c = 0
		tmp = empty
		for each field in request.form
			if not c = 0 then
				 tmp = tmp & " --- "
			end if
			tmp = tmp & field & ": " & request.form(field)
			c = c + 1
		next
		printFormCollection = tmp
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	sends an email with the error and returns the status
	'******************************************************************************************************************
	public function sendErrorEmail()
		toReturn = empty
		
		if notifyViaMail then
			add_EmailContent "<span style=font-family:verdana;font-size:10pt;>", empty, empty
			add_EmailContent "<strong>" & TXT_MAIL_HEADLINE_GENERAL & "</strong><BR><BR>", empty, empty
			add_EmailContent "<strong>" & TXT_MAIL_DURING & ": </strong>", errorDuring, "<BR>"
			add_EmailContent "<strong>" & TXT_MAIL_DEBUGVAR & ": </strong>", p_debuggingVar, "<BR>"
			add_EmailContent "<strong>" & TXT_MAIL_ERRDESCRIPTION & ": </strong>", errorDesc ,"<BR>"
			add_EmailContent "<strong>" & TXT_MAIL_ERRSOURCE & ": </strong>", errorSource, "<BR>"
			add_EmailContent "<strong>" & TXT_MAIL_REQUESTED_URL & ": </strong>", "http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL"), "<BR>"
			add_EmailContent "<strong>" & TXT_MAIL_REFFERED_URL & ": </strong>", Request.ServerVariables("HTTP_REFERER"), "<BR>"
			add_EmailContent "<strong>" & TXT_MAIL_QUERYSTRING_COLLECTION & ": </strong>", request.queryString, "<BR>"
			add_EmailContent "<strong>" & TXT_MAIL_FORM_COLLECTION & ": </strong>", printFormCollection(), "<BR>"
			add_EmailContent "<BR><BR><strong>" & TXT_MAIL_HEADLINE_SESSIONS & "</strong><BR><BR>", showAllSessionVars(HTMLemail), "</span>"
			
			set mailer = new customSendMail
			with mailer
				.subject = TXT_SUBJECTPREFIX & " - "& Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL")
				.fromEmail = consts.automated_bot_email
				.fromName = consts.automated_bot_name
				.body = emailContent
				
				if HTMLemail then
					.htmlEmail = true
				end if
				
				'add all webadmins as recipients
				myArray = consts.webadmins
				for i = 0 to ubound(myArray)
					.addRecipient "TO", myArray(i,1), myArray(i,0)
				next
				
				if not .send() then
					toReturn = "<BR><em>" & TXT_ERROR_ON_SEND_NOTIFICATION_MAIL & ":<BR>" & .errormessage & "</em>"
				else
					toReturn = toReturn & "<BR><em>" & TXT_ERROR_MAILED_TO & ":<BR>"
					myArr = consts.webadmins
					for i = 0 to ubound(myArr)
						if not i = 0 then toReturn = toReturn & " ," end if
						toReturn = toReturn & "<a href=mailto:" & myArr(i,1) & ">" & myArr(i,0) & "</a>"
					next
					toReturn = toReturn & "<BR>" & TXT_AND_FIXED_SOON & "</em></BODY></HTML>"
				end if
				
			end with
			set mailer = nothing
		else
			toReturn = "<BR><em>" & TXT_EMAIL_NOTIFICATION_DISABLED & "</em>"
		end if
		
		sendErrorEmail = toReturn
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Logs the error for example to database or a file.
	'' @DESCRIPTION:	It is an interface for the method logError from customLib-class. You have to implement this
	''					method yourself in customLib. The following parameter will be available: Chronological order of parameters<BR>
	''					1. errorDuring [string] (during what did the error happend)<BR>
	''					2. errorDescription [string] (the description of the error-object)<BR>
	''					3. errorSource [string] (whats the source of the error)<BR>
	''					4. debuggingVariable [variable] (the debugging variable)<BR>
	''					5. requestedUrl [string] (incl. QueryString - on what page did the error happend)<BR>
	''					6. refferedUrl [string] (whats the referring url)<BR>
	''					7. sessionVariables [string] (all session-variables. HTMLformatted)<BR>
	''					The method should return true/false wheater the logging-procedure was successfull or not.
	'' @RETURN:			[bool] logging successfully or not
	'******************************************************************************************************************
	public function logError()
		if not request.QueryString = empty then
			QS = "?" & request.queryString
		else
			QS = empty
		end if
		
		if lib.custom.logError(errorDuring, errorDesc, errorSource, p_debuggingVar,_
							"http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL") & QS,_
							Request.ServerVariables("HTTP_REFERER"), showAllSessionVars(true)) then
			logError = true
		else
			logError = false
		end if
	end function
	
	'******************************************************************************************************************
	'* add_Output 
	'******************************************************************************************************************
	private sub add_Output(value)
		output = output & value
	end sub
	
	'******************************************************************************************************************
	'* add_EmailContent 
	'******************************************************************************************************************
	private sub add_EmailContent(beginningTags, values, endingTags)
		if HTMLemail then
			emailContent = emailContent & beginningTags & values & endingTags
		else
			emailContent = emailContent & values & chr(13)
		end if
	end sub
	
	'******************************************************************************************************************
	'* initStyles 
	'******************************************************************************************************************
	private function initStyles()
		tmp = empty
		tmp = tmp & "<STYLE>" & vbcrlf
		tmp = tmp & "	body 		{ font-family:""Tahoma""; font-size:12pt; }" & vbcrlf
		tmp = tmp & "	a			{ color:#0000FF; }" & vbcrlf
		tmp = tmp & "	a:visited	{ color:#0000FF; }" & vbcrlf
		tmp = tmp & "	a.active	{ color:#0000FF; }" & vbcrlf
		tmp = tmp & "	a:hover	{ color:#FF0000; }" & vbcrlf
		tmp = tmp & "</STYLE>" & vbcrlf
		initStyles = tmp
	end function
	
	'******************************************************************************************************************
	'' init_errorObject 
	'******************************************************************************************************************
	private sub init_errorObject()
		if typename(p_errorObject) = "Object" then
			errorDesc = cstr(p_errorObject.number) & " - " & p_errorObject.description
			errorSource = p_errorObject.source
		else
			errorDesc = empty
			errorSource = empty
		end if
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Draws the occured error-message with all the information needed. 
	'******************************************************************************************************************
	public sub draw()
		init_errorObject()
		
		add_Output initStyles()
		add_Output "<HTML><HEAD><TITLE>" & TXT_ERRORDURING & " " & errorDuring & "</TITLE></HEAD>"
		add_Output "<BODY bgcolor=""#FFFFFF"">"
		add_Output "<img src='" & consts.logo & "' alt=""" & consts.company_name & """ border=0 align=absmiddle>"
		add_Output "&nbsp;&nbsp;&nbsp;&nbsp;<span style=color:#FF2727;><strong>" & TXT_ERROR & "</strong></span><BR><BR>"
		add_Output "<strong>" & TXT_ERRORDURING & " " & errorDuring & ":</strong><div>" & p_debuggingVar & "</div>"
		
		if not additionalInfo = empty then
			add_Output "<BR>"
			add_Output "<DIV style=""background-color:#DDD;padding:10px;"">" & additionalInfo & "</DIV>"
		end if
		
		if typename(p_errorObject) = "Object" then
			add_Output "<BR>"
			add_Output "<strong>" & TXT_ERRORCODE & ":</strong> " & errorDesc & "<br>"
			add_Output "<strong>" & TXT_SOURCE & ":</strong> " & errorSource & "<br>"
		end if
		
		add_Output sendErrorEmail()
		
		if typename(p_errorObject) = "Object" then
			p_errorObject.clear()
			set p_errorObject = nothing
		end if
		
		if logging then
			add_Output "<BR><BR><span style='font-size:8pt;'>ErrorLog - "
			if logError() then
				add_Output "done."
			else
				add_Output "<strong>failed.</strong>"
			end if
			add_Output "</span>"
		end if
		
		response.write output & "<BR><BR>"
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Returns all Session-variables.
	'' @DESCRIPTION:	Also checks if there are Arrays in a session-var and displays every item of the array
	'' @PARAM:			- HTMLformatted [bool]: Should the output be html formatted or not?
	'' @RETURN:			[string]
	'******************************************************************************************************************
	public function showAllSessionVars(HTMLformatted)
		tp = empty
		for each field in session.contents
			'if there is an array in a session-var we catch it. cause if not then the whole thing wont work!
			if (isArray(session(field))) then
				for i = 0 to ubound(session(field))
					myArr = session(field)
					if i = 0 then
						if HTMLformatted then
							tp = tp & "<strong>" & field & "</strong>"
							tp = tp & "<em>(array)</em>"
						else
							tp = tp & field & "(array)"
						end if
						tp = tp & ": (" & i & ") " & myArr(i)
					else
						tp = tp & ", (" & i & ") " & myArr(i)
					end if
				next
				if HTMLformatted then
					tp = tp & "<BR>"
				else
					tp = tp & chr(13)
				end if
			else
				if not isObject(session(field)) then
					sessionContent = session(field)
				else
					sessionContent = "{Object: " & typename(session(field)) & "}"
				end if
				
				if HTMLformatted then
					tp = tp & "<strong>" & field & ": </strong>" & sessionContent & "<BR>"
				else
					tp = tp & field & ": " & sessionContent & chr(13)
				end if
			end if
		next
		showAllSessionVars = tp
	end function

end class
%>