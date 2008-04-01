<!--#include virtual="/gab_LibraryConfig/_errorHandler.asp"-->
<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		ErrorHandler
'' @CREATOR:		Michal Gabrukiewicz - gabru @ grafix.at
'' @CREATEDON:		20.11.2003
'' @CDESCRIPTION:	This class is here to handle errors which are defined by the gabLib and the user.
''					You can use this class also by yourself to catch the errors. It sends automatically
''					an email to all webadmins when an error occured and email is configured. Furthermore
''					the consts.email_send_on_error must be enabled.
''					Moreover it is possible to store the error e.g. to a database by implementing the
''					method lib.custom.logError() method.
'' @VERSION:		1.0
'' @REQUIRES:		TextTemplate

'**************************************************************************************************************
class ErrorHandler

	private feedbackURL
	
	public errorDuring			''[string] During what did the error happen. e.g. executing Sql-query
								''Its just for the user. Wont be logged or sent or anything else. Just displayed.
	public notifyViaMail		''[bool] Send Email to all WebAdmins with error details?
	public logging				''[bool] Log this error? calls lib.custom.logError() if true. default = taken from consts.errorLogging
	public debuggingVar			''[string] Sets a variable you want to display with the error. So it is faster for debugging.
								''e.g. an Sql-query. So will see the variable when the error occurs.
								''by defualt the value from the sessions named "GL_lastErrorHelperVar" is taken.
	public errorObject			''[error], [aspError] sets the errorobject if available. data of these errors will be used for mail and log
								''default is nothing, so no additional information about the error will be displayed.
								''normally get with server.getLastError()
	public alternativeText		''[string] alternative message which will be displayed when no errorObject is available.
	public cssLocation			''[string] location of the stylesheet which will be taken for formatting the errorhandler on drawing.
								''default is taken from the config
	
	public property get executingFile ''[string] gets the fullname of the file which has been executed when the error was generated (incl. Querystring)
		executingFile = request.serverVariables("SCRIPT_NAME") & lib.iif(lib.QS("") <> "", "?" & lib.QS(""), "")
	end property
	
	'******************************************************************************************************************
	'* constructor 
	'******************************************************************************************************************
	private sub class_Initialize()
		errorDuring = "code execution"
		notifyViaMail = consts.email_send_on_error
		logging	= consts.errorLogging
		debuggingVar = session("GL_lastErrorHelperVar")
		set errorObject = nothing
		alternativeText = ""
		cssLocation = lib.init(GL_EH_CSSLOCATION, consts.gabLibLocation & "class_errorHandler/std.css")
		feedbackURL = lib.init(GL_EH_FEEDBACKURL, "/feedback/?for={0}&title={1}")
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Generates the handler. this means that all actions will be done. drawing, mailing, logging etc.
	'' @DESCRIPTION:	use this method to do all actions. alternatively each action can be executed manually.
	''					first the error is drawn, then mail is sent and last but not least the log is done.
	'******************************************************************************************************************
	public sub generate()
		draw()
		sendMail()
		me.log()
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Draws the error with the set properties. 
	'' @DESCRIPTION:	if buffering is on then the buffer will be cleared thus the error will be shown without
	''					any already rendered content. this prevents rendering errors (e.g. if error happen within an
	''					unclosed HTML-tag.
	'******************************************************************************************************************
	public sub draw()
		if response.buffer then
			response.clear()
			'if we cleared the buffer then we need the basic html-tags
			str.write("<html><head><meta http-equiv=Content-Type content=""text/html; charset=utf-8"">")
			str.write("<title>" & TXT_ERROR_TITLE & "</title></head><body>")
			lib.page.loadJSCore()
		end if
		
		lib.page.loadStylesheetFile cssLocation, ""
		
		with str
			.write("<div id=GLErrorContainer>")
			.write("<div id=GLErrorHeadline onclick=""alert(document.getElementById('GLErrorDetails').innerText)"">")
			.write("<img src=" & consts.STDAPP("icons/exclamation.gif") & " align=absmiddle>")
			.writeln(TXT_ERROR_TITLE & " (" & errorDuring & ")")
			.write("</div>")
			.write("<div id=GLErrorDetailsContainer><div id=GLErrorDetails>")
			if errorObject is nothing and alternativeText <> "" then
				.writeln("<p>" & alternativeText & "</p>")
			elseif not errorObject is nothing and (lib.custom.isWebadmin() or consts.isDevelopment()) then
				isASPError = lib.iif(typename(errorObject) = "IASPError", true, false)
				if isASPError then
					drawValue "", errorObject.category
					drawValue "", errorObject.file
				end if
				drawValue "", errorObject.description
				if isASPError then
					drawValue "", str.format(TXT_ERROR_POS, array(errorObject.line, errorObject.column))
				end if
				.write("<br>")
				drawValue TXT_ERROR_NUMBER, errorObject.number
				if isASPError then
					drawValue TXT_ERROR_CODE, errorObject.ASPCode
					drawValue TXT_ERROR_DESCRIPTION, errorObject.ASPDescription
				end if
				drawValue TXT_ERROR_SOURCE, errorObject.source
				.write("<br>")
				if debuggingVar <> "" then
					.write("<label>" & TXT_ERROR_HELPERVAR & ":</label><br>")
					.write("<textarea rows=3 cols=60 onclick=""this.select()"" readonly>" & debuggingVar & "</textarea>")
				end if
			'feedback only for logged in users because not logged in "could" hack...
			elseif lib.custom.userLoggedIn() then
				.write("<p>")
				.write(TXT_ERROR_USERFRIENDLY & "<br><br>")
				if feedbackURL <> "" then
					.write("<a href=""javascript:openCenteredModal('" & str.format(feedbackURL, array(server.URLEncode(executingFile), server.URLEncode(TXT_ERROR_PROCDESC))) & "', 300, 330, false)"">")
					.write("<strong>" & TXT_ERROR_CLICKTODESCRIBE & "</strong>")
					.write("</a><br><br>")
				end if
				.write(TXT_ERROR_TALKBACK)
				.write("</p>")
			'for not logged in users we show as less as possible.
			else
				.writeln("<p>" & TXT_ERROR_TALKBACK & "</p>")
			end if
			.write("</div></div>")
			.write("</div>")
		end with
		
		if response.buffer then str.write("</body></html>")
	end sub
	
	'******************************************************************************************************************
	'* drawValue 
	'******************************************************************************************************************
	public sub drawValue(name, value)
		if name <> "" then str.write("<label>" & name & ": </label>")
		str.writeln(str.defuseHTML(value) & "<br>")
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	sends an email with the error to the webadmins
	'' @DESCRIPTION:	mail is just send when notifyViaMail is set to true. the errorMail.html template is used
	''					for the mail
	'******************************************************************************************************************
	public sub sendMail()
		if not notifyViaMail then exit sub
		
		set template = new TextTemplate
		with template
			.fileName = consts.gabLibLocation & "class_errorHandler/errorMail.html"
			.addVariable "FILENAME", executingFile
			.addVariable "DURING", errorDuring
			.addVariable "DEBUGVAR", str.HTMLEncode(debuggingVar)
			.addVariable "ALTERNATIVETEXT", alternativeText
			if not errorObject is nothing then
				if typename(errorObject) = "IASPError" then
					.addVariable "CATEGORY", errorObject.category
					.addVariable "FILE", errorObject.file
					.addVariable "ERR_LINE", errorObject.line
					.addVariable "ERR_COLUMN", errorObject.column
					.addVariable "CODE", errorObject.ASPCode
					.addVariable "ASP_DESCRIPTION", errorObject.ASPDescription
				end if
				.addVariable "DESCRIPTION", errorObject.description
				.addVariable "NUMBER", errorObject.number
				.addVariable "SOURCE", errorObject.source
			end if
			.addVariable "USER_AGENT", request.serverVariables("HTTP_USER_AGENT")
			.addVariable "REFERRED_URL", request.serverVariables("HTTP_REFERER")
			.addVariable "REQUEST_METHOD", request.serverVariables("REQUEST_METHOD")
			.addVariable "REQUESTED_URL", request.serverVariables("SCRIPT_NAME")
			.addVariable "QUERYSTRING", lib.QS("")
			
			'add the form fields
			set block = new TextTemplateBlock
			if request.serverVariables("REQUEST_METHOD") = "POST" then
				for each f in lib.form
					block.addItem(array("NAME", f, "VALUE", lib.RF(f)))
				next
			end if
			.addVariable "FORMFIELDS", block
			
			set block = new TextTemplateBlock
			'add the session vars
			for each s in session.contents
				if isArray(session(s)) then
					val = str.arrayToString(session(s), "; ")
				elseif isObject(session(s)) then
					val = ""
				else
					val = session(s)
				end if
				block.addItem(array("NAME", s, "TYPE", typename(session(s)), "VALUE", val))
			next
			.addVariable "SESSIONS", block
		end with
		
		'now send the mail
		set mail = new CustomSendMail
		with mail
			.subject = template.getFirstLine()
			.body = template.getAllButFirstLine()
			admins = consts.webadmins
			for i = 0 to ubound(admins)
				email = admins(i, 1)
				if email <> "" then .addRecipient "TO", email, admins(i, 0)
			next
			.send()
		end with
		set mail = nothing
		set template = nothing
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Logs the error for example to database or a file. uses the lib.custom.logError()
	'' @DESCRIPTION:	if errors should be logged then the lib.custom.logError() should be implemented
	'******************************************************************************************************************
	public sub [log]()
		if logging then lib.custom.logError(me)
	end sub

end class
lib.registerClass("ErrorHandler")
%>