<!--#include virtual="/gab_Library/class_lib/lib.asp"-->
<!--#include virtual="/gab_Library/class_string/string.asp"-->
<!--#include virtual="/gab_Library/class_debuggingConsole/debuggingConsole.asp"-->
<!--#include virtual="/gab_LibraryConfig.asp"-->
<%
'**************************************************************************************************************
'* GABLIB Copyright (C) 2003 - This file is part of GABLIB - http://gablib.grafix.at
'**************************************************************************************************************
'* This program is free software; you can redistribute it and/or modify it under the terms of
'* the GNU General Publ. License as published by the Free Software Foundation; either version
'* 2 of the License, or (at your option) any later version. 
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		generatePage
'' @CREATOR:		Michal Gabrukieiwcz - gabru@gmx.at
'' @CREATEDON:		09.07.2003
'' @CDESCRIPTION:	This is the class for every page. Every page should be generated using an instance of this
''					class. Automatically created Header (header.asp), Footer (footer.asp), Database-connection, instance of library-class
''					for often used functions. It is real simple to add methods afterwards to whole project when
''					every page is created with this class.
'' @VERSION:		0.6

'**************************************************************************************************************
dim lib, cnn, DebugInfo 'its very important to dim these two objects. They are used in the whole page
dim consts : set consts = new constants
dim str : set str = new StringOperations

class generatePage

	private p_dev_warning_message	
	private p_access_denied_url		
	private startTime				
	private p_title					'The title of you page
	private p_loginRequired			'Login Required?
	private p_maintenanceSite		
	private p_accessID				
	private printCssLocation		
	
	public DBConnection				''[bool] do you need a Database-connection?
	public BgColor					''[string] Page-Background-color as Hex. e.g. #FFFFFF
	public showFooter				''[bool] you want to show the custom-footer
	public showHeader				''[bool] you want to show the custom-header
	public contentSub				''[string] The name of the content-procedure
	public bodyAttribute			''[string] Give the body an attribute e.g. onload=....
	public loadCss					''[bool] Load the stylesheets or not
	public frameSetter				''[bool] load the frameset if its not loaded? You need to modify the frameSetter.asp yourself.
	public HTTPHeader				''[bool] Do you want to write the HTTP Header?
	public debugMode				''[bool] Turn on the debugmode? class_debuggingConsole
	public drawBody					''[bool] Should the body be drawn? e.g. if you use frameset then you dont need body
	public devWarning				''[bool] display a warning if the page is on development-server.
	public pageEnterEffect			''[bool] enable page-Enter Effect. Only IE. Fade Effect!
	public pageEnterEffectDuration	''[string] e.g: 0.5 - The duration for the pageEffect.
	public accessRule				''[int] module-Section-ID which the user must have to view this page.
									''If the accessID returns 0 then a redirect to access_denied will happen.
									''If the user has access you can use the accessID-member to get the accessID for that module-Section
	public loadJavascript			''[bool] should the page.js be loaded within the page. it includes common javascripts
	public loadTooltips				''[bool] load the windows xp style balloon tooltips (loads javascript & stylesheet)
	public isModalDialog			''[bool] is this page a modal dialog. if yes then direct access to the page (without loading in modal) will be denied. default = false
	public DBConnectionParameter	''[string] lib.makeDBConnection works with one parameter. you can use this parameter to initialize the page with another parameter
									''for your database-connection. usefull when deriving
	
	'Konstruktor => set the default values
	private sub Class_Initialize()
		p_dev_warning_message	= "Warning! You are on the development server!"
		p_access_denied_url		= "/gab_Library/class_customLib/access_denied.asp"
		DBConnection			= false
		p_title					= consts.company_name
		p_loginRequired			= true
		BgColor					= consts.bgColor
		showFooter				= true
		showHeader				= true
		contentSub				= "main"
		bodyAttribute			= empty
		loadCss					= true
		loginRequired			= true
		frameSetter				= false
		HTTPHeader				= true
		debugMode				= false
		drawBody				= true
		devWarning				= true
		p_maintenanceSite		= "/gab_Library/class_customLib/maintenance.asp"
		pageEnterEffect			= false
		pageEnterEffectDuration	= "0.3"
		accessRule				= empty
		loadJavascript			= true
		loadTooltips			= false
		isModalDialog			= false
		DBConnectionParameter	= empty
		printCssLocation		= "/intranet_style/print.css"
	end sub
	
	public property get access_denied_url() ''[string] Returns the url of the page which should be shown if the user has no access to the page
		access_denied_url = p_access_denied_url
	end property
	
	public property get dev_warning_message() ''[string] Returns the message which should be showed when on development-server
		dev_warning_message = p_dev_warning_message
	end property
	
	public property get accessID() ''[int] Gets the Access ID, e.g. 1 - Read, ... Only returns a value if accessRule was used.
		accessID = p_accessID
	end property
	
	public property get maintenanceSite() ''[string] Returns the url of the maintenanceSite
		maintenanceSite = p_maintenanceSite
	end property
	
	public property get title() ''[string] Gets the page-Title
		title = p_title
	end property
	public property let title(value) ''[string] Sets the page-Title
		p_title = value
	end property
	
	public property get loginRequired() ''[bool] Returns if the user need to be logged in to watch this page.
		loginRequired = p_loginRequired
	end property
	public property let loginRequired(value) ''[bool] Sets if the page need a logged-in user to be showed or not?
		p_loginRequired = value
	end property
	
	'******************************************************************************************************************
	''@SDESCRIPTION:	Returns wheather the page was already sent to server or not.
	''@DESCRIPTION:		It should be the same as the isPostback from asp.net.
	''@RETURN:			[bool] posted back (true) or not (false)
	'******************************************************************************************************************
	public function isPostback()
		if request.form.count > 0 then
			isPostback = true
		else
			isPostback = false
		end if
	end function
	
	'************************************************************************************************************
	''@SDESCRIPTION:	Redirects you to the access denied page
	'************************************************************************************************************
	public sub showAccessDenied()
		response.redirect(access_denied_url)
		response.end
	end sub
	
	'************************************************************************************************************
	''@SDESCRIPTION:	Draws the page. Beware :)
	'************************************************************************************************************
	public function draw()
		'LIB-INSTANCE
		set lib = new library
		
		if (consts.maintenance_work and accessDuringMaintenance()) or not consts.maintenance_work then
			if p_loginRequired then
				if isUserLoggedIn() then 'logged in
					call page_main()
				else
					requestedUrl = request.serverVariables("URL")
					if not request.queryString = empty then
						requestedUrl = requestedUrl & "?" & replace(request.queryString, "&", "@")
					end if
					
					currentLoginName = trim(request.cookies("alu"))
					redirectPage = consts.domain & "emeaapps/autoLogin.asp?url=" & requestedUrl
					
					'user is not logged into local INTRANET. now we check if he is logged in to EMEA or not
					if currentLoginName = "" then
						response.redirect(consts.insideWyethGateway & redirectPage)
					else
						response.redirect(redirectPage)
					end if
				end if
			else
				call page_main()
			end if
		else
			server.transfer(p_maintenanceSite)
		end if
	end function
	
	'******************************************************************************************************************
	''@SDESCRIPTION:	Does someone has access to the site although maintenance work is going on.
	''@DESCRIPTION:		Thats an interface to accessDuringMaintenance-method of customLib. You have to implement this
	''					method in customLib yourself. It should return true/false if a user has access to the page
	''					although the maintenance work.
	''@RETURN:			[bool] access or not
	'******************************************************************************************************************
	public function accessDuringMaintenance()
		if lib.custom.accessDuringMaintenance() then
			accessDuringMaintenance = true
		else
			accessDuringMaintenance = false
		end if
	end function
	
	'******************************************************************************************************************
	''@SDESCRIPTION:	Checks if user is logged in or not
	''@DESCRIPTION:		This is just something like an interface. You need to implement a method in customLib called
	''					userLoggedIn which returns true when user is logged in and false if user is not logged in.
	''					Example: you could check in userLoggedIn-method if a special session-var is empty or not.
	''					If "loginRequired"-property is set to true and user is not logged in then noLogin.asp will be
	''					shown. Nologin is located in custom-class-Dir
	''@RETURN:			[bool] user logged in or not
	'******************************************************************************************************************
	public function isUserLoggedIn()
		if lib.custom.userLoggedIn() then
			isUserLoggedIn = true
		else
			isUserLoggedIn = false
		end if
	end function
	
	'******************************************************************************************************************
	''@SDESCRIPTION:	Adds a variable to the debug-informations. 
	''@DESCRIPTION:		Will be displayed if you use debugMode = true otherwise the variable wont be stored. 
	''					If you do this in the whole page for all important vars then you wont need to do it later again.
	''					Just debugMode enable/disable will also disable all debugvars.
	''@PARAM:			- description [string]: Description of the variable
	''@PARAM:			- var [variable]: The variable itself
	'******************************************************************************************************************
	public sub addDebugVar(description, byVal var)
		if debugMode then
			DebugInfo.Print description, var
		end if
	end sub
	
	'******************************************************************************************************************
	' loadFrameSetterJavascript 
	'******************************************************************************************************************
	private sub loadFrameSetterJavascript()
		%><!--#include virtual="/gab_Library/class_customLib/frameSetter.asp"--><%
	end sub
	
	'******************************************************************************************************************
	' init_debugger 
	'******************************************************************************************************************
	private sub init_debugger()
		if debugMode then
			Set DebugInfo = new debuggingConsole
			with DebugInfo
				.Enabled = true
				.AllVars = true
				.Show = "0,1,1,1,0,0,0,0,0,0,0" 'variables, Querystring and form are opened
			end with
		end if
	end sub
	
	'******************************************************************************************************************
	' display_debugger 
	'******************************************************************************************************************
	private sub display_debugger()
		if debugMode then
			DebugInfo.draw()
			Set DebugInfo = Nothing
		end if
	end sub
	
	'******************************************************************************************************************
	' PAGE MAIN 
	'******************************************************************************************************************
	private sub page_main()
		init_debugger()
		
		startTime = timer
		
		'DB-CONNECTION
		if DBConnection then
			lib.makeDbConnection(DBConnectionParameter)
			if debugMode then
				DebugInfo.grabDatabaseInfo(cnn)
			end if
		end if
		
		'Check Access
		if accessRule <> "" then
			if not DBConnection then
				str.writeln("IllestComplexException (ICE-2033) for Programmers.<br>Set DBConnection = true if checking access.")
				response.end
			end if
			p_accessID = lib.custom.hasAccess(accessRule)
			if accessID = ACCESS_NO then
				call lib.logAndForget("accessDenied", request.serverVariables("SCRIPT_NAME") & " (" & accessRule & ")")
				call showAccessDenied()
			end if
		end if
		
		'HTTPHEADER
		if HTTPHeader then
			response.buffer	= false
			response.expires = 0
		end if
		
		'CHECK IF MODAL DIALOG CALLED DIRECTLY
		call checkModalDialog()
		
		'FRAMESETTER
		if frameSetter then
			call loadFrameSetterJavascript()
		end if
		
		'HEADER
		call page_Header()
		
		'CONTENT
		if not contentSub = empty then
			execute("call " & contentSub)
		end if
		
		'CUSTOM FOOTER
		if showFooter then
			call page_custom_footer()
		end if
		
		'FOOTER
		call page_footer
		
		'DB-CONNECTION CLOSE
		if DBConnection then
			lib.closeDbConnection()
		end If
		
		'KILL objects
		set lib 	= nothing
		set consts 	= nothing
		set str 	= nothing
		
		display_debugger()
		if Response.Buffer then Response.Flush
	end sub
	
	'******************************************************************************************************************
	' HEADER FOR EVERYPAGE. 
	'******************************************************************************************************************
	private sub page_Header()
		if devWarning then
			if consts.isDevelopment() then
				str.writeln("<div class=notForPrint style=""text-align:center;background-color:red;font-size:10pt;color:white;position:absolute;top:0px;left:0px;z-index:10000;width:100%;filter:alpha(opacity=50);font-weight:bold;"">" & DEV_WARNING_MESSAGE & "</div>")
			end if
		end if
		
		if not bodyAttribute = empty then
			myAtt = " " & bodyAttribute
		else
			myAtt = empty
		end if
		with str
			.writeln("<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">")
			.writeln("<html>")
			.writeln("	<!--")
			.writeln("			PAGE GENERATED WITH GABLIB - ASP Class-Library by Michal Gabrukiewicz")
			.writeln("			COPYRIGHT " & ucase(consts.company_name))
			.writeln("			" & consts.domain)
			.writeln("	-->")
			.writeln("<head>")
			.writeln("<meta http-equiv=""content-type"" content=""text/html; charset=ISO-8859-1"">")
			.writeln("<meta http-equiv=""cache-control"" content=""no-cache"">")
			.writeln("<meta http-equiv=""pragma"" content=""no-cache"">")
			
			if pageEnterEffect then
				.writeln("	<meta http-equiv=""Page-Enter"" content=""blendtrans(duration=" & pageEnterEffectDuration & ")"">")
			end if
			
			.writeln("	<title>" & title & "</title>")
			.writeln("</head>")
			if loadJavascript then
				.writeln("<script language=""JavaScript"" src=""/gab_Library/class_page/page.js""></script>")
			end if
			if loadCss then
				.writeln("<link rel=""stylesheet"" type=""text/css"" href=""" & consts.page_stylesheet & """>")
				.writeln("<link rel=""stylesheet"" type=""text/css"" media=""print"" href=""" & printCssLocation & """>")
			end if
			if drawBody then
				.writeln("<body bgcolor=" & BgColor & " " & myAtt & ">")
			end if
			
			if loadTooltips then
				.writeln("<script language=""JavaScript"" src=""/gab_Library/class_tooltip/tooltip.js""></script>")
				.writeln("<link rel=""stylesheet"" type=""text/css"" href=""/gab_Library/class_tooltip/tooltip.css"">")
			end if
		end with
		if showHeader then
			%><!--#include file="custom_Header.asp"--><%
		end if
		
		if isModalDialog then
			str.writeln("<base target=""_self"">")
		end if
	end sub
	
	'******************************************************************************************************************
	' CUSTOM FOOTER INCL COMPANY SIGNATURE ETC. 
	'******************************************************************************************************************
	private sub page_custom_footer()
		showURL = left(consts.domain, (len(consts.domain)-1))
		pageLoadTime = formatnumber((timer - startTime),3)
		%><!--#include file="custom_Footer.asp"--><%
	end sub
	
	'******************************************************************************************************************
	' MAIN FOOTER. JUST TO CLOSE HTML PAGE 
	'******************************************************************************************************************
	private sub page_footer()
		if drawBody then
			response.write "</body>"
		end if
		response.write "</HTML>"
	end sub
	
	'******************************************************************************************************************
	' checkModalDialog 
	'******************************************************************************************************************
	private sub checkModalDialog()
		if isModalDialog then %>
			<script language="javascript">
				if(!window.dialogArguments) {
					location.href = "<%= access_denied_url %>";
				}
			</script>
	<%	end if
	end sub
	
end class
%>