<!--#include file="../class_lib/lib.asp"-->
<!--#include file="../class_debuggingConsole/debuggingConsole.asp"-->
<!--#include virtual="/gab_libraryConfig/_generatepage.asp"-->
<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003 
'* License refer to license.txt   
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		generatePage
'' @CREATOR:		Michal Gabrukiewicz - gabru @ grafix.at
'' @CREATEDON:		09.07.2003
'' @CDESCRIPTION:	This class represents a webpage for pages in your applications. Some people like to call
''					it a masterpage. It handles all the common functionality which a page should handle and
''					do for you like loading all needed libraries e.g. asp, javascript, css, etc., 
''					drawing headers & footers, handle login, print all neccessary HTML and much more.
''					This class should be used for every page. Thats the only way you can stay flexible for
''					feature changes. Its also recommended to derive from it when creating new applications
''					IMPORTANT: whenever this class is included it must be the first includ! so always the
''					first line.
'' @POSTFIX:		page
'' @VERSION:		1.0

'**************************************************************************************************************

class GeneratePage

	private p_dev_warning_message, p_access_denied_url
	private startTime, p_title, p_loginRequired, p_maintenanceSite
	private printCssLocation, standardCssLocation
	private standardModCssLocation, debugConsole
	private p_maintenanceWarningMSG, loadedSources
	
	public DBConnection				''[bool] do you need a Database-connection?
	public backgroundColor			''[string] Page-Background-color as Hex. e.g. #FFFFFF
	public showFooter				''[bool] you want to show the custom-footer
	public showHeader				''[bool] you want to show the custom-header
	public bodyAttribute			''[string] Give the body an attribute e.g. onload=....
	public loadCss					''[bool] Load the stylesheets or not
	public framesetter				''[bool] load the page within the frameset if its not in a frameset? default = false
	public framesetURL				''[string] if framesetter is used then this is the URL of the frameset which is used to display the page
									''ideally it should be changed when derriving because each application uses its own frameset (if it uses frames).
									''default: consts.domain (place a {0} which will be replaced by the url to load in the frameset)
	public isFrameset				''[bool] is this instance a frameset? this is important because e.g. body-tag is not written, etc. default = false
	public HTTPHeader				''[bool] OBSOLETE! Do you want to write the HTTP Header? 
	public buffering				''[bool] should the output be buffered? if false then e.g. response.redirect does not work. default = true
									''if turning this off then the all ASP errors which are thrown after some response has been made wont be handled.
									''this is only important if using the 500-100.asp errorpage
	public debugMode				''[bool] Turn on the debugmode? class_debuggingConsole
	public drawBody					''[bool] Should the body be drawn? e.g. if you use frameset then you dont need body
	public devWarning				''[bool] display a warning if the page is on development-server.
	public pageEnterEffect			''[bool] enable page-Enter Effect. Only IE. Fade Effect!
	public pageEnterEffectDuration	''[string] e.g: 0.5 - The duration for the pageEffect.
	public loadJavascript			''[bool] should the page.js be loaded within the page. it includes common javascripts
	public loadTooltips				''[bool] load the windows xp style balloon tooltips (loads javascript & stylesheet)
	public isModalDialog			''[bool] is this page a modal dialog. if yes then direct access to the page (without loading in modal) will be denied. default = false
	public DBConnectionParameter	''[string] lib.makeDBConnection works with one parameter. you can use this parameter to initialize the page with another parameter
									''for your database-connection. usefull when deriving
	public maintenanceWork			''[bool] maintenance-work on or of. value taken from constants
	public showOnMaintenanceWork	''[bool] does this page should be rendered on maintenance-work? default=false. 
	public enableModalStyles		''[bool] this will only work in modal dialogs - disables default styles and activates modal.css. default=false
	public isXML					''[bool] indicates if the page should be rendered as an XML-file. default = false
	public metaDescription			''[string] description of the page. will be used within the description-meta-tag
	public metaKeywords				''[string] description for the page. will be used within the keywords-meta-tag
	public forceStandardApp			''[bool] if on then it forces that the appereance of a standard application will be loaded. default = false
									''this is useful when creating an application which should use the appereance of a standard application and not the
									''defined appereance of the customized gabLibrary
	public onlyWebDev				''[bool] is this page accessible only for webdevs? use this if page is sensitive. default = false
	public loadPrototypeJS			''[bool] loads the protype.js library (http://www.prototypejs.org). default = false
									''for the documentation of the library refer to their website. it eases the JS development and allows easy
									''work with AJAX technology
	public plain					''[bool] should a plain page be generated? that means that just the stuff from main()-method will be rendered.
									''no html-, body-, etc tags are rendered out. useful for Ajax requests which use the HTML of a given page. default = false
	public contentType				''[string] the contenttype of the page. by default its the default of response.contentType (text/html). 
									''e.g. When setting isXML to true contenttype is set to text/xml automatically
	public charset					''[string] default = utf-8
	
	public property get access_denied_url() ''[string] Returns the url of the page which should be shown if the user has no access to the page
		access_denied_url = p_access_denied_url
	end property
	
	public property get maintenanceWarningMSG() ''[string] Returns the message which should be showed when access during maintenance-work.
		maintenanceWarningMSG = p_maintenanceWarningMSG
	end property
	
	public property get dev_warning_message() ''[string] Returns the message which should be showed when on development-server
		dev_warning_message = p_dev_warning_message
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
	
	public property get loginRequired() ''[bool] Returns if the user need to be logged in to view this page.
		loginRequired = p_loginRequired
	end property
	
	public property let loginRequired(value) ''[bool] Sets if the page needs a logged-in user to be showed or not?
		p_loginRequired = value
	end property
	
	'************************************************************************************************************
	' constructor 
	'************************************************************************************************************
	public sub class_initialize()
		set lib.page			= me
		startTime 				= timer()
		p_dev_warning_message	= "Warning! You are on the development server!"
		p_maintenanceWarningMSG	= "Maintenance work on! Access granted!"
		p_access_denied_url		= lib.init(GL_GP_ACCESSDENIEDURL, consts.gabLibLocation & "class_page/accessDenied.asp")
		standardCssLocation		= lib.init(GL_GP_STANDARDCSSLOCATION, consts.STDAPP("std.css"))
		standardModCssLocation	= lib.init(GL_GP_MODALCSSLOCATION, consts.STDAPP("modal.css"))
		printCssLocation		= lib.init(GL_GP_PRINTCSSLOCATION, consts.STDAPP("print.css"))
		backgroundColor			= "#ffffff"
		p_title					= consts.company_name
		p_maintenanceSite		= lib.init(GL_GP_MAINTENANCEURL, consts.gabLibLocation & "class_page/maintenance.asp")
		DBConnection			= lib.init(GL_GP_DBCONNECTION, true)
		p_loginRequired			= lib.init(GL_GP_LOGINREQUIRED, true)
		showFooter				= true
		showHeader				= true
		bodyAttribute			= empty
		loadCss					= true
		loginRequired			= lib.init(GL_GP_LOGINREQUIRED, true)
		frameSetter				= false
		HTTPHeader				= true
		debugMode				= false
		drawBody				= true
		devWarning				= true
		pageEnterEffect			= false
		pageEnterEffectDuration	= "0.3"
		loadJavascript			= true
		loadTooltips			= false
		isModalDialog			= false
		DBConnectionParameter	= empty
		maintenanceWork			= consts.maintenance_work
		showOnMaintenanceWork	= false
		isFrameset				= false
		enableModalStyles		= true
		set loadedSources		= server.createObject("scripting.dictionary")
		isXML					= false
		framesetURL				= consts.domain & "?url={0}"
		buffering				= true
		forceStandardApp		= false
		onlyWebDev				= false
		loadPrototypeJS			= false
		plain					= false
		charset					= "utf-8"
	end sub
	
	'************************************************************************************************************
	' destructor 
	'************************************************************************************************************
	private sub class_terminate()
		set loadedSources = nothing
	end sub
	
	'************************************************************************************************************
	'' @SDESCRIPTION:	Draws the page. Beware :)
	'' @DESCRIPTION:	If you derive from this class you can call this method in your new class or you call
	''					every method which is inside. this allows you to add own headers, footers, etc.
	'************************************************************************************************************
	public function draw()
		checkMaintenanceWork()
		checkLogin()
		initDebugger()
		openDatabaseConnection()
		checkAccess()
		setHTTPHeader()
		checkModalDialog()
		checkFramesetter()
		drawPageHeader()
		drawCustomHeader()
		drawDevWarning()
		drawMaintenanceWarning()
		drawContent()
		drawCustomFooter()
		drawPageFooter()
		closeDatabaseConnection()
		drawDebugger()
		destruct()
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Returns wheather the page was already sent to server or not.
	'' @DESCRIPTION:	It should be the same as the isPostback from asp.net.
	'' @RETURN:			[bool] posted back (true) or not (false)
	'******************************************************************************************************************
	public function isPostback()
		isPostback = (lib.form.count > 0)
	end function
	
	'************************************************************************************************************
	'' @SDESCRIPTION:	Redirects you to the access denied page
	'************************************************************************************************************
	public sub showAccessDenied()
		response.redirect(access_denied_url)
		response.end()
	end sub
	
	'************************************************************************************************************
	'* checkMaintenanceWork
	'************************************************************************************************************
	public sub checkMaintenanceWork()
		if not hasAccessOnMaintenance then server.transfer(p_maintenanceSite)
	end sub
	
	'************************************************************************************************************
	'' @SDESCRIPTION:	is the site accessible due to all maintenance rules?
	'' @RETURN:			[bool] true if can be shown...
	'************************************************************************************************************
	public function hasAccessOnMaintenance()
		hasAccessOnMaintenance = true
		if not showOnMaintenanceWork then
			if maintenanceWork and not lib.custom.accessDuringMaintenance() then hasAccessOnMaintenance = false
		end if
	end function
	
	'************************************************************************************************************
	'* checkLogin
	'************************************************************************************************************
	public sub checkLogin()
		if (loginRequired and not isUserLoggedIn()) then
			url = server.urlEncode(request.serverVariables("URL") & "?" & lib.QS(""))
			response.redirect(str.format(consts.loginPageUrl, array(url, lib.iif(isModalDialog, "1", ""))))
		end if
	end sub
	
	'******************************************************************************************************************
	'* checkFramesetter 
	'******************************************************************************************************************
	public sub checkFramesetter()
		if framesetter and not isXML and not isFrameset and not plain then
			rURL = str.format(framesetURL, array(server.URLEncode(request.serverVariables("URL") & "?" & lib.QS(""))))
			lib.execJS("if(self==top || !parent)top.location.href = '" & rURL & "';")
		end if
	end sub
	
	'******************************************************************************************************************
	'* openDatabaseConnection 
	'******************************************************************************************************************
	public sub openDatabaseConnection()
		if DBConnection then
			lib.custom.establishDatabaseConnection DBConnectionParameter
			if debugMode then
				if lib.custom.isWebadmin() then debugConsole.grabDatabaseInfo(lib.databaseConnection)
			end if
		end if
	end sub
	
	'******************************************************************************************************************
	'* checkAccess 
	'******************************************************************************************************************
	public sub checkAccess()
		if onlyWebDev then 
			if not lib.custom.isWebadmin() then showAccessDenied()
		end if
	end sub
	
	'******************************************************************************************************************
	'* closeDatabaseConnection 
	'******************************************************************************************************************
	public sub closeDatabaseConnection()
		if DBConnection then lib.databaseConnection.close()
	end sub
	
	'******************************************************************************************************************
	'* setHTTPHeader 
	'******************************************************************************************************************
	public sub setHTTPHeader()
		with response
			.charset = charset
			.codePage = 65001
			if isXML then
				.contentType = "text/xml"
			elseif contentType <> "" then
				.contentType = contentType
			end if
			.expires = 0
			.buffer = buffering
		end with
	end sub
	
	'******************************************************************************************************************
	'* drawContent 
	'******************************************************************************************************************
	public sub drawContent()
		main()
	end sub
	
	'******************************************************************************************************************
	'* drawPageFooter 
	'******************************************************************************************************************
	public sub drawPageFooter()
		if not isXML and not plain then
			if drawBody and not isFrameset then str.write("</body>")
			str.write("</html>")
		end if
	end sub
	
	'******************************************************************************************************************
	'* drawCustomHeader 
	'******************************************************************************************************************
	public sub drawCustomHeader()
		if showHeader and not isXML and not isModalDialog and not isFrameset and not plain then
			%><!--#include virtual="/gab_LibraryConfig/pageHeader.asp"--><%
		end if
	end sub
	
	'******************************************************************************************************************
	'* drawCustomFooter 
	'******************************************************************************************************************
	public sub drawCustomFooter()
		if showFooter and not isXML and not isModalDialog and not isFrameset and not plain then
			%><!--#include virtual="/gab_LibraryConfig/pageFooter.asp"--><%
			if lib.custom.isWebadmin() then
				with str
					.write("<div id=developerToolbar class=""notForPrint"">")
					.write("<span title=""Querystring: " & str.NAIfEmpty(str.HTMLEncode(request.queryString)) & """>" & request.serverVariables("SCRIPT_NAME") & "</span>: ")
					if consts.isDevelopment() then
						liveURL = consts.liveServerName & request.serverVariables("SCRIPT_NAME") & "?" &  request.queryString
						.write("<a href=""http://" & liveURL & """ target=""_blank"">load@live</a> | ")
					end if
					.write("<a href=""javascript:viewSource();"">Source</a> | ")
					.write("<a href=""javascript:location.reload();"">Reload</a>")
					.write("</div>")
				end with
			end if
		end if
	end sub
	
	'******************************************************************************************************************
	'* destruct 
	'******************************************************************************************************************
	public sub destruct()
		set lib = nothing
		set str = nothing
		set consts = nothing
		'IE7 bug!
		'if response.buffer then response.flush()
	end sub
	
	'******************************************************************************************************************
	'* drawMaintenanceWarning 
	'******************************************************************************************************************
	public sub drawMaintenanceWarning()
		if maintenanceWork and not showOnMaintenanceWork and not isFrameset and not isXML and not plain then
			if lib.custom.accessDuringMaintenance() then drawInformationBanner maintenanceWarningMSG, 50, "green"
		end if
	end sub
	
	'******************************************************************************************************************
	'* drawDevWarning 
	'******************************************************************************************************************
	public sub drawDevWarning()
		if devWarning and not isFrameset and not isXML and not plain then
			if consts.isDevelopment() then drawInformationBanner dev_warning_message, 0, "red"
		end if
	end sub
	
	'******************************************************************************************************************
	'* drawInformationBanner 
	'******************************************************************************************************************
	private sub drawInformationBanner(msg, leftPercent, backColor)
		str.writeln("<div class=notForPrint style=""text-align:center;background-color:" & backColor & ";font-size:8pt;color:white;position:absolute;top:0px;left:" & leftPercent & "%;z-index:10000;width:50%;filter:alpha(opacity=50);"">" & msg & "</div>")
	end sub
	
	'******************************************************************************************************************
	'* drawPageHeader 
	'******************************************************************************************************************
	public sub drawPageHeader()
		if isXML or plain then exit sub
		
		'if we force the appearence we need to change the CSS-locations
		if forceStandardApp then
			standardCssLocation = consts.STDAPP("std.css")
			printCssLocation = consts.STDAPP("print.css")
			standardModCssLocation = consts.STDAPP("modal.css")
		end if
		
		with str
			.writeln("<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">")
			.writeln("<html>")
			.writeln("	<!--")
			.writeln("		generated with 'gabLibrary' a Web Application Framework by Michal Gabrukiewicz")
			.writeln("		CONTENT COPYRIGHT BY " & ucase(consts.company_name))
			.writeln("	-->")
			.writeln("<head>")
			.writeln("	<meta http-equiv=""content-type"" content=""text/html; charset=ISO-8859-1""/>")
			.writeln("	<meta http-equiv=""expires"" content=""0""/>")
			.writeln("	<meta http-equiv=""cache-control"" content=""no-cache""/>")
			.writeln("	<meta http-equiv=""pragma"" content=""no-cache""/>")
			.writeln("	<link rel=""shortcut icon"" href=""/favicon.ico""/>")
			if metaDescription <> "" then writeMetaTag "description", metaDescription
			if metaKeywords <> "" then writeMetaTag "keywords", metaKeywords
			
			if pageEnterEffect then .writeln("	<meta http-equiv=""Page-Enter"" content=""blendtrans(duration=" & pageEnterEffectDuration & ")""/>")
			
			.writeln("	<title>" & title & "</title>")
			
			loadStyles()
			
			if isModalDialog then .writeln("<base target=""_self"">")
			.writeln("</head>")
			
			if drawBody and not isFrameset then
				bAttr = bodyAttribute
				if isModalDialog then bAttr = "onkeydown=""handleModalEscape()"" " & bodyAttribute
				.writeln("<body bgcolor=""" & backgroundColor & """" & lib.iif(bAttr <> empty, " " & bAttr, empty) & ">")
			end if
			
			if loadJavascript then loadJSCore()
			if loadPrototypeJS then
				loadJavascriptFile(consts.gabLibLocation & "class_page/prototype.js")
				if not isFrameset then
					str.writeln("<img id=""ajaxLoader"" class=""notForPrint"" style=""display:none;position:absolute;right:10;bottom:10;"" src=""" & consts.stdapp("images/loading_ajax.gif") & """>")
					lib.execJS("Ajax.Responders.register({ onCreate: function() { $('ajaxLoader').show(); }, onComplete: function() { $('ajaxLoader').hide(); } }); ")
				end if
			end if
			
			if loadTooltips then
				loadJavascriptFile(consts.gabLibLocation & "class_tooltip/tooltip.js")
				loadStylesheetFile consts.gabLibLocation & "class_tooltip/tooltip.css", empty
			end if
			
		end with
	end sub
	
	'******************************************************************************************************************
	'* loadCSS 
	'******************************************************************************************************************
	private sub loadStyles()
		if not loadCss then exit sub
		loadStylesheetFile standardCssLocation, empty
		loadStylesheetFile printCssLocation, "print"
		if isModalDialog and enableModalStyles then loadStylesheetFile standardModCssLocation, empty
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Loads a the javascript core which is used within each page
	'******************************************************************************************************************
	public sub loadJSCore()
		loadJavascriptFile(consts.gabLibLocation & "class_page/page.js")
	end sub
	
	'******************************************************************************************************************
	'* write 
	'******************************************************************************************************************
	private sub writeMetaTag(metaName, metaContent)
		str.writeln("<meta name=""" & metaName & """ content=""" & str.HTMLEncode(metaContent) & """/>")
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Loads a specified javscript-file
	'' @DESCRIPTION:	it will load the file only if it has not been already loaded on the page before.
	''					so you can load the file and dont have to worry if it will be loaded more than once.
	''					differentiation between the files is the filename (case-sensitive!).
	''					- has no effect on plain pages!
	'' @PARAM:			url [string]: url of your javascript-file
	'******************************************************************************************************************
	public sub loadJavascriptFile(url)
		if isXML or plain then exit sub
		sourceID = "JS" & url
		if not loadedSources.exists(sourceID) then
			str.writeln("<script type=""text/javascript"" language=""javascript"" src=""" & url & """></script>")
			loadedSources.add sourceID, empty
		end if
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Loads a specified stylesheet-file
	'' @DESCRIPTION:	- has no effect on plain pages!
	'' @PARAM:			url [string]: url of your stylesheet
	'' @PARAM:			media [string]: what media is this stylesheet for. screen, etc. leave it blank if for every media
	'******************************************************************************************************************
	public sub loadStylesheetFile(url, media)
		if isXML or plain then exit sub
		sourceID = "CSS" & url & media
		if not loadedSources.exists(sourceID) then
			str.writeln("<link rel=""stylesheet"" type=""text/css""" & lib.iif(media <> "", " media=""" & media & """", empty) & " href=""" & url & """>")
			loadedSources.add sourceID, empty
		end if
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Checks if user is logged in or not
	'' @DESCRIPTION:	This is just something like an interface. You need to implement a method in customLib called
	''					userLoggedIn which returns true when user is logged in and false if user is not logged in.
	''					Example: you could check in userLoggedIn-method if a special session-var is empty or not.
	''					If "loginRequired"-property is set to true and user is not logged in then noLogin.asp will be
	''					shown. Nologin is located in custom-class-Dir
	'' @RETURN:			[bool] user logged in or not
	'******************************************************************************************************************
	public function isUserLoggedIn()
		isUserLoggedIn = lib.custom.userLoggedIn()
	end function
	
	'******************************************************************************************************************
	'* initDebugger 
	'******************************************************************************************************************
	public sub initDebugger()
		if debugMode then
			if not lib.custom.isWebadmin() then exit sub
			set debugConsole = new DebuggingConsole
			with debugConsole
				.enabled = true
				.allVars = true
				.show = "0,1,1,1,0,0,0,0,0,0,0,0" 'variables, Querystring and form are opened
			end with
		end if
	end sub
	
	'******************************************************************************************************************
	'* drawDebugger 
	'******************************************************************************************************************
	public sub drawDebugger()
		if debugMode and not isXML then
			if lib.custom.isWebAdmin() then
				debugConsole.print "Number of Database accesses", lib.numberOfDBAccess
				debugConsole.draw()
			end if
			set debugConsole = Nothing
		end if
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Adds a variable to the debug-informations. 
	'' @DESCRIPTION:	Will be displayed if you use debugMode = true otherwise the variable wont be stored. 
	''					If you do this in the whole page for all important vars then you wont need to do it later again.
	''					Just debugMode enable/disable will also disable all debugvars.
	'' @PARAM:			description [string]: Description of the variable
	'' @PARAM:			var [variable]: The variable itself
	'******************************************************************************************************************
	public sub addDebugVar(description, byVal var)
		if debugMode then debugConsole.print description, var
	end sub
	
	'******************************************************************************************************************
	'* checkModalDialog 
	'******************************************************************************************************************
	public sub checkModalDialog()
		if not isXML and isModalDialog then
			lib.execJS("if(!window.dialogArguments) location.href = '" & access_denied_url & "';")
		end if
	end sub

end class
lib.registerClass("GeneratePage")
%>