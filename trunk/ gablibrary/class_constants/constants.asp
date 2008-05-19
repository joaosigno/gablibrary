<!--#include virtual="/gab_LibraryConfig.asp"-->
<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		Constants
'' @CREATOR:		Michal Gabrukieiwcz - gabru @ grafix.at
'' @CREATEDON:		02.12.2003
'' @CDESCRIPTION:	This class stores all the configs needed for the project. They are all readonly.
''					Using this class we make our code more "nicer" to read. It is loaded within the 
''					page-class with the name called "consts". All vars are loaded from the config file.
''					if system is running on the livesystem and no live var is available then the var from
''					from the dev is taken. are both no available then the default is taken.
'' @STATICNAME:		consts
'' @VERSION:		0.2

'**************************************************************************************************************
class Constants

	private p_isDevelopment, p_company_name
	private p_domain, p_loginPageUrl, p_errorLogging
	private p_hostIP, p_smtp_server, p_logo, p_logo_Small
	private p_automated_bot_email, p_automated_bot_name
	private p_email_send_on_error, p_maintenance_work
	private p_logs_path, p_liveServerName, p_NotAvailable, p_userFiles, p_liveWebservicesName
	private p_webadmins, p_allowed_on_maintenance, p_onlyLiveServer, p_UTF8, p_webservicesHost
	
	'******************************************************************************************************************
	'* constructor
	'******************************************************************************************************************
	public sub class_Initialize()
		'the liveserver stuff needs to be initialized first
		p_liveServerName 			= initVar(GL_CONST_LIVERSERVERNAME, "localhost")
		p_liveWebservicesName		= initVar(GL_CONST_LIVEWEBSERVICESHOST, "localhost")
		p_onlyLiveServer			= initVar(GL_CONST_ONLYLIVESERVER, false)
		
		'the rest of the vars
		p_allowed_on_maintenance 	= init(GL_CONST_ALLOWEDONMAINTENANCE, GL_CONST_ALLOWEDONMAINTENANCE_DEV, array())
		p_webadmins 				= init(GL_CONST_WEBADMINS, GL_CONST_WEBADMINS_DEV, array())
		p_domain 					= init(GL_CONST_DOMAIN, GL_CONST_DOMAIN_DEV, "http://localhost/")
		p_company_name 				= init(GL_CONST_COMPANYNAME, GL_CONST_COMPANYNAME_DEV, "My company")
		p_hostIP 					= init(GL_CONST_HOSTIP, GL_CONST_HOSTIP_DEV, "127.0.0.1")
		p_loginPageUrl 				= init(GL_CONST_LOGINURL, GL_CONST_LOGINURL_DEV, "http://localhost/login.asp?url={0}")
		p_smtp_server 				= init(GL_CONST_SMTP, GL_CONST_SMTP_DEV, "127.0.0.1")
		p_userFiles 				= init(GL_CONST_USERFILES, GL_CONST_USERFILES_DEV, "/userFiles/")
		p_email_send_on_error 		= init(GL_CONST_EMAILONERROR, GL_CONST_EMAILONERROR_DEV, true)
		p_maintenance_work 			= init(GL_CONST_MAINTENANCE, GL_CONST_MAINTENANCE_DEV, false)
		p_errorLogging 				= init(GL_CONST_ERRORLOGGING, GL_CONST_ERRORLOGGING_DEV, true)
		p_logo 						= init(GL_CONST_LOGO, GL_CONST_LOGO_DEV, "/images/logo.gif")
		p_logo_small 				= init(GL_CONST_LOGOSMALL, GL_CONST_LOGOSMALL_DEV, "/images/logoSmall.gif")
		p_automated_bot_email 		= init(GL_CONST_BOTMAIL, GL_CONST_BOTMAIL_DEV, "bot@localhost")
		p_automated_bot_name 		= init(GL_CONST_BOTNAME, GL_CONST_BOTNAME_DEV, "Email bot")
		p_logs_path 				= init(GL_CONST_LOGSPATH, GL_CONST_LOGSPATH_DEV, "/log/")
		p_NotAvailable 				= init(GL_CONST_NOTAVAILABLE, GL_CONST_NOTAVAILABLE_DEV, "N/A")
		p_UTF8 						= init(GL_CONST_UTF8, GL_CONST_UTF8_DEV, true)
		p_webservicesHost			= init(GL_CONST_WEBSERVICESHOST, GL_CONST_WEBSERVICESHOST_DEV, "http://localhost/")
	end sub
	
	public property get env ''[string] gets the environment. 'dev' or 'live'
		env = "dev"
		if not isDevelopment() then env = "live"
	end property
	
	public property get webservicesHost ''[bool] host where the webservices are running. OBSOLETE! only for backwards compatibility with old webservice component
		webservicesHost = p_webservicesHost
	end property
	
	public property get UTF8 ''[bool] is the charset set to UTF-8 for the whole gablibrary? if false then server default is taken. default = true.
		UTF8 = p_UTF8
	end property
	
	public property get onlyLiveServer ''[bool] is the system configured only with a live system (there is no live/dev scenario). default = false
		onlyLiveServer = p_onlyLiveServer
	end property
	
	public property get errorLogging ''[bool] should error logging be done?
		errorLogging = p_errorLogging
	end property
	
	public property get gabLibLocation ''[string] Returns the virtual path of the gabLibrary and ends with a slash (example: /gab_Library/)
		gabLibLocation = "/gab_Library/"
	end property
	
	public property get gabLibConfigLocation ''[string] Returns the virtual path of the gabLibrary and ends with a slash (example: /gab_Library/)
		gabLibConfigLocation = "/gab_LibraryConfig/"
	end property
	
	public property get liveServerName ''[string] Returns the server-name of the live-webServer
		liveServerName = p_liveServerName
	end property
	
	public property get logs_path ''[string] Returns the path to the logs
		logs_path = p_logs_path
	end property
	
	public property get company_name ''[string] Returns the Companies name
		company_name = p_company_name
	end property
	
	public property get userFiles ''[string] gets the virtual path to the user files. user files are file uploaded by the user or modified by the user
		userFiles = p_userFiles
	end property
	
	public property get NA ''[string] Returns a string which will indicates "no data available"
		NA = p_NotAvailable
	end property
	
	public property get domain ''[string] 'Thats the URL of the machine where all these files are located. e.g. http://inside.me.com
		domain = p_domain
	end property
	
	public property get loginPageUrl ''[string] gets the URL of the loginpage. {0} placeholder will be replaced by the requested URL
		loginPageUrl = p_loginPageUrl
	end property
	
	public property get hostIP ''[string] Returns the IP-address of the host
		hostIP = p_hostIP
	end property
	
	public property get smtp_server ''[string] Returns the SMTP-server
		smtp_server = p_smtp_server
	end property
	
	public property get logo ''[string] Returns the Path to the company-Logo.
		logo = p_logo
	end property
	
	public property get logoSmall ''[string] Returns the Path to the SMALL company-Logo.
		logoSmall = p_logo_Small
	end property
	
	public property get automated_bot_email ''[string] Returns the Name of the Automated Bot. e.g. used for email sending as sender.
		automated_bot_email = p_automated_bot_email
	end property
	
	public property get automated_bot_name ''[string] Returns the Email of the Automated Bot. e.g. used for email sending as sender.
		automated_bot_name = p_automated_bot_name
	end property
	
	public property get email_send_on_error ''[bool] Returns if there will be email sending on errors.
		email_send_on_error = p_email_send_on_error
	end property
	
	public property get maintenance_work ''[bool] Returns if maintanance work is enabled or disabled
		maintenance_work = p_maintenance_work
	end property
	
	public property get webadmins ''[array] Retuns a 2 dimensional array with the Webadmins. 1st column is the name 2nd column is the email of the webadmin, 3rd column the userid.
		webadmins = p_webadmins
	end property
	
	public property get allowed_on_maintenance ''[array] Retuns an array with the IP-addresses which are allowed to view the page during maintanance work.
		allowed_on_maintenance = p_allowed_on_maintenance
	end property
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	gets a full virtual path for a file from the standard application.
	'' @PARAM:			filename [string]: the requested file name you wish from the standard application e.g. std.css
	''					provide empty to get just the path to the standard application
	'' @RETURN:			[string] a full virtual path to the requested file. e.g. /gab_Library/STANDARDAPP/std.css
	'******************************************************************************************************************
	public function STDAPP(filename)
		STDAPP = gabLibLocation & "STANDARDAPP/" & filename
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Tells if site is running on development server or not
	'' @RETURN:			[bool] if the site currently runs on developmentServer or not
	'******************************************************************************************************************
	public function isDevelopment()
		if isEmpty(p_isDevelopment) then
			if onlyLiveServer then
				p_isDevelopment = false
			else
				p_isDevelopment = (not hostIs(liveServerName) and not hostIs(p_liveWebservicesName))
			end if
		end if
		isDevelopment = p_isDevelopment
	end function
	
	'******************************************************************************************************************
	'* hostIs
	'******************************************************************************************************************
	private function hostIs(hostname)
		hostIs = (ucase(request.serverVariables("SERVER_NAME")) = ucase(hostname))
	end function
	
	'******************************************************************************************************************
	'* initialisation of the vars.
	'******************************************************************************************************************
	private function init(liveVar, devVar, default)
		if isDevelopment() then
			init = initvar(devVar, initvar(liveVar, default))
		else
			init = initvar(liveVar, initvar(devVar, default))
		end if
	end function
	
	'******************************************************************************************************************
	'* initvar
	'******************************************************************************************************************
	private function initvar(var, default)
		initvar = var
		if isEmpty(var) then initvar = default
	end function

end class
lib.registerClass("Constants")
%>