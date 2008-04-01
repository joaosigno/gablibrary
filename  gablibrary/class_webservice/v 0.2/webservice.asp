<!--#include virtual="/gab_Library/class_page/generatePage.asp"-->
<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		Webservice
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		23.02.2006 13:39
'' @CDESCRIPTION:	Gives you the possibility to create a webservice using common ASP
'' 					In general it just lets you create a XML file.
''					provide a main()-method and use the XMLDOM-property to set all your data, then
''					use the generate()-method to generate the webservice.
''					You will always get a valid XML-file, even if there are any errors in the webservice.
''					On an error you get an ERROR-node with the error. If the client don't adds a root-node
''					then also an error will be returned because of no root-node
'' @REQUIRES:		-
'' @VERSION:		0.2

'**************************************************************************************************************
class Webservice

	'private members
	private myBase
	
	'public members
	public XMLDOM
	public errorHandling	''[bool] disable/enable error handling. good for development. default = true
	
	public property let DBConnectionParameter(val)
		myBase.DBConnectionParameter = val
	end property
	
	public property let DBConnection(val)
		myBase.DBConnection = val
	end property
	
	public property get ERRCODE_CUSTOM ''[int] gets the error-code for custom-errors
		ERRCODE_CUSTOM = 1
	end property
	
	public property get ERRCODE_WRONGHOST ''[int] gets the error-code if webservice called not on webservice host
		ERRCODE_WRONGHOST = 2
	end property
	
	public property get ERRCODE_UNEXPECTED ''[int] gets error-code if any unexpected ASP error happens
		ERRCODE_UNEXPECTED = 4
	end property
	
	public property get ERRCODE_INVALIDXML ''[int] gets error-code if the XML is not valid. e.g. no root-node
		ERRCODE_INVALIDXML = 8
	end property
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		set myBase = new GeneratePage
		set XMLDOM = server.createObject("Microsoft.XMLDOM")
		lib.disableErrorHandling = true
		myBase.isXML = true
		errorHandling = true
	end sub
	
	'**********************************************************************************************************
	'* destructor 
	'**********************************************************************************************************
	public sub class_terminate()
		set myBase = nothing
		set XMLDOM = nothing
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	generates the webservice
	'' @DESCRIPTION:	if some error happens then rootnode "ERROR" is returned with the
	''					Errormessage and errorcode. e.g. <Error coder="1000">..</Error>
	''					See for codes at properties part. ERRCODE...
	'**********************************************************************************************************
	public sub generate()
		with myBase
			if errorHandling then on error resume next
			.openDatabaseConnection()
			.setHTTPHeader()
			if isWebserviceHost() and err = 0 then
				.drawContent()
			elseif not isWebserviceHost() then
				errorMsg = "Call not on appropriate Host. " & ucase(request.serverVariables("SERVER_NAME")) & " is not a valid webservice-host."
				addError ERRCODE_WRONGHOST, errorMsg
			end if
			if err <> 0 then
				addError ERRCODE_UNEXPECTED, err.source & " - " & err.description
			end if
			
			if not XMLDOM.hasChildNodes() then addError ERRCODE_INVALIDXML, "Invalid Webservice (XML)"
			
			with XMLDOM
				set PINF = .createProcessingInstruction("xml", "version=""1.0""")
				.insertBefore PINF, .childNodes(0)
				.save(response)
			end with
			.closeDatabaseConnection()
			.destruct()
		end with
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	gets a new node of the tree
	'' @PARAM:			name [string]: name of your node
	'' @PARAM:			value [string]: value the node should have
	'' @RETURN:			[node]
	'**********************************************************************************************************
	public function getNewNode(name, value)
		set getNewNode = XMLDOM.createElement(name)
		getNewNode.appendChild(XMLDOM.createTextNode(value))
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	you can add custom-errors. e.g if something fails you want to display an error
	'' @DESCRIPTION:	use it more than once and it will add more messages to the error-node
	''					it will use an ERROR-node with the errorCode ERRCODE_CUSTOM
	'' @PARAM:			errMessage [string]: Errormessage
	'**********************************************************************************************************
	public sub addCustomError(errMessage)
		addError ERRCODE_CUSTOM, errMessage
	end sub
	
	'**********************************************************************************************************
	'* addError 
	'**********************************************************************************************************
	private sub addError(errCode, errMessage)
		if not XMLDOM.hasChildNodes() then
			set errNode = getNewNode("Error", empty)
			errNode.setAttribute "code", errCode
			XMLDOM.appendChild(errNode)
		elseif XMLDOM.documentElement.nodeName <> "Error" then
			'we have to remove all inserted nodes till now
			for each node in XMLDOM.childnodes
				XMLDOM.removeChild(node)
			next
		end if
		
		set errNode = XMLDOM.selectSingleNode("Error")
		errNode.appendChild(getNewNode("Message", errMessage))
		
		'we use the errorhandler for any unexpected error happend.
		if errCode > ERRCODE_WRONGHOST then
			set myErroro = new ErrorHandler
			with myErroro
				.errorDuring = "consuming Webservice"
				.debuggingVar = Request.ServerVariables("URL")
				.errorObject(err)
				.logError()
				.sendErrorEmail()
			end with
			set myErroro = nothing
		end if
	end sub
	
	'**********************************************************************************************************
	'* isWebserviceHost 
	'**********************************************************************************************************
	private function isWebserviceHost()
		serverName = ucase(request.serverVariables("SERVER_NAME"))
		isWebserviceHost = (serverName = ucase(consts.webservicesHost))
	end function

end class
%>