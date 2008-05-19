<!--#include file="class_webserviceParameter.asp"-->
<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		Webservice
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		23.02.2006 13:39
'' @CDESCRIPTION:	Gives you the possibility to create a "webservice" using common ASP and also to load (consume) the webservice
''					and get access to its metadata and the data itself. It is not a real webservice
''					according to the W3C standard but it allows communication between systems with a given
''					simple XML-contract.
''					provide a main()-method and use the XMLDOM-property to set all your data, then
''					use the generate()-method to generate the webservice.
''					The webservice XML output has an envelope-node as root which holds a header- and body-node.
''					Header has the information about the service itself whereas the body holds the data specific
''					for the webservice.
''					You will always get a valid XML-file, even if there are any errors in the webservice.
''					On an error you get an ERROR-node within the body-node holding the error.
''					NOTE: You need to load the GeneratePage in order to use it
'' @REQUIRES:		GeneratePage, ErrorHandler, TextTemplate
'' @POSTFIX:		WS
'' @VERSION:		0.4

'**************************************************************************************************************
class Webservice

	'private members
	private myBase
	
	'public members
	public XMLDOM					''[xmldom] the xmldom used within the webservice. and it represents the webservice body.
	public errorHandling			''[bool] disable/enable error handling. good for development. default = true
	public method					''[string] the method which should be used to call the webservice (POST or GET). default = post
	public description				''[string] a verbal description of the webservice. what does it do
	public name						''[string] name of the webservice. set automatically when creating.
	public url						''[string] url. set automatically when creating
	public params					''[webserviceParameter] collection of the parameters the webservice can approach
	
	public property let DBConnectionParameter(val)
		myBase.DBConnectionParameter = val
	end property
	
	public property let onlyWebDev(val)
		myBase.onlyWebDev = val
	end property
	
	public property let DBConnection(val)
		myBase.DBConnection = val
	end property
	
	public property get ERRCODE_CUSTOM ''[int] gets the error-code for custom-errors
		ERRCODE_CUSTOM = 1
	end property
	
	public property get ERRCODE_WRONGHOST ''[int] used to indicate that the webservice has been executed on the wrong host.
		ERRCODE_WRONGHOST = 2
	end property
	
	public property get ERRCODE_UNEXPECTED ''[int] gets error-code if any unexpected ASP error happens
		ERRCODE_UNEXPECTED = 4
	end property
	
	public property get ERRCODE_INVALIDXML ''[int] OBSOLETE! used in older versions.
		ERRCODE_INVALIDXML = 8
	end property
	
	public property get ERRCODE_MAINTENANCE	''[int] get error-code if the page is on maintenance
		ERRCODE_MAINTENANCE = 16
	end property
	
	public property get body ''[node] gets the body node. nothing if not available
		set body = XMLDOM.selectSingleNode("envelope/body")
	end property
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		lib.require(array("TextTemplate", "ErrorHandler", "GeneratePage"))
		set myBase = new GeneratePage
		myBase.loginRequired = false
		set XMLDOM = server.createObject("Microsoft.XMLDOM")
		set params = server.createObject("scripting.dictionary")
		myBase.isXML = true
		errorHandling = true
		method = "POST"
		description = ""
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
		'every webservice must have the extension .webservice!
		if lCase(str.splitValue(request.serverVariables("URL"), ".", -1)) <> "webservice" then
			'TODO!!!!'
			'lib.error("Webservice file must have the extension .webservice")
		end if
		
		if not errorHandling then myBase.isXML = false
		
		with myBase
			if errorHandling then on error resume next
			.openDatabaseConnection()
			.setHTTPHeader()
			
			if not .hasAccessOnMaintenance() then
				errorMsg = "Maintenance work."
				addError ERRCODE_MAINTENANCE, errorMsg
			elseif err = 0 then
				.drawContent()
			end if
			
			if err <> 0 then
				src = err.source
				descr = err.description
				'disable error handling because the handling of the webservice content is done.
				on error goto 0
				addError ERRCODE_UNEXPECTED, src & " - " & descr
			end if
			
			with XMLDOM
				set nEnv = getNewNode("envelope", "")
				nEnv.appendChild(getHeaderNode())
				set nBody = getNewNode("body", "")
				nBody.appendChild(.childNodes(0))
				nEnv.appendChild(nBody)
				XMLDOM.appendChild(nEnv)
				set PINF = .createProcessingInstruction("xml", "version=""1.0""  encoding=""UTF-8""")
				.insertBefore PINF, .childNodes(0)
				.save(response)
			end with
			
			.closeDatabaseConnection()
			.destruct()
		end with
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	adds a parameter to the input parameters list of the webservice
	'' @PARAM:			param [webserviceParameter]: the parameter you want to add
	'**********************************************************************************************************
	public sub addParam(param)
		params.add lib.getUniqueID(), param
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	has the webservice an error. useful after calling consume(). if yes it returns an array
	''					with the code and the message
	'' @RETURN:			[array] if array returned then an error is held by the webservice.. otherwise empty
	'**********************************************************************************************************
	public function getError()
		getError = empty
		set eNode = body.selectSingleNode("Error")
		if not eNode is nothing then getError = array(eNode.getAttribute("code"), eNode.text)
	end function
	
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
	'* getHeaderNode 
	'**********************************************************************************************************
	private function getHeaderNode()
		set getHeaderNode = getNewNode("header", "")
		with getHeaderNode
			.appendChild(getNewNode("name", str.splitValue(str.splitValue(request.serverVariables("URL"), ".", 0), "/", -1)))
			.appendChild(getNewNode("url", "http://" & request.serverVariables("HTTP_HOST") & request.serverVariables("URL")))
			.appendChild(getNewNode("description", description))
			.appendChild(getNewNode("method", uCase(method)))
			set nParams = getNewNode("parameters", "")
			for each p in params.items
				set nP = getNewNode("param", p.description)
				nP.setAttribute "name", p.name
				nP.setAttribute "dataType", p.dataType
				nP.setAttribute "defaultValue", p.defaultValue
				nParams.appendChild(nP)
			next
			.appendChild(nParams)
		end with
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
	'' @SDESCRIPTION:	consumes a given webservice which was created with the webservice class
	'' @DESCRIPTION:	webservice can be called without any parameters in order to get the metadata about the
	''					service itself. dont forget that the url has to start with http.
	''					Its useful to call getError() to check if an error happend. use body property to access the body
	'' @PARAM:			requestedUrl [string]: full url of the webservice. e.g. http://www....
	'' @PARAM:			requestMethod [string]: POST or GET are supported. leave empty if just metadata needed (GET will be used)
	'' @PARAM:			inputParams [dictionary], [array]: dictionary with parameters. key = parameter-name, item = param value
	''					or an array which looks like array(array(name, value), ..)
	'' @PARAM:			timoutAfter [int]: timout for the request. amount of seconds. 0 = infinite
	'' @RETURN:			[webservice] returns a webservice object. nothing if it cannot be loaded for some reasons.
	''					e.g. not available, could not parse XML, not a gablib webservice, etc.
	'**********************************************************************************************************
	public function consume(requestedUrl, requestMethod, inputParams, timoutAfter)
		failed = false
		set consume = new Webservice
		
		if errorHandling then on error resume next
		'4.0 version cannot be used due to the following problem on WIN2003 server
		'http://support.microsoft.com/default.aspx?scid=kb;en-us;820882#6
		set xmlhttp = createObject("Msxml2.ServerXMLHTTP.3.0")
		
		'apply the parameters
		if isArray(inputParams) then
			for each p in inputParams
				pQS = pQS & p(0) & "=" & server.URLEncode(p(1)) & "&"
			next
		elseif not inputParams is nothing then
			for each pName in inputParams.keys
				pQS = pQS & pName & "=" & server.URLEncode(inputParams(pName)) & "&"
			next
		end if
		
		async = false
		timeout = timoutAfter * 1000
		xmlhttp.setTimeouts timeout, timeout, timeout, timeout 'resolve, connect, send, receive
		
		if lCase(requestMethod) = "get" then
			xmlhttp.open "GET", requestedUrl & "?" & pQS, async
			xmlhttp.send()
		else
			xmlhttp.open "POST", requestedUrl, async
			xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			xmlhttp.setRequestHeader "Encoding", "UTF-8"
			xmlhttp.send(pQS)
		end if
		
		failed = not consume.xmldom.loadxml(xmlhttp.responseText)
		if consume.xmldom.parseError.errorCode <> 0 or err <> 0 then failed = true
		on error goto 0
		if failed then
			set consume = nothing
			exit function
		'check if gablib webservice
		elseif consume.xmldom.selectSingleNode("envelope/header") is nothing then
			set consume = nothing
			exit function
		end if
		
		consume.name = consume.xmldom.selectSingleNode("envelope/header/name").text
		consume.description = consume.xmldom.selectSingleNode("envelope/header/description").text
		consume.url = consume.xmldom.selectSingleNode("envelope/header/url").text
		consume.method = consume.xmldom.selectSingleNode("envelope/header/method").text
		set nParams = consume.xmldom.selectSingleNode("envelope/header/parameters")
		if nParams.hasChildNodes() then
			for each nParam in nParams.childNodes
				set p = new WebserviceParameter
				p.name = nParam.getAttribute("name")
				p.dataType = nParam.getAttribute("dataType")
				p.defaultValue = nParam.getAttribute("defaultValue")
				p.description = nParam.text
				consume.addParam(p)
			next
		end if
	end function
	
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
		if errCode = ERRCODE_UNEXPECTED then
			set aError = new ErrorHandler
			with aError
				.errorDuring = "consuming Webservice"
				.debuggingVar = request.ServerVariables("URL")
				.alternativeText = errMessage
				.sendMail()
				.log()
			end with
			set aError = nothing
		end if
	end sub

end class
lib.registerClass("Webservice")
%>