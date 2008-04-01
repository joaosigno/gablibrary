<% set lib = new Library %>
<!--#include virtual="/gab_Library/class_constants/constants.asp"-->
<!--#include virtual="/gab_Library/class_string/string.asp"-->
<!--#include virtual="/gab_LibraryConfig/class_customLib.asp"-->
<!--#include virtual="/gab_Library/class_createDropdown/createDropdown.asp"-->
<!--#include virtual="/gab_library/class_logger/logger.asp"-->
<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		Library
'' @CREATOR:		Michal Gabrukiewicz - gabru @ grafix.at
'' @CREATEDON:		12.09.2003
'' @STATICNAME:		lib
'' @CDESCRIPTION:	Often used functions. Often used in classes too. This class is loaded within the
''					Page-class. So you dont need to init it when working with page-class.
''					An instance called "lib" is available. You will have access to the lib methods with
''					"lib.myMethod". This class also makes an instance of customLib. All general methods
''					are available in lib-class. All custom methods should appear in lib.custom-class.
''					It also represents the gabLibrary itself. e.g. you register used classes, query
''					how many database access(es) has been done, etc.
'' @VERSION:		1.0

'**************************************************************************************************************

dim consts, lib, str
set consts = new Constants
set str = new StringOperations

class Library

	'private members
	private uniqueID, p_numberOfDBAccess, p_extensionIcons, p_gablibIconsPath, p_browser, p_form
	
	'public members
	public databaseConnection	''[ADODBConnection] Holds the Database-connection
	public custom				''[customLib] CustomLib-instance loaded automatically with library.
								''So you can use your custom-methods with using lib.custom.myMethod and lib-mthods using lib.myMethod
	public page					''[GeneratePage] currently executing page. just needed for control developers.
	public FSO					''[FileSystemObject] an instance of fileSystemObject for quick use ;)
	public registeredClasses	''[dictionary] collection of registered classes. key = name of the class (lcase)
	
	public property set form(val) ''[NameValueCollection] sets the reference to the form which is used. when using a multipart/form-data form, then its useful to assign the collection to a components formcollection. e.g. W3Upload. ONLY ADVANCED USE!
		'if the type is the same we dont set it because
		'e.g. for 3rd party upload components its only allowed to refer to one.
		if typename(p_form) = typename(val) then exit property
		set p_form = val
	end property
	
	public property get form ''[NameValueCollection] gets the currently used form. Accessible only if page is set because then the encoding for the response has been already set
		if me.page is nothing then throwError(array(1, "lib.form", "Cannot access 'lib.form' without an existing instance of 'GeneratePage'."))
		set form = p_form
	end property
	
	public property get numberOfDBAccess ''[int] returns the number of executed queries on the database
		numberOfDBAccess = p_numberOfDBAccess
	end property
	
	public property get version ''[string] gets the version of the gabLibrary installation
		version = "1.1"
	end property
	
	public property get browser	''[string] gets the browser (uppercased shortcut) which is used. version is ignored. e.g. FF, IE, etc. empty if unknown
		if p_browser = "" then
			agent = uCase(request.serverVariables("HTTP_USER_AGENT"))
			if instr(agent, "MSIE") > 0 then
				p_browser = "IE"
			elseif instr(agent, "FIREFOX") > 0 then
				p_browser = "FF"
			end if
		end if
		browser = p_browser
	end property
	
	'***********************************************************************************************************
	'* constructor 
	'***********************************************************************************************************
	public sub class_Initialize()
		set me.page = nothing
		set databaseConnection = nothing
		set p_extensionIcons = server.createObject("scripting.dictionary")
		set registeredClasses = server.createObject("scripting.dictionary")
		set FSO = server.createObject("Scripting.FileSystemObject")
		set form = request.form
		p_gablibIconsPath = empty
		p_numberOfDBAccess = 0
		uniqueID = 0
		set custom = new CustomLib
		p_browser = ""
	end sub
	
	'***********************************************************************************************************
	'* Destructor 
	'***********************************************************************************************************
	public sub class_Terminate()
		set FSO = nothing
		set p_extensionIcons = nothing
		set custom = nothing
		set databaseConnection = nothing
		set registeredClasses = nothing
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	calls a given function/sub if it exists
	'' @DESCRIPTION:	tries to call a given function/sub with the given parameters.
	''					the scope is the scope when calling exec. 
	'' @PARAM:			params [variant]: you choose how you provide your params. provide empty to call a procedure without parameters
	'' @RETURN:			[variant] whatever the function returns
	'**********************************************************************************************************
	function exec(functionName, params)
		on error resume next
		set func = getRef(functionName)
		available = (err = 0)
		on error goto 0
		if available then
			if isEmpty(params) then
				exec = func
			else
				exec = func(params)
			end if
		end if
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	gets a new dictionary filled with a list of values
	'' @PARAM:			values [array]: values to fill into the dictionary. array( array(key, value), arrray(key, value) )
	''					if the fields are not arrays (valuepairs) then the key is generated automatically. if no array
	''					provided then an empty dictionary is returned
	'' @RETURN:			[dictionary] dictionary with values.
	'**********************************************************************************************************
	public function newDict(values)
		set newDict = server.createObject("scripting.dictionary")
		if not isArray(values) then exit function
		for each v in values
			if isArray(v) then
				newDict.add v(0), v(1)
			else
				newDict.add lib.getUniqueID(), v
			end if
		next
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	throws an ASP runtime Error which can be handled with on error resume next
	'' @DESCRIPTION:	if you want to throw an error where just the user should be notified use lib.error instead
	'' @PARAM:			args [array], [string]: 
	''					- if array then fields => number, source, description
	''					- The number range for user errors is 512 (exclusive) - 1024 (exclusive)
	''					- if args is a string then its handled as the description and an error is raised with the
	''					number 1024 (highest possible number for user defined VBScript errors)
	'**********************************************************************************************************
	public sub throwError(args)
		if isArray(args) then
			if ubound(args) < 2 then me.throwError("To less arguments for throwError. must be (number, source, description)")
			if args(0) <= 0 then me.throwError("Error number must be greater than 0.")
			nr = 512 + args(0)
			source = args(1)
			description = args(2)
		else
			nr = 1024
			description = args
			source = request.serverVariables("SCRIPT_NAME")
			if lib.QS(empty) <> "" then source = source & "?" & lib.QS(empty)
		end if
		'user errors start after 512 (VB spec)
		err.raise nr, source, description
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	expands a given array by a given value or values of a given array and returns the new array
	'' @DESCRIPTION:	useful if you have an array and want to add a value to the end. Note: as a side effect
	''					you can use it to reverse a given array by providing no array to expand and the values
	''					as array (values will be reversed)
	'' @PARAM:			arr [array]: array which should be expanded. if no given then a new will be created
	'' @PARAM:			value [variant], [array]: value which should be added to the end. if its an array
	''					then all values will be added to the end. Note: the values will not be added in the right
	''					order. order will be reversed!
	'' @RETURN:			[array] the expanded array
	'**********************************************************************************************************
	public function expandArray(arr, byVal value)
		if not isArray(arr) then arr = array()
		if isArray(value) then
			if uBound(value) > -1 then
				redim preserve arr(ubound(arr) + 1)
				arr(uBound(arr)) = value(uBound(value))
				redim preserve value(uBound(value) - 1)
				expandArray = expandArray(arr, value)
			end if
		else
			redim preserve arr(ubound(arr) + 1)
			arr(uBound(arr)) = value
		end if
		expandArray = arr
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	registers a class that it has been loaded (included) and therefore can be used
	'' @DESCRIPTION:	this method should only be used for classes and by convention only at the end 
	''					of the class-definition. If the class has been already registered then an error is
	''					thrown. Makes it easier to debug.
	'' @PARAM:			className [string]: name of the class which should be registered
	'**********************************************************************************************************
	public sub registerClass(className)
		clsName = lCase(className)
		if clsName = "" then lib.error("Cannot register empty class name.")
		if registeredClasses.exists(clsName) then lib.error("'" & className & "' has already been registered (included).")
		registeredClasses.add clsName, empty
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	assumes that given class/classes is/are registered. if not error will be thrown.
	'' @DESCRIPTION:	should be used by component developers which require specific classes for their
	''					component (control, etc) and want to inform the client if its not registered.
	'' @PARAM:			classNames [string], [array]: name(s) of classes which are required at a certain place
	'**********************************************************************************************************
	public sub require(classNames)
		classesArr = classNames
		if not isArray(classesArr) then classesArr = array(classesArr)
		for i = 0 to uBound(classesArr)
			if not registeredClasses.exists(lCase(classesArr(i))) then lib.error("'" & classesArr(i) & "' is required (needs to be included).")
		next
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	gets a new globally unique Identifier.
	'' @RETURN:			[string]: new guid without hyphens. (hexa-decimal)
	'**********************************************************************************************************
	public function getGUID()
		getGUID = ""
	    set typelib = server.createObject("scriptlet.typelib")
	    getGUID = typelib.guid
	    set typelib = nothing
	    getGUID = mid(replace(getGUID, "-", ""), 2, 32)
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	initializes a variable with a default value if the variable is not set set (isEmpty)
	'' @PARAM:			var [variant]: some variable
	'' @PARAM:			default [variant]: the default value which should be taken if the var is not set
	'' @RETURN:			[variant] if var is set then the var otherwise the default value.
	'**********************************************************************************************************
	public function init(var, default)
		init = var
		if isEmpty(var) then init = default
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	gets an HTML-string with the appropriate icon for the filetype
	'' @PARAM:			file [file] file from FSO
	'' @RETURN:			[string] <img>-tag 
	'**********************************************************************************************************
	public function getFileIcon(file)
		extension = FSO.getExtensionName(file.path)
		
		'we check if an extension-icon exists on the file-system. to avoid checking on every-file
		'we store already checked extensions in an own collection
		if not p_extensionIcons.exists(extension) then
			if isEmpty(p_gablibIconsPath) then p_gablibIconsPath = server.mappath(consts.STDAPP("ICONS/"))
			if FSO.fileExists(p_gablibIconsPath & "\file_" & extension & ".gif") then
				p_extensionIcons.add extension, file.type
			else
				p_extensionIcons.add extension, "unknown"
			end if
		end if
		
		ext = extension
		if p_extensionIcons(extension) = "unknown" then
			ext = "unknown"
		end if
		getFileIcon = "<img border=0 align=absmiddle src=""" & consts.STDAPP("ICONS/file_" & ext & ".gif") & """ title=""" & uCase(extension) & " (" & p_extensionIcons(extension) & ")"">"
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes an error and ends the response. for unexpected error
	'' @DESCRIPTION:	this error should be used for any unexpected errors. for Exceptions!
	'' @PARAM:			msg [string], [array]: message or an array with messages which will be interpreted as lines
	'******************************************************************************************************************
	public sub [error](msg)
		response.clear()
		str.writeln("Erroro: ")
		if not isArray(msg) then
			str.write(str.HTMLEncode(msg))
		else
			str.write(vbCrLf)
			for each m in msg
				str.write(str.HTMLEncode(m) & vbCrLf)
			next
		end if
		str.end()
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	gets the value from a given form field after postback
	'' @DESCRIPTION:	just an equivalent for request.form.
	'' @PARAM:			name [string]: name of the value you want to get
	'' @RETURN:			[string], [object] value from the request-form-collection. if using an uploader then its likely
	''					that an object is returned for files.
	'******************************************************************************************************************
	public function RF(name)
		set theForm = form
		set RF = theForm(name)
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	gets the value from a given form field and encodes the string into HTML.
	''					useful if you want the value be HTML encoded. e.g. inserting into value fields
	'' @PARAM:			name [string]: name of the value you want to get
	'' @RETURN:			[string] value from the request-form-collection
	'******************************************************************************************************************
	public function RFE(name)
		RFE = str.HTMLEncode(RF(name))
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	gets the value of a given form field and trims it automatically.
	'' @PARAM:			name [string]: name of the formfield you want to get
	'' @RETURN:			[string] value from the request-form-collection
	'******************************************************************************************************************
	public function RFT(name)
		RFT = trim(RF(name))
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	returns true if a given value exists in the request.form
	'' @PARAM:			name [string]: name of the value you want to get
	'' @RETURN:			[bool] false if there is not value returned. true if yes
	'******************************************************************************************************************
	public function RFHas(name)
		RFHas = (trim(RF(name)) <> "")
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	just an equivalent for request.querystring. if empty then returns whole querystring.
	'' @PARAM:			name [string]: name of the value you want to get. leave it empty to get the whole querystring
	'' @RETURN:			[string] value from the request-querystring-collection
	'******************************************************************************************************************
	public function QS(name)
		if name = "" then
			QS = request.querystring
		else
			QS = request.querystring(name)
		end if
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	gets a value from the querystring. 
	'' @DESCRIPTION:	example: you want an ID from the querystring and be sure you get a number.
	''					so you call the method getFromQS("ID", 0) and you can be sure you get a number and if the
	''					parameter in the querystring is not a number then you get 0.
	'' @PARAM:			name [string]: name of the value you want to get. leave it empty to use the whole
	''					querystring as the value
	'' @PARAM:			alternative [variant]: an integer value with what you want to replace the 
	''					value when it cannot be parsed into the given type of alternative
	'' @RETURN:			[string] value from the request-querystring-collection
	'******************************************************************************************************************
	public function getFromQS(name, alternative)
		getFromQS = alternative
		value = QS(name)
		if isNumeric(alternative) then getFromQS = str.toInt(value, alternative)
	end function
	
	'***********************************************************************************************************
	'' @SDESCRIPTION:	executes a given javascript. input may be a string or an array. each array field = a line
	'' @PARAM:			javascriptCode [string]. [array]: your javascript-code. e.g. window.location.reload()
	'***********************************************************************************************************
	public sub execJS(javascriptCode)
		with str
			.writeln("<script language=JavaScript>")
			if isArray(javascriptCode) then
				for k = 0 to uBound(javascriptCode)
					.writeln(javascriptCode(k))
				next
			else
				.writeln(javascriptCode)
			end if
			.writeln("</script>")
		end with
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Returns HTML code which can be used within any tag to display a tooltip.
	'' @DESCRIPTION:	loadTooltips of page must be turned on (additional javascript will be loaded) before
	''					calling this method
	'' @PARAM:			title [string]: the title which should be displayed in the tooltip
	'' @PARAM:			msg [string]: the message of the tooltip
	'' @RETURN:			[string] onMouseover="...."
	'******************************************************************************************************************
	public function tooltip(title, msg)
		if not me.page.loadTooltips then lib.error("loadTooltips of the current page instance need to be set properly.")
		tooltip = " onmouseover=""balloonTooltip.showTooltip('" & str.JSEncode(title) & "', '" & str.JSEncode(msg) & "');"" onmouseout=""balloonTooltip.hideTooltip();"" "
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Sleeps for a specified time. for a while :)
	'' @PARAM:			seconds [int]: how many seconds. Minmum 1, Maximum 20. Value will be autochanged if value
	''					is incorrect. 
	'******************************************************************************************************************
	public sub sleep(seconds)
	    if seconds < 1 then
		    seconds = 1
	    elseIf seconds > 20 then
		    seconds = 20
	    end if
		
	    x = timer()
	    do until x + seconds = timer()
		    'sleep
	    loop
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Returns an unique ID for each pagerequest, starting with 1
	'' @RETURN:			uniqueID [int]: the unique id
	'******************************************************************************************************************
	public function getUniqueID()
		uniqueID = uniqueID + 1
		getUniqueID = uniqueID
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Does a simple log using the Logger-class
	'' @DESCRIPTION: 	This function exists because of simplier and faster logging. It uses the Logger-class
	''					Use it if you want to log something quite fast without "thinking" ;)
	''					Only use it if you log once within a site. If you log more than once please create your
	''					own Logger-instance. This method uses lib.custom.logAndForget. Its like an interface.
	''					Provide your own logging routine in customlib.
	''					Why log and forget? its inherited from fire and forget ...
	'' @PARAM:			identification [string]: The log identification. For more details check Logger-class
	'' @PARAM:			logMessage [string]: The string you want to log
	'******************************************************************************************************************
	public sub logAndForget(identification, logMessage)
		lib.custom.logAndForget identification, logMessage
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Opposite of server.URLEncode
	'' @PARAM:			- endcodedText [string]: your string which should be decoded. e.g: Haxn%20Text (%20 = Space)
	'' @RETURN:			[string] decoded string
	'' @DESCRIPTION: 	If you store a variable in the queryString then the variables input will be automatically
	''					encoded. Sometimes you need a function to decode this %20%2H, etc.
	'******************************************************************************************************************
	public function URLDecode(endcodedText)
    	decoded = endcodedText
	    
		set oRegExpr = server.createObject("VBScript.RegExp")
	    oRegExpr.pattern = "%[0-9,A-F]{2}"
	    oRegExpr.global = true
	    set matchCollection = oRegExpr.execute(endcodedText)
		
	    for each match in matchCollection
	    	decoded = replace(decoded, match.value, chr(cint("&H" & right(match.value, 2))))
	    next
		
	    URLDecode = decoded
    end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Checks if a file exists on the filesystem
	'' @PARAM:			- path for the file. e.g: "/images/haxn.gif"
	'' @RETURN:			[bool] if the file exists or not
	'******************************************************************************************************************
	public function fileExists(virtualPath)
		fileExists = (FSO.fileExists(server.mappath(virtualPath)))
	end function
	
	'******************************************************************************************************************
	'' @DESCRIPTION: 	This will replace the IIf function that is missing from the intrinsic functions of ASP
	'' @PARAM:			i [variant]: condition
	'' @PARAM:			j [variant]: expression 1
	'' @PARAM:			k [variant]: expression 2
	'' @RETURN:			[string]
	'******************************************************************************************************************
	public function iif(i, j, k)
    	if i then iif = j else iif = k
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	executes an sql-query and returns the first value of the first row.
	'' @DESCRIPTION:	if there is no record given then the noRecordReplacer will be returned.
	''					the returned value (if available) will be converted to the type of which the noRecordReplacer is.
	''					example: calling getScalar("...", 0) will convert the returned value into an integer. if no record
	''					then 0 will be returned
	'' @PARAM:			sql [string]: sql-query to be executed
	'' @PARAM:			noRecordReplacer [variant]: what should be returned when there is no record returned
	'' @RETURN:			[variant] the first value of the result converted to the type of noRecordReplacer
	''					or the noRecordReplacer
	'******************************************************************************************************************
	public function getScalar(sql, noRecordReplacer)
		if trim(sql) = "" then throwError(array(100, "lib.getScalar", "SQL-Query cannot be empty"))
		getScalar = noRecordReplacer
		set aRS = lib.getRecordset(sql)
		if not aRS.eof then getScalar = str.parse(aRS(0) & "", noRecordReplacer)
		set aRS = nothing
	end function
	
	'******************************************************************************************************************
	'' @DESCRIPTION: 	Default method which should be always used to get a LOCKED recordset. Example for use:
	''					set rs = lib.getRecordset("SELECT * FROM users")
	'' @PARAM:			sql [string]: Your SQL query
	'' @RETURN:			[object] recordset-Object
	'******************************************************************************************************************
	public function getRecordset(sql)
		if trim(sql) = "" then throwError(array(100, "lib.getRecordset", "SQL-Query cannot be empty"))
		if databaseConnection is nothing then lib.error("lib.databaseConnection is nothing. Check lib.custom.establishDatabaseConnection")
		on error resume next
 		set getRecordset = databaseConnection.execute(sql)
		if err <> 0 then
			errdesc = err.description
			on error goto 0
			throwError(array(101, "lib.getRecordset", "Could not execute '" & sql & "'. Reason: " & errdesc, sql))
		end if
		on error goto 0
		p_numberOfDBAccess = p_numberOfDBAccess + 1
	end Function
	
	'******************************************************************************************************************
	'' @DESCRIPTION: 	Default method which should be always used to get an UNLOCKED recordset. Example for use:
	''					set rs = lib.getUnlockedRecordset("SELECT * FROM users")
	'' @PARAM:			sql [string]: Your SQL query
	'' @RETURN:			[object] recordset-Object
	'******************************************************************************************************************
	public function getUnlockedRecordset(sql)
		if trim(sql) = "" then throwError(array(100, "lib.getUnlockedRecordset", "SQL-Query cannot be empty"))
		if databaseConnection is nothing then lib.error("lib.databaseConnection is nothing. Check lib.custom.establishDatabaseConnection")
		on error resume next
		set getUnlockedRecordset = server.createObject("ADODB.RecordSet")
		getUnlockedRecordset.cursorLocation = 3
		getUnlockedRecordset.cursorType = 3
		getUnlockedRecordset.open sql, databaseConnection
		p_numberOfDBAccess = p_numberOfDBAccess + 1
		if err <> 0 then
			errdesc = err.description
			on error goto 0
			throwError(array(101, "lib.getUnlockedRecordset", "Could not execute '" & sql & "'. Reason: " & errdesc, sql))
		end if
		on error goto 0
	end Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Get QueryString without specified variables.
	'' @DESCRIPTION: 	You can get the Querystring without unwanted QueryString-variables. 
	''					Seperate more parameters with ",". Example. You have a QueryString which looks like
	''					var=1&var2=3&unwanted=hallo and you want to get a QueryString without the "unwanted"-variable
	''					so you just call the function like this: getAllFromQueryStringBut("unwanted")
	'' @PARAM:			- query [string]: the parameter(s) you dont want to have. Seperate more with ","
	'' @RETURN:			[string] QueryString-String
	'******************************************************************************************************************
	public function getAllFromQueryStringBut(query)
		if instr(query, ",") then
			myQuery = split(query, ",")
		else
			dim myQuery(0)
			myQuery(0) = Query
		end if
		
		for each q in request.QueryString
			returnItem = true
			for i = 0 to uBound(myQuery)
		    	if lcase(q) = lCase(myQuery(i)) then returnItem = false
		  	next
			if returnItem then s = s & q & "=" & request.queryString(q).item & "&"
		next
		
		if len(s) > 0 then
			getAllFromQueryStringBut = Left(s, Len(s) - 1)
		else
			getAllFromQueryStringBut = ""
		end if
	end function

end class
%>
