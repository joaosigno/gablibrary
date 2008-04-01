<!--#include virtual="/gab_libraryConfig/_fileUpload.asp"-->
<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		Fileupload
'' @CREATOR:		Michael Rebec, Michal Gabrukiewicz
'' @CREATEDON:		31.10.2003
'' @CDESCRIPTION:	Uploads a file to a specified location on the server
''					Note: Don't forget to set the enctype="multipart/form-data" on the form tag
'' @VERSION:		1.0
'' @POSTFIX:		fu

'**************************************************************************************************************
const ERR_FILEUPLOAD_NOFILE = "NFS"
const ERR_FILEUPLOAD_OVERWRITE = "OVR"
const ERR_FILEUPLOAD_FILESIZE = "FTL"
const ERR_FILEUPLOAD_EXTENSION = "EXT"
const ERR_FILEUPLOAD_BADFILENAME = "BFN"

class Fileupload

	private logID, m_uFilename
	private m_strFile, m_errorDesc, m_errorCode
	private p_uploadPath
	
	public fieldName							''[string] the name of the formField in your form. The field must be type=file.
	public defaultFilename						''[string] name the file should be renamed when saving it on the server. 
												''name without extension. extension will be taken from the source-file.
												''leave empty to keep the original name.
	public maxFileSize							''[int] the maximum size of the file in bytes. default is 100000 (100KB)
	public overwrite							''[bool] true if overwriting is allowed; default is false
	public allowedExtensions					''[string] includes the allowed extensions - default are txt files. seperate more with ","
	public uploaderObj							''[object] thats the uploaderObj-object. e.g. w3.upload. you can access it
	public serverLimit							''[int] the value of AspMaxRequestEntityAllowed set on the server in bytes. defines the
												''maximum allowed size of the form. if the user tries to upload more than this allowed size
												''this class throws an error.
	public saveUnique							''[bool] if overwriting is off use this to "index" the file when it exists. example: haxn.gif -> haxn(1).gif, etc.
												''useful if one name is neeeded for a file. default = false
	
	public property let uploadPath(val) ''[string] virtual target path where the file will be uploaded to e.g. /dir/subdir/. default = consts.userFiles
		p_uploadPath = str.ensureSlash(val)
	end property
	
	public property get uploadPath ''[string] gets the path where the file will be uploaded
		uploadPath = p_uploadPath
	end property
	
	public property get getErrorMsg ''[string] if an error occurs you can get the errorMsg with this property after using upload()
		getErrorMsg = m_errorDesc
	end property
	
	public property get getErrorCode ''[string] gets the errorcode if an error occured after upload(). e.g. NFS for no file given, etc.
		getErrorCode = m_errorCode
	end property
	
	public property let fileName(val) ''[string] OBSOLETE! use fieldname instead
		fieldname = val
	end property
	
	'**************************************************************************************************************
	'* constructor 
	'**************************************************************************************************************
	public sub class_Initialize()
		serverLimit = lib.init(GL_FU_SERVERLIMIT, 150000)
		if request.totalBytes > serverLimit then lib.error("Server does not accept this filesize (maximum: " & formatnumber(serverLimit / 1000) & "KB).")
		set uploaderObj	= server.createObject("w3.upload")
		'set the form of the library to the form of the uploader object
		set lib.form = uploaderObj.form
		maxFileSize = 100000
		p_uploadPath = consts.userFiles
		fieldName = empty
		overwrite = false
		allowedExtensions = "txt"
		defaultFilename	= empty
		logID = "fileUpload"
		saveUnique = false
	end sub
	
	'**************************************************************************************************************
	'* destructor 
	'**************************************************************************************************************
	public sub class_terminate()
		set uploaderObj = nothing
	end sub
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	call this to upload the file 
	'' @DESCRIPTON:		you can get the errorMsg with "getErrorMsg" or "getErrorCode"
	''					the following error codes are returned:<br>
	''					<strong>NFS</strong> = no file selected<br>
	''					<strong>OVR</strong> = overwriting not allowed<br>
	''					<strong>FTL</strong> = file toooo large<br>
	''					<strong>EXT</strong> = some problem with the extension<br>
	''					<strong>BFN</strong> = bad filename and/or extension
	'' @RETURN:			[string] returns name of uploaded file if successfull (name + extension). otherwise empty string
	'**************************************************************************************************************
	public function upload()
		upload = empty
		
		m_errorDesc = empty
		m_errorCode = empty
		set m_uFilename = lib.RF(fieldName)
		
		'validate some things...
		if not m_uFilename.isFile then
			errorHandle ERR_FILEUPLOAD_NOFILE, "No file selected."
			exit function
		end if
		
		m_strFile = m_uFilename.filename
		if defaultFilename <> empty then m_strFile = defaultFilename & "." & str.splitValue(m_strFile, ".", -1)
		fileSize = round((m_uFilename.size), 2)
		
		if not isValidFilename() then
			errorHandle ERR_FILEUPLOAD_BADFILENAME, "The file and/or its extension has an invalid name. Remove all " & getBadChars(false)
			exit function
		end if
		if not isAllowedExtension() then
			errorHandle ERR_FILEUPLOAD_EXTENSION, "Wrong extension"
			exit function
		end if
		if fileSize > maxFileSize then
			errorHandle ERR_FILEUPLOAD_FILESIZE, "File is too large: " & fileSize / 1000 & "KB"
			exit function
		end if
		if not overwriteAllowed() then
			errorHandle ERR_FILEUPLOAD_OVERWRITE, "The file already exists and overwriting is not allowed! You need to rename the file."
			exit function
		end if
		
		'the actual file uploading...
		targetPath = server.mapPath(p_uploadPath & m_strFile)
		if not overwrite and saveUnique then
			upload = str.splitValue(m_uFilename.saveAsUniqueFile(targetPath), "\", -1)
		else
		    m_uFilename.saveToFile(targetPath)
			upload = m_strFile
		end if
		lib.LogAndForget logID, "Uploaded: " & p_uploadPath & upload
	end function
	
	'**************************************************************************************************************
	'* errorHandle 
	'**************************************************************************************************************
	private sub errorHandle(errCode, errMessage)
		m_errorDesc = "ERROR: " & errMessage
		m_errorCode	= errCode
		if m_uFilename.isFile then
			lib.logAndForget logID, errCode & " " & server.mappath(p_uploadPath) & "\" & m_strFile
		else
			lib.logAndForget logID, errCode
		end if
	end sub
	
	'**************************************************************************************************************
	'* isValidFilename 
	'**************************************************************************************************************
	private function isValidFilename()
		set regex = new Regexp
		with regex
			.pattern = getBadChars(true)
			.ignoreCase = true
			.global = true
			isValidFilename = not .test(m_strFile)
		end with
		set regex = nothing
	end function
	
	'**************************************************************************************************************
	'* getBadChars 
	'**************************************************************************************************************
	private function getBadChars(asRegexPattern)
		getBadChars = "\,<>?:*/"
		if asRegexPattern then getBadChars = "[\\,<>\?:\*/]"
	end function
	
	'**************************************************************************************************************
	'* overwriteAllowed 
	'**************************************************************************************************************
	private function overwriteAllowed()
		overwriteAllowed = false
		if overwrite or (not overwrite and saveUnique) then
			overwriteAllowed = true
		else
			overwriteAllowed = not lib.fileExists(str.ensureSlash(p_uploadPath) & m_strFile)
		end if
	end function
	
	'**************************************************************************************************************
	'* isAllowedExtension 
	'**************************************************************************************************************
	private function isAllowedExtension()
		isAllowedExtension = false
		fileTypeArray = split(trim(allowedExtensions), ",")
		for i = 0 to uBound(fileTypeArray)
			if str.endsWith(uCase(m_strFile), uCase(trim(fileTypeArray(i)))) then isAllowedExtension = true
		next
	end function

end class
lib.registerClass("FileUpload")
%>