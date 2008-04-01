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
'' @VERSION:		0.3

'**************************************************************************************************************
const ERR_FILEUPLOAD_NOFILE = "NFS"
const ERR_FILEUPLOAD_OVERWRITE = "OVR"
const ERR_FILEUPLOAD_FILESIZE = "FTL"
const ERR_FILEUPLOAD_EXTENSION = "EXT"

class Fileupload

	private logID, m_uFilename, m_offset_old, m_offset
	private m_strFile, m_fileSize, m_errorDesc, m_errorCode
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
		set uploaderObj	= server.createObject("w3.upload")
		if request.totalBytes > serverLimit then lib.error("Server does not accept this filesize.")
		'set the form of the library to the form of the uploader object
		'if the form is already a form of an uploader we leave it...
		if typename(lib.form) <> typename(uploaderObj.form) then set lib.form = uploaderObj.form
		maxFileSize = 100000
		p_uploadPath = consts.userFiles
		fieldName = empty
		overwrite = false
		allowedExtensions = "txt"
		defaultFilename	= empty
		logID = "fileUpload"
	end sub
	
	'**************************************************************************************************************
	'* destructor 
	'**************************************************************************************************************
	public sub class_terminate()
		set uploaderObj = nothing
	end sub
	
	'**************************************************************************************************************
	'' @DESCRIPTION:	call this to upload the file 
	'' @RETURN:			[string] returns name of uploaded file if successfull (name + extension). otherwise empty string
	''					you can get the errorMsg with "getErrorMsg" or "getErrorCode"
	'**************************************************************************************************************
	public function upload()
		upload = empty
		if init() then
			if checkExtension() then
				if defaultFilename <> empty then m_strFile = defaultFilename & "." & str.splitValue(m_strFile, ".", -1)
			    m_uFilename.SaveToFile(server.Mappath(p_uploadPath & m_strFile))
				upload = m_strFile
				lib.LogAndForget logID, "Uploaded: " & p_uploadPath & m_strFile
			else
				errorHandle ERR_FILEUPLOAD_EXTENSION, "wrong extension"
			end if
		end if
	end function
	
	'**************************************************************************************************************
	'* errorHandle 
	'**************************************************************************************************************
	private sub errorHandle(errCode, errMessage)
		m_errorDesc = "ERROR: " & errMessage
		m_errorCode	= errCode
		if m_uFilename.isFile then
			lib.logAndForget logID, errCode & " " & server.mappath(p_uploadPath & m_strFile)
		else
			lib.logAndForget logID, errCode
		end if
	end sub
	
	'**************************************************************************************************************
	'* init 
	'**************************************************************************************************************
	private function init()
		m_errorDesc = empty
		m_errorCode = empty
		
		init = false
		
		set m_uFilename = lib.form(fieldName)
		
		if m_uFilename.isFile then
			Do
	        	m_offset_old = m_offset
		        m_offset = inStr(m_offset + 1, m_uFilename.filename, "\")
	    	Loop Until m_offset = 0
			m_strFile = mid(m_uFilename.filename, m_offset_old + 1)
			m_fileSize = round((m_uFilename.size), 2)
			if m_fileSize > maxFileSize then
				errorHandle ERR_FILEUPLOAD_FILESIZE, "Filesize to large: " & m_fileSize / 1000 & "KB"
				exit function
			end if
			if not overwriteAllowed() then
				errorHandle ERR_FILEUPLOAD_OVERWRITE, "The file already exists and overwriting is not allowed! You need to rename the file."
				exit function
			end if
		else
			errorHandle ERR_FILEUPLOAD_NOFILE, "No file selected."
			exit function
		end if
		
		init = true
	end function
	
	'**************************************************************************************************************
	'* overwriteAllowed 
	'**************************************************************************************************************
	private function overwriteAllowed()
		overwriteAllowed = true
		if lib.fileExists(str.ensureSlash(p_uploadPath) & m_strFile) and not overwrite then overwriteAllowed = false
	end function
	
	'**************************************************************************************************************
	'* checkExtension 
	'**************************************************************************************************************
	private function checkExtension()
		checkExtension = false
		fileTypeArray = split(trim(allowedExtensions), ",")
		for i = 0 To uBound(fileTypeArray)
			if str.endsWith(uCase(m_strFile), uCase(trim(fileTypeArray(i)))) then
				checkExtension = True
			end if
		next
	end function

end class
lib.registerClass("FileUpload")
%>