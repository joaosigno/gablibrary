<%
'**************************************************************************************************************

'' @CLASSTITLE:		Fileupload
'' @CREATOR:		Michael Rebec
'' @CREATEDON:		31.10.2003
'' @CDESCRIPTION:	Uploads a file to a specified location on the server
'' @VERSION:		0.2

'**************************************************************************************************************
const ERR_FILEUPLOAD_NOFILE = "NFS"
const ERR_FILEUPLOAD_OVERWRITE = "OVR"
const ERR_FILEUPLOAD_FILESIZE = "FTL"
const ERR_FILEUPLOAD_EXTENSION = "EXT"

class Fileupload

	private m_extensions						' includes the allowed file extensions
	private m_objFSO							' file system object
	private m_uFilename							' the new target when file was uploaded
	private m_offset_old						' needed to get the filename without path
	private m_offset							' needed to get the filename without path
	private m_strFile							' inlcudes the relative filename, e.g. test.txt
	private m_fileSize							' includes the calculated filesize
	private m_errorDesc							' saves the Errormsg - you can check the msg with ".getErrorMsg"
	private m_errorCode							' saves the Errorcode - you can check the code with ".getErrorCode"
	private m_logFileName						' the name of the logfile
	
	public uploaderObj							'' [object] thats the uploaderObj-object. e.g. w3.upload. you can access it
	public maxFileSize							'' the maximum size of the file - in [bytes]
	public uploadPath							'' the target location of the file. e.g. /dir/subdir/
	public fileName								'' the name of the formField in your form. The field must be type=file.
	public file									'' In some cases (multiple files in one form) you need to provide the file object as source
												'' If you provide this member than fileName will be ignored and this object will be taken.
	public overwrite							'' true if overwriting is allowed; default is false
	public uploadTool							'' the used Upload Tool - default is "w3.upload"
	public allowedExtensions					'' includes the allowed extensions - default are all file types. seperate more with ","
	public defaultFilename						'' if this includes a filename, the targetfile will be named like this
	
	'Konstruktor => set the default values
	private sub Class_Initialize()
		maxFileSize 		= empty
		uploadPath 			= empty
		fileName 			= empty
		'set file			= nothing
		overwrite 			= false
		uploadTool 			= "w3.upload"
		set uploaderObj		= server.createObject(uploadTool)
		allowedExtensions	= "txt"
		defaultFilename		= empty
		m_logFileName		= "fileUpload"
	end sub
	
	'**************************************************************************************************************
	' writes the errors on screen
	'**************************************************************************************************************
	private sub errorHandle(short, strErr)
		m_errorDesc = "ERROR: " & strErr
		m_errorCode	= short
		if m_uFilename.IsFile then
			call lib.LogAndForget(m_logFileName, short & " " & server.mappath(uploadPath) & "\" & m_strFile)
		else
			call lib.LogAndForget(m_logFileName, short)
		end if
	end sub
	
	'**************************************************************************************************************
	' initialize the class and checks if everything is correct
	'**************************************************************************************************************
	private function init()
		init = false
			if not (allowedExtensions = empty) then m_extensions = Split(Trim(allowedExtensions),",")
			if not (uploadTool = empty) then
				'set uploaderObj = Server.CreateObject(uploadTool)
				if isObject(file) then
					set m_uFilename = file
				else
					set m_uFilename = uploaderObj.form(fileName)
				end if
			end if
			
			if m_uFilename.IsFile then
				Do
		        	m_offset_old = m_offset
			        m_offset = InStr(m_offset + 1, m_uFilename.filename, "\")
		    	Loop Until m_offset = 0
				m_strFile = Mid(m_uFilename.filename, m_offset_old + 1)
				m_fileSize = round((m_uFilename.size),2)
				if m_fileSize > maxFileSize then
					call errorHandle(ERR_FILEUPLOAD_FILESIZE, "Filesize to large: " & m_fileSize / 1000 & "KB")
					exit function
				end if
				if not overwriteAllowed(m_strFile) then
					call errorHandle(ERR_FILEUPLOAD_OVERWRITE, "Overwriting not allowed! File already exists. You can rename the file.")
					exit function
				end if
			else
				call errorHandle(ERR_FILEUPLOAD_NOFILE, "No file selected.")
				exit function
			end if
		init = true
	end function
	
	'**************************************************************************************************************
	' checks if a file allready exists and if overwriting is allowed
	'**************************************************************************************************************
	private function overwriteAllowed(myfilename)
		overwriteAllowed = true
		Set m_objFSO = Server.CreateObject("Scripting.FileSystemObject")
		if m_objFSO.FileExists(server.mappath(uploadPath) & "\" & myfilename) and not overwrite then _
			overwriteAllowed = false
	end function
	
	'**************************************************************************************************************
	' saves/uploads the file on the server
	'**************************************************************************************************************
	private function uploadFile()
		if checkExtension(m_strFile, m_extensions) then
			if not (defaultFilename = empty) then m_strFile = defaultFilename & "." & str.splitValue(m_strFile, ".", -1)
		    m_uFilename.SaveToFile(Server.Mappath(uploadPath & m_strFile))
			uploadFile = m_strFile
		else
			call errorHandle(ERR_FILEUPLOAD_EXTENSION, "wrong extension")
			uploadFile = ""
		end if
	end function
	
	'**************************************************************************************************************
	' checks if the file extensions are right, e.g. "txt" / "txt, jpg"
	'**************************************************************************************************************
	private function checkExtension(myfilename, fileTypeArray)
		checkExtension = False
		For i = 0 To UBound(fileTypeArray)
			If UCASE(Right(myfilename, Len(fileTypeArray(i)))) = UCASE(fileTypeArray(i)) Then
				checkExtension = True
			end if
		Next
	End Function
	
	'**************************************************************************************************************
	'' @DESCRIPTION:	call this to upload the file 
	'' @RETURN:			returns filename of uploaded file if successfull. Else you can get the errorMsg with .getErrorMsg
	'**************************************************************************************************************
	public function upload()
		if init() then tmp = uploadFile()
		if (tmp = "") then 
			upload = false
		else
			upload = tmp
			call lib.LogAndForget(m_logFileName, "Uploaded: " & uploadPath & m_strFile)
		end if
	end function
	
	public property get getErrorMsg ''if an error occurs you can get the errorMsg with this property
		getErrorMsg = m_errorDesc
	end property
	
	public property get getErrorCode
		getErrorCode = m_errorCode
	end property

end class
%>