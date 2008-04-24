<!--#include virtual="/log/index/index.idx"-->
<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		Logger
'' @CREATOR:		Michael Rebec / Michal Gabrukiewicz - gabru @ grafix.at
'' @CREATEDON:		22.04.2004
'' @CDESCRIPTION:	Gives you an opportunity to log things to a log-file quick and easy.
'' @VERSION:		0.2

'**************************************************************************************************************
class Logger

	private ASP_IN					'"< %"
	private ASP_OUT					'"% >"
	private ASP_ENDLINE				''*endline*
	private PREFIX_PLACEHOLDER		'used for the message preifx
	private indexFilePath			'the virtual path of the index file, do NOT change this
	private loggerLCIDTimeformat	'3079 for Austria
	private fso						'the FileSystemObject
	private currentLogfilePath		'the Absolute Path of the log file, e.g. C:\LOGS\common_log.txt
	private currentLogFile			'the File Object for the current log file
	private logFileFromIndex		'the name of the logfile saved in the index file
	private indexModifier			'the modifier in the index file, e.g. "dim logidx_test : "
	private indexPrefix				'the prefix in the index file, e.g. "logidx_"
	private indexKeyExists			'[bool] true if a key allready exists in the index file
	private logMessagePrefix		'the prefix message in the log file, e.g. "[01.01.2004 13:48] "
	private logsPath				'the virtual path of the log file, default is "consts.logs_path"
	private logfileExtension		'the extension of the log file, default is ".txt"
	private modeForAppending, modeForReading, modeForWriting
	private openAsASCII, openAsUnicode, openAsSystem
	
	public identification			''[string] the name of the log file, default is "common_logs"
	public splittingSize			''[int] the maximum file size of the current log file in bytes, default is "100 000". 0 = never split
	public onlyOneLogFile			''[bool] if this flag is true and the splitting size has been reached, the old file will
									''be deleted and a new one will be created
	
	'**************************************************************************************************************
	'* constructor 
	'**************************************************************************************************************
	public sub class_initialize()
		set fso					= nothing
		set currentLogFile		= nothing
		ASP_IN					= "<%"
		ASP_OUT					= chr(37) & ">"
		ASP_ENDLINE				= "'*endline*"
		PREFIX_PLACEHOLDER		= "<<<DATE>>>"
		modeForAppending 		= 8
		modeForReading 			= 1
		modeForWriting 			= 2
		openAsASCII				= 0
		openAsUnicode			= -1
		openAsSystem			= -2
		loggerLCIDTimeformat	= 3079
		currentLogfilePath		= empty
		logFileFromIndex		= empty
		identification 			= "common_logs"
		logfileExtension 		= "log"
		logsPath				= str.ensureSlash(consts.logs_path)
		splittingSize			= 100000
		logMessagePrefix		= "[" & request.servervariables("REMOTE_ADDR") & " | " & PREFIX_PLACEHOLDER & "]" & vbTab
		indexFilePath			= "/log/index/index.idx"
		indexPrefix				= "logidx_"
		indexKeyExists			= false
		onlyOneLogFile			= false
	end sub
	
	'**************************************************************************************************************
	'* destructor 
	'**************************************************************************************************************
	public sub class_terminate()
		set fso = nothing
		set currentLogFile = nothing
	end sub
	
	'**************************************************************************************************************
	'' @SDESCRIPTION: 	use this to log
	'' @DESCRIPTION:	logs the message into the specified log file
	'' @PARAM:			- message: the text you want to log
	'**************************************************************************************************************
	public sub [log](message)
		indexModifier = "dim " & indexPrefix & identification & " : "
		if init() then
			if needsToBeSplitted() then
				if onlyOneLogFile then deleteOldFile()
				newFileName = getNewFileName()
				writeToIndexFile indexModifier & indexPrefix & identification & " = """ & newFileName & """"
				refreshFileNames(newFileName)
			end if
			writeToLogFile(message)
		else
			createNewLogFile(message)
		end if
	end sub
	
	'**************************************************************************************************************
	'' @SDESCRIPTION: 	use this to delete all log files, with the same name, e.g all "common_logs_...."
	'**************************************************************************************************************
	public sub deleteIdent()
		init()
		
		set folder = fso.getFolder(server.mappath(logsPath))
		for each file in folder.files
			'filename without extension
			filename = ucase(str.trimEnd(file.name, 4))
			if str.startsWith(filename, ucase(identification)) then file.delete()
		next
		set folder = nothing
	end sub
	
	'**************************************************************************************************************
	'' @SDESCRIPTION: 	Attention: this function deletes all log files from the logFilePath
	'**************************************************************************************************************
	public sub deleteAll()
		set fso = server.createObject("Scripting.FileSystemObject")
		set folder = fso.getFolder(server.mappath(logsPath))
		for each file in folder.files
			if str.endsWith(file.name, logfileExtension) then file.delete()
		next
		set folder = nothing
	end sub
	
	'**************************************************************************************************************
	'' @SDESCRIPTION: 	use this to reset the index file to default state
	'' @DESCRIPTION:	clears everything from the index file and writes the default properties into it
	'**************************************************************************************************************
	public sub clearIndex()
		set fso = server.createObject("Scripting.FileSystemObject")
		
		set streamWriter = fso.OpenTextFile(server.mappath(indexFilePath), 2)
		With streamWriter
			.WriteLine(ASP_IN)
			.WriteLine(ASP_ENDLINE)
			.WriteLine(ASP_OUT)
			.Close()
		End With
		set streamWriter = nothing
	end sub
	
	'**************************************************************************************************************
	'* init - initialize the FileObject
	'**************************************************************************************************************
	private function init()
		set fso = server.createObject("Scripting.FileSystemObject")
		refreshFileNames(empty)
		
		if lib.fileExists(logsPath & logFileFromIndex) then
			set currentLogFile = fso.getFile(currentLogfilePath)
			init = true
		else
			set currentLogFile = nothing
			init = false
		end if
	end function
	
	'**************************************************************************************************************
	'* deleteOldFile - deletes the old log file if onlyOneLogfile is enabled
	'**************************************************************************************************************	
	private sub deleteOldFile()
		currentLogFile.delete()
		set currentLogFile = nothing
	end sub
	
	'**************************************************************************************************************
	'* refreshFileNames - refreshes some needed paths/files/...
	'**************************************************************************************************************
	private sub refreshFileNames(tempIndexFileName)
		logFileFromIndex = tempIndexFileName
		if tempIndexFileName = empty then logFileFromIndex = getIndexFileEntry()
		
		currentLogfilePath = server.mappath(logsPath & logFileFromIndex)
		
		if lib.fileExists(logsPath & logFileFromIndex) then
			set currentLogFile = fso.getFile(currentLogfilePath)
		else
			createNewLogFile(empty)
		end if
	end sub
	
	'**************************************************************************************************************
	'* needsToBeSplitted - returns true if the file size of the log file is bigger than the splittingSize
	'**************************************************************************************************************
	private function needsToBeSplitted()
		needsToBeSplitted = (splittingSize <> 0 and currentLogFile.size >= splittingSize)
	end function
	
	'**************************************************************************************************************
	'* getIndexFileEntry - (pseudo) reads the index file, to get the last entry and creates a new one in there isn't one
	'**************************************************************************************************************
	private function getIndexFileEntry()
		execute("getIndexFileEntry = " & indexPrefix & identification)
		
		if getIndexFileEntry = "" then
			newIdent = getNewFileName()
			writeToIndexFile indexModifier & indexPrefix & identification & " = """ & newIdent & """"
			getIndexFileEntry = newIdent
		end if
	end function
	
	'**************************************************************************************************************
	'* getNewFileName - returns a new filename with the name and the date
	'**************************************************************************************************************
	private function getNewFileName()
		if splittingSize = 0 then
			getNewFileName = identification & "." & logFileExtension
		else
			getNewFileName = identification  & "_" & Year(now) & "_" & Month(now) & "_" & Day(now) & "_" & _
				Hour(now) & Minute(now) & Second(now) & "." & logFileExtension
		end if
	end function
	
	'**************************************************************************************************************
	'* getIndexFileStream - reads the index file
	'**************************************************************************************************************
	private function getIndexFileStream(msg)
		idxStream = empty
		set streamReader = fso.OpenTextFile(server.mappath(indexFilePath), 1)
		do while not streamReader.AtEndOfStream
			tmpStream = parseIndexStream(streamReader.ReadLine, msg)
			if not tmpStream = "" then idxStream = idxStream & tmpStream & vbCrLf
		loop
		getIndexFileStream = idxStream
		set streamReader = nothing
	end function
	
	'**************************************************************************************************************
	'* parseIndexStream - parses the index file
	'**************************************************************************************************************
	private function parseIndexStream(stream, msg)
		intSep = InStr(stream, "=")
		if intSep > 1 and not indexKeyExists then
			if str.startsWith(msg, left(stream, intSep)) then
				indexKeyExists = true
				if not stream = msg then
					parseIndexStream = msg & vbCrLf
					exit function
				end if
			end if
		end if
		
		if not indexKeyExists and stream = ASP_ENDLINE then stream = msg & vbCrLf & ASP_ENDLINE
		parseIndexStream = stream
	end function
	
	'**************************************************************************************************************
	'* writeToIndexFile - writes the index file
	'**************************************************************************************************************
	private sub writeToIndexFile(msg)
		idxStream = getIndexFileStream(msg)
		set streamWriter = fso.OpenTextFile(server.mappath(indexFilePath), 2)
		streamWriter.WriteLine(idxStream)
		streamWriter.Close()
		set streamWriter = nothing
	end sub
	
	'**************************************************************************************************************
	'* createNewLog - creates a new log file
	'**************************************************************************************************************
	private sub createNewLogFile(msg)
		'create the file now
		set streamX = fso.CreateTextFile(currentLogfilePath)
		streamX.close()
		set streamX = nothing
		
		'we created a new logFile and so we need to change currentLogFile
		set currentLogFile = fso.getFile(currentLogfilePath)
		
		if not msg = empty then writeToLogFile(msg)
	end sub
	
	'**************************************************************************************************************
	'* getMessagePrefix - sets and resets the LCID and replaces the Placeholder in the messagePrefix
	'**************************************************************************************************************
	private function getMessagePrefix()
		temporaryLCID = session.LCID
		session.LCID = loggerLCIDTimeformat
		tempPrefix = Replace(logMessagePrefix, PREFIX_PLACEHOLDER, now())
		getMessagePrefix = tempPrefix
		session.LCID = temporaryLCID
	end function
	
	'**************************************************************************************************************
	'* writeToLogFile - writes the message into the log file
	'**************************************************************************************************************
	private sub writeToLogFile(msg)
		set streamX = currentLogFile.openAsTextStream(modeForAppending, openAsASCII)
		with streamX
			.writeLine(getMessagePrefix() & msg)
			.close()
		end with
		set streamX = nothing
	end sub
	
end class
lib.registerClass("Logger")
%>