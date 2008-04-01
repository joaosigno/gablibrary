<!--#include file="config.asp"-->
<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003 - This file is part of GAB_LIBRARY		
'* For license refer to the license.txt in the root    						
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		DocumentHolder
'' @CREATOR:		Michael Rebec, Michal Gabrukiewicz
'' @CREATEDON:		2006-09-28 16:28
'' @CDESCRIPTION:	A control which provides the interface for uploading/deleting/viewing a file
'' @REQUIRES:		fileUpload
'' @VERSION:		0.1

'**************************************************************************************************************
class DocumentHolder

	private fso
	private successfullyPerformed 'stores if the perform was successfull or not.
	
	private property get CBName
		CBName = name & "CB"
	end property
	
	private property get valueFull 'full path of the value.
		valueFull = uploader.uploadPath & value
	end property
	
	public name					''[string] the name of the control
	public value				''[string] name of the file which will be displayed. path is taken from the uploader. e.g. hugo.jpg
	public deleteCaption 		''[string] the description of the text standing next to the checkbox (e.g. Delete)
	public cssLocation			''[string] location of the stylesheet file. by default it is taken from the config.asp
	public autoHandleFileTypes 	''[bool] a value indicating wheter the displaying of the file types should be handled automatic.
								''e.g. Images will be displayed in a nice box, text files will be opened in a new window, ...
								''if false all files will be displayed in a new browser window
	public uploader				''[fileUpload] the fileUploader associated with the documentHolder. set the properties
								''like targetpath, etc. in order to define how the upload will be handled.
	public onDelete				''[string] name of the SUB which will be executed when the user wants to delete the file.
								''e.g. you want to update your database
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		set uploader = new FileUpload
		name = "documentHolder"
		deleteCaption = "Delete"
		value = empty
		onDelete = empty
		cssLocation = DH_CSS_CLASS_LOCATION
		autoHandleFileTypes = true
		set fso = server.createObject("Scripting.FileSystemObject")
		successfullyPerformed = false
	end sub
	
	'**********************************************************************************************************
	'* destructor 
	'**********************************************************************************************************
	public sub class_terminate()
		set fso = nothing
		set uploader = nothing
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	Draws the control
	'**********************************************************************************************************
	public sub draw()
		loadScripts()
		with str
			.writeln("<div class=""docHolder"">")
			if existsFile() then
				str.writeln("<div>")
				if autoHandleFileTypes then
					ext = lCase(getExtension())
					'detected images
					if ext = "jpeg" or ext = "jpg" or ext = "gif" or ext = "png" then attr = " rel=""lightbox"""
				end if
				.writeln("<a href=""" & valueFull & """ target=""_blank""" & attr & ">")
				.writeln(getFileIcon())
				.writeln("View</a>")
				'when the perform was successful then we dont want to tick the delete box anymore
				checked = lib.iif(lib.RFHas(CBName) and not successfullyPerformed, "checked", empty)
				.writeln("&nbsp;<input type=""checkbox"" onclick=""" & name & ".disabled=this.checked"" name=""" & CBName & """ id=""" & CBName & """ value=""1"" " & checked & ">")
				.writeln("<label for=""" & CBName & """>" & deleteCaption & "</label>")
				.writeln("</div>")
			end if
			disabled = lib.iif(checked <> "", " disabled", empty)
			.writeln("<div><input type=""file"" size=25 name=""" & name & """ id=""" & name & """ class=""fileBrowser""" & disabled & "></div>")
			.writeln("</div>")
		end with
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	performs the actions of the control. delete, upload, etc.
	'' @DESCRIPTION:	1 file will be uploaded if selected
	''					2. if no file selected and "delete"-ticked then onDelete event will be raised.
	''					- it deletes the file on the filesystem as well
	'' @RETURN:			[bool] true if successfull. otherwise false, get the error from the uploader then.
	'**********************************************************************************************************
	public function perform()
		successfullyPerformed = false
		perform = true
		if lib.RFHas(name) then
			uploader.fieldName = name
			uploader.overwrite = true
			success = uploader.upload()
			if success <> "" then
				value = success
			else
				perform = false
			end if
		elseif lib.RFHas(CBName) then
			if onDelete <> "" then execute(onDelete)
			'delete from filesystem
			if existsFile() then fso.deleteFile server.mappath(valueFull), true
			'now we can reset the value. because it has been deleted
			value = ""
			perform = true
		end if
		successfullyPerformed = perform
	end function
	
	'**********************************************************************************************************
	'* getFileIcon - gets the corresponding file icon associated with the file
	'**********************************************************************************************************
	private function getFileIcon()
		if existsFile() then
			set file = fso.getFile(server.mappath(valueFull))
			getFileIcon = lib.getFileIcon(file)
		end if
	end function
	
	'**********************************************************************************************************
	'* existsFile - checks if the file specified in the filename exists in the directory
	'**********************************************************************************************************
	private function existsFile()
		existsFile = fso.fileExists(server.mappath(valueFull))
	end function
	
	'**********************************************************************************************************
	'* getExtension - gets the extension of the file
	'**********************************************************************************************************
	private function getExtension()
		if existsFile() then getExtension = fso.getExtensionName(server.mappath(valueFull))
	end function
	
	'**********************************************************************************************************
	'* loadScripts - loads the needed stylesheet and javascript files
	'**********************************************************************************************************
	private sub loadScripts()
		lib.execJS("var loadingImage = '" & DH_CLASS_LOCATION & "lightbox/loading.gif';")
		lib.execJS("var closeButton = '" & DH_CLASS_LOCATION & "lightbox/close.gif';")
		str.writeln("<script type=""text/javascript"" src=""" & DH_CLASS_LOCATION & "lightbox/lightbox.js""></script>")
		str.writeln("<link rel=""stylesheet"" type=""text/css"" href=""" & DH_CLASS_LOCATION & "lightbox/lightbox.css"">")
		str.writeln("<link rel=""stylesheet"" href=""" & cssLocation & """>")
	end sub

end class
lib.registerClass("DocumentHolder")
%>