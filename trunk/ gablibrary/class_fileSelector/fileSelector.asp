<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		fileSelector
'' @CREATOR:		Michal Gabrukiewicz - gabru @ grafix.at
'' @CREATEDON:		2005-06-10 11:58
'' @CDESCRIPTION:	Lets you easily browse files, folders in a given folder and select any of them.
'' @VERSION:		1.0

'**************************************************************************************************************
class fileSelector

	'private members
	private p_sourcePath				
	private virtualSourcePath			
	private classLocation				
	private uniqueID					
	private selectedFiles				
	private output						
	private imagePath					
	private lastFileUnknownExtension	
	
	'public members
	public fso							''[object] filesystem-object. its public that the client can use it too.
	public defaultStyles				''[bool] load the default styles for the control? default = true
	public indentWidth					''[int] width of every indent in pixels. default = 15
	public filesSelectable				''[bool] are the files selectable? default = true
	public foldersSelectable			''[bool] are the folders selectable? default = false
	public height						''[int] height of the control in pixels. 
	public filesOpenable				''[bool] should the user be able to open the files on double-click? default = true
	public showFiles					''[bool] show the files within folder? useful when you want only the folders. default = true
	public useStringBuilder				''[bool] use the stringbuilder? DLL need to be installed. default = true
	public multipleSelection			''[bool] multiple selection is allowed? default = true
	public showExtensionOfKnownTypes	''[bool] show the extension of a known file-type. default = false
										''a file is known if the icon is available. just copy the icon to the images-directory
	public showTreeLines				''[bool] show the treelines? default = true
	public onItemClicked				''[string] The java script to execute when an item is clicked
	public name							''[string] name of the inputs checkbox/radio (used for postback). default = FSSelectedObject
	
	public property let selected(val) ''[array] array with full virtual-paths of the files which should be selected. If multipleSelection is off then the first item is taken. Note: Folder must always end with a "/"
		for i = 0 to uBound(val)
			pPath = server.mapPath(val(i))
			'check if the selected item is a folder.
			if str.endsWith(val(i), "/") then pPath = pPath & "\"
			selectedFiles.add val(i), pPath
		next
	end property
	
	public property get selected ''[array] gets the selected files. usefull after postback
		if request.form.count > 0 then
			selected = split(lib.RF(name), ", ")
		else
			selectedArray = array()
			for each key in selectedFiles.keys
				redim preserve selectedArray(uBound(selectedArray) + 1)
				selectedArray(uBound(selectedArray)) = key
			next
			selected = selectedArray
		end if
	end property
	
	public property let sourcePath(val) ''[string] virtual-path of the folder you want to browse
		virtualSourcePath = val
		if not (right(val, 1) = "/") then virtualSourcePath = virtualSourcePath & "/"
		p_sourcePath = server.mapPath(virtualSourcePath)
	end property
	
	private property get sourcePath
		sourcePath = p_sourcePath
	end property
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		p_sourcePath				= empty
		set fso						= server.createObject("scripting.fileSystemObject")
		set selectedFiles			= server.createObject("scripting.dictionary")
		defaultStyles 				= true
		classLocation				= consts.gabLibLocation & "class_fileSelector/" 'must start and end with a slash
		imagePath					= server.mappath(classLocation & "images")
		indentWidth					= 15
		filesSelectable				= true
		uniqueID					= 0
		height						= 200
		filesOpenable				= true
		useStringBuilder			= true
		multipleSelection			= true
		showExtensionOfKnownTypes	= false
		lastFileUnknownExtension	= false
		showTreeLines				= true
		foldersSelectable			= false
		showFiles					= true
		name						= "FSSelectedObject"
	end sub
	
	'**********************************************************************************************************
	'* desctructor 
	'**********************************************************************************************************
	private sub class_terminate()
		set fso	= nothing
		set selectedFiles = nothing
		set extensionIcons = nothing
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	draws the control
	'**********************************************************************************************************
	public sub draw()
		initStringBuilder()
		
		'we need to manipulate the selected-items if just one selection is allowed
		if not multipleSelection then
			for each key in selectedFiles.keys
				firstKey = key
				firstValue = selectedFiles(key)
				exit for
			next
			selectedFiles.removeAll()
			selectedFiles.add firstKey, firstValue
		end if
		
		if fso.folderExists(sourcePath) then
			if defaultStyles then print("<link rel=""stylesheet"" type=""text/css"" href=""" & classLocation & "std.css"">")
			'if treelines needed we add the vertical-treeline as background. not in css-file because
			'we want to use the classlocation
			if showTreeLines then
				print("<style>.fileSelector .vtl{background:url(" & classLocation & "images/tree_v.gif) repeat-y;}</style>")
			end if
			print("<script language=JavaScript src=" & classLocation & "javascript.js></script>")
			print("<div class=fileSelector style=""height:" & height & """>")
			
			set sourceFolder = fso.getFolder(sourcePath)
			drawTree sourceFolder, 0
			printFiles sourceFolder, 0
			print("</div>")
		else
			print(sourcePath & " does not exist.")
		end if
		if useStringBuilder then response.write(output.toString())
	end sub
	
	'**********************************************************************************************************
	'* initStringBuilder 
	'**********************************************************************************************************
	private function initStringBuilder()
		if useStringBuilder then
			Set output = Server.CreateObject("StringBuilderVB.StringBuilder")
			output.init 70000, 7500
		end if
	end function
	
	'**********************************************************************************************************
	'* print 
	'**********************************************************************************************************
	private sub print(outputStr)
		if useStringBuilder then
			output.append(outputStr)
		else
			response.write(outputStr)
		end if
	end sub
	
	'**********************************************************************************************************
	'* drawTree 
	'**********************************************************************************************************
	private sub drawTree(folder, depth)
		for each subFolder in folder.subfolders
			containerID = "c" & getUniqueID()
			
			'we ident every hierarchie
			if depth > 0 then
				print("<div class=""idt")
				if showTreeLines then print(" vtl")
				if depth > 1 or not showTreeLines then print(" idt2")
				print(""">")
			end if
			
			'check if we should open the folder or not. depends on selected files within the folder
			anyFilesSelected = false
			subFolderUCased = uCase(subFolder) & "\"
			for each field in selectedFiles.items
				if showFiles or (not showFiles and str.endsWith(field, "\")) then
					uCasedField = uCase(field)
					if str.startsWith(uCasedField, subFolderUCased) and not subFolderUCased = uCasedField then
						anyFilesSelected = true
						exit for
					end if
				end if
			next
			
			'print the current folder
			printFolder subFolder, containerID, anyFilesSelected, depth
			
			'print the whole container with all subfolders and files
			print("<div style=""display:")
			if anyFilesSelected then
				print("block")
			else
				print("none")
			end if
			print("""")
			print(" id=" & containerID)
			print(">")
			drawTree subFolder, depth + 1
			printFiles subFolder, depth + 1
			print("</div>")
			
			if depth > 0 then print("</div>")
		next
	end sub
	
	'**********************************************************************************************************
	'* printFiles 
	'**********************************************************************************************************
	private sub printFiles(folder, depth)
		if not showFiles then exit sub
		print("<div")
		if depth > 0 then
			print(" class=""idt")
			if showTreeLines then print(" vtl")
			if depth > 1 or not showTreeLines then print(" idt2")
			print("""")
		end if
		print(">")
		for each file in folder.files
			printFile file, depth
		next
		print("</div>")
	end sub
	
	'**********************************************************************************************************
	'* printFolder 
	'**********************************************************************************************************
	private sub printFolder(folder, folderID, opened, depth)
		folderIMG = "folder.gif"
		if opened then folderIMG = "folder_open.gif"
		print("<div class=folder>")
		if depth > 0 then printHorizontalTreeLine()
		
		if foldersSelectable then
			objectLocation = physicalToVirtualPath(folder.path) & "/"
			objectSelected = false
			objectSelected = selectedFiles.exists(objectLocation)
			print("<input onclick=""" & onItemClicked & """")
			print(" type=")
			print(lib.iif(multipleSelection, "checkbox", "radio"))
			print(" value=""" & objectLocation & """")
			print(" name=" & name)
			if objectSelected then print(" checked")
			print(">")
		end if
		print("<a href=#>")
		print("<span onclick=""FSToggleFolder('" & folderID & "',this)""><img")
		print(" src=" & classLocation & "images/" & folderIMG)
		print(" id=img" & folderID)
		print(">")
		print(folder.name)
		print("</a></span>")
		print("</div>")
	end sub
	
	'**********************************************************************************************************
	'* physicalToVirtualPath 
	'* Example: D:\xx\xx\ => /xx/xx/
	'**********************************************************************************************************
	private function physicalToVirtualPath(aPath)
		physicalToVirtualPath = virtualSourcePath & mid(replace(aPath, "\", "/"), len(sourcePath) + 2, len(aPath))
	end function
	
	'**********************************************************************************************************
	'* printFile 
	'**********************************************************************************************************
	private sub printFile(file, depth)
		fileID = "f" & getUniqueID()
		fileLocation = physicalToVirtualPath(file.path)
		
		print("<div class=file>")
		if depth > 0 then printHorizontalTreeLine()
		
		fileSelected = false
		if filesSelectable then
			fileSelected = selectedFiles.exists(fileLocation)
			print("<input onclick=""FSselect('lb" & fileID & "',this);" & onItemClicked & """")
			print(" id=" & fileID)
			print(" type=")
			if multipleSelection then
				print("checkbox")
			else
				print("radio")
				print(" ondblclick=FSuc(this)")
			end if
			print(" value=""" & fileLocation & """")
			print(" name=" & name)
			if fileSelected then print(" checked")
			print(">")
			'if its a single selection we need to store the selected label for the javascript.
			'so it can be unselected if another file will be selected
			if fileSelected and not multipleSelection then print("<script>FSSelectedFileID='lb" & fileID & "'</script>")
		end if
		
		print("<label id=lb" & fileID)
		print(" for=" & fileID)
		if filesOpenable then print(" ondblclick=""window.open('" & fileLocation & "')""")
		print(" title=""" & file.name & " (" & byteToKiloByte(file.size) & " KB)""")
		if fileSelected then print(" class=s")
		print(">")
		
		print(getExtensionIcon(file))
		
		if showExtensionOfKnownTypes then
			print(file.name)
		else
			print(fso.getBaseName(file))
		end if
		print("</label>")
		print("</div>")
	end sub
	
	'**********************************************************************************************************
	'* printHorizontalTreeLine 
	'**********************************************************************************************************
	private function printHorizontalTreeLine()
		if showTreeLines then print("<img src=" & classLocation & "images/tree_h.gif>")
	end function
	
	'**********************************************************************************************************
	'* byteToKiloByte 
	'**********************************************************************************************************
	private function byteToKiloByte(bytes)
		byteToKiloByte = formatnumber(bytes / 1024)
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	gets an HTML-string with the appropriate icon for the filetype
	'' @PARAM:			file [file] file from FSO
	'' @RETURN:			[string] <img>-tag 
	'**********************************************************************************************************
	public function getExtensionIcon(file)
		getExtensionIcon = lib.getFileIcon(file)
	end function
	
	'**********************************************************************************************************
	'* getUniqueID 
	'**********************************************************************************************************
	private function getUniqueID()
		getUniqueID = uniqueID
		uniqueID = uniqueID + 1
	end function

end class
lib.registerClass("FileSelector")
%>