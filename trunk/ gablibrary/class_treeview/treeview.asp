<!--#include virtual="/gab_Library/class_treeview/class_treeviewGroup.asp"-->
<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		Treeview
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		2005-09-22 11:04
'' @CDESCRIPTION:	NOT WORKING YET!!!! UNFINISHED! Generate a treeview using a Recordset.
''					naming, etc. taken most from the .NET-treeview-control
'' @VERSION:		0.1

'**************************************************************************************************************
class Treeview

	'private members
	private classLocation
	private groupID					'this ID is used to number the groups in the groups-collection
	
	'public members
	public datasource				''[ADODB-Recordset] your data
	public defaultStyles			''[bool] load the default styles for the control? default = true
	public groups					''[Dictionary] the groups collection
	public ID						''[string] ID of the control
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		groupID 			= 0
		classLocation		= consts.gabLibLocation & "class_fileSelector/" 'must start and end with a slash
		set datasource 		= nothing
		set groups			= server.createObject("scripting.dictionary")
	end sub
	
	'**********************************************************************************************************
	'* destructor 
	'**********************************************************************************************************
	private sub class_terminate()
		set datasource = nothing
		set groups = nothing
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	adds a group. need to be done before draw()-ing
	'' @DESCRIPTION:	the last added group is always the leaf!
	'' @PARAM:			groupObj [TreeviewGroup]: create a new object or use the getNewGroup()-method
	''					for more readable code.
	'**********************************************************************************************************
	public sub addGroup(groupObj)
		groups.add groupID, groupObj
		groupID = groupID + 1
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	gets a new Group with the wanted parameters. useful as input for the addGroup()-method
	'**********************************************************************************************************
	public function getNewGroup(IDFieldname, textFieldname)
		set getNewGroup = new TreeviewGroup
		getNewGroup.IDFieldname = IDFieldname
		getNewGroup.textFieldname = textFieldname
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	draws the control
	'**********************************************************************************************************
	public sub draw()
		if defaultStyles then print("<link rel=""stylesheet"" type=""text/css"" href=""" & classLocation & "std.css"">")
		
		print("<div id=" & ID & ">")
		
		if not datasource.eof then
			currentPath = empty
			lastPath = empty
			depth = 0
			lastID = "-1"
			while not datasource.eof
				
				currentPath = datasource(groups(0).IDFieldname)
				diff = replace(lastPath, currentPath, ":")
				if instr(diff, ":") > 0 then depth =  depth - 1
				valueField = groups(depth).IDFieldname
				currentID = datasource(valueField)
				textField = groups(depth).textFieldname
				
				if cint(lastID) <> cint(currentID) then depth = depth + 1
				
				print("<div>" & str.clone("&nbsp;&nbsp;&nbsp;", depth) & datasource(textField) & " (" & currentPath & ")</div>")
				
				lastPath = currentPath
				lastID = currentID
				datasource.movenext()
			wend
		end if
		
		print("</div>")
	end sub
	
	'**********************************************************************************************************
	'' print 
	'**********************************************************************************************************
	private sub print(value)
		str.write(value)
	end sub

end class
%>