<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		TreeviewGroup
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		2005-09-22 11:28
'' @CDESCRIPTION:	Represents a group for the treeview-control
'' @VERSION:		0.1

'**************************************************************************************************************
class TreeviewGroup

	'public members
	public IDFieldname			''[string] name of the field which holds the ID of the group
	public textFieldname		''[string] name of the field which holds the text of the group
	public treeview				''[dropdown] instance of the treeview the group belongs to
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	private sub class_initialize()
		IDFieldname = empty
		textFieldname = empty
		set treeview = nothing
	end sub

end class
%>