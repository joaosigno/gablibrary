<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		columnFilter
'' @CREATOR:		Michael Rebec
'' @CREATEDON:		24.11.2003
'' @CDESCRIPTION:	Needed for menuXP Class
'' @VERSION:		0.1

'**************************************************************************************************************
class MenuPoint

	public parent		'' [string] the parent-menu-item. Leave it blank if this is a Main-Menu Point, otherwise enter the title of the Main-Menu Point
	public title		'' [string] the display string of the Menu Point
	public link			'' [string] the URL of the target file
	public image		'' [string] the URL of the icon
	public target		'' [string] the name of the frame in which you want to display the site, or "_blank" for a new page
	
	'Konstruktor => set the default values
	private sub Class_Initialize()
		parent 		= empty
		title  		= empty
		link		= empty
		image		= empty
		target		= empty
	end sub

end class
%>