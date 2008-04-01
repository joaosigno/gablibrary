<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		TableRow
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		2006-01-10 18:25
'' @CDESCRIPTION:	Represents a complete datarow of the Drawtable. it is drawn as a TR-element of a table
'' @REQUIRES:		-
'' @VERSION:		0.1

'**************************************************************************************************************
class TableRow

	'public members
	public index				''[int] index (position) of the row in the table. starts with 0
	public ID					''[string] ID-attribute of the TR-element
	public recordID				''[string] the ID of the current data-record. depends on the fieldlinkID of the Drawtable
	public hoverEffect			''[bool] hovereffect (on mouseover) turned on or not? default taken from Drawtable setting
	public CSSClass				''[string] CSS-Class of the row (TR-element).
	public attributes			''[string] additional attributes inside the TR-element
	public BGColor				''[string] background-color of the row
	public disabled				''[bool] should the row be disabled? clicking then is not allow, radiobuttons are disabled, etc.
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		index = 0
		ID = empty
		recordID = empty
		hoverEffect = false
		CSSClass = empty
		attributes = empty
		BGColor = empty
		disabled = false
	end sub

end class
%>