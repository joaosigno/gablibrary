<!--#include virtual="/gab_Library/class_page/generatepage.asp"-->
<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003 - This file is part of GAB_LIBRARY		
'* For license refer to the license.txt in the root    						
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		TestFixture
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		2007-04-03 16:18
'' @CDESCRIPTION:	Represents a Test Fixture which is used for creating tests and running them.
''					It derives from the page because tests can be run and therefore return output (xml) about the
''					execution itself. Test fixtures must have the fileextension .test
'' @REQUIRES:		-
'' @VERSION:		0.1

'**************************************************************************************************************
class TestFixture

	'private members
	private myBase
	
	'public members
	public name
	public description
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		set myBase = new GeneratePage
		myBase.isXML = true
		myBase.onlyWebDev = true
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	
	'' @DESCRIPTION:	
	'' @PARAM:			name [type]: the text you want to log
	'' @RETURN:			[type] 
	'**********************************************************************************************************
	public sub draw()
		
	end sub

end class
%>