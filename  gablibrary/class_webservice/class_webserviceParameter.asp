<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003 - This file is part of GAB_LIBRARY		
'* For license refer to the license.txt in the root    						
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		WebserviceParameter
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		2007-02-22 14:31
'' @CDESCRIPTION:	Represents a parameter for a gabLibrary Webservice
'' @REQUIRES:		-
'' @FRIENDOF:		Webservice
'' @VERSION:		0.1

'**************************************************************************************************************
class WebserviceParameter

	'private members
	private p_dataType
	
	'public members
	public name				''[string] name of the parameter
	public defaultValue		''[string] the default value which will be used if its an optional parameter
	public description		''[string] verbal description of the parameter
	
	public property get dataType ''[string] gets the type of the parameter. supported: bool, string, int. default = string
		dataType = uCase(p_dataType)
	end property
	
	public property let dataType(val) ''[string] sets the type of the parameter.
		p_dataType = val
	end property
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		name = ""
		dataType = "STRING"
	end sub

end class
%>