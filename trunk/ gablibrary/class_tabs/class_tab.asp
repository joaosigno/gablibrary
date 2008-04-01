<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		tab
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		03.09.2003
'' @CDESCRIPTION:	Needed for tabs-class. This object is one tab
'' @VERSION:		0.1

'**************************************************************************************************************
class tab

	'private members
	private p_attributes
	private p_onClick
	
	'public members
	public caption				''[string] Value to display on the Tab.
	public procedure			''[string] The name of the procedure which should be called when clicking on the tab
	public show					''[bool] do you want to show the tab? default = true
	public disabled				''[bool] indicates wheater the tab is disabled or not. default = false
	public alwaysClickable		''[bool] if true the selected tab will not be disabled. default = false
	
	private passed_querystring	'The querystring passed from the tab-class
	
	'Konstruktor => set the default values
	private sub Class_Initialize()
		caption				= empty
		procedure			= empty
		disabled			= false
		passed_querystring 	= empty
		p_attributes		= empty
		p_onClick 			= empty
		show				= true
		alwaysClickable		= false
	end sub
	
	public property let onClick(val) ''[string] sets the onclick attribute. if you overwrite the value then auto-postback is disabled.
		p_onClick = val
	end property
	
	private property get onClick()
		onClick = p_onClick
	end property
	
	public property let attributes(val) ''[string] sets the attributes of the tab (button-tag)
		p_attributes = val
	end property
	
	public property get attributes() ''[string] gets the attributes of the tab
		if p_attributes <> "" then
			attributes =  " " & p_attributes
		else
			attributes = empty
		end if
	end property
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Draws one tab 
	'' @PARAM:			- active [bool]: Active tab or not
	'' @PARAM:			- ID [int]: id of the tab
	'' @PARAM:			- attribute [string]
	'' @PARAM:			- QSnotUse [string]
	'******************************************************************************************************************
	public sub draw(active, ID, attribute, QSnotUse)
		if show then
			if active or disabled then
				if active then
					myClass = "tab_active"
				end if
				if not alwaysClickable then
					disabled = " disabled"
				else
					disabled = ""
				end if
			else
				myClass = "tab_inactive"
				disabled = empty
			end if
			
			dontShow = "activeTab"
			if not QSnotUse = empty then
				dontShow = dontShow & "," & QSnotUse
			end if
			
			changed = lib.getAllFromQueryStringBut(dontShow)
			if changed <> empty then
				passed_querystring = "&" & changed
			end if
			
			if onClick = "" then onClick = "window.location.href='" & request.servervariables("URL") & "?activeTab=" & ID & passed_querystring & "'"
			%><!--#include file="display_tab.asp"--><%
		end if
	end sub

end class
%>