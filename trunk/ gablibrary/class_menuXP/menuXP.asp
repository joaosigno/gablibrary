<!--#include virtual="/gab_Library/class_sort/sort.asp"-->
<!--#include file="menuPointXP.asp"-->
<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'*****************************************************************************************************************************************

'' @CLASSTITLE:		menuXP
'' @CREATOR:		Basic Idea from Timothy Marin - http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=8167&lngWId=4
''					Rewritten to class & modified by Michael Rebec
'' @CREATEDON:		24.11.2003
'' @CDESCRIPTION:	Draws a menu in windows XP - style. Several skins available
''					With this class you can easily create a menu dynamically and completely OOP with ASP
'' @VERSION:		0.3.1

'*****************************************************************************************************************************************
class menuXP
	
	private p_menu				'Dictionary Item Object for drawing
	private p_menuPoints		'Dictionary Object 
	private p_Counter			'counter for the Dictionary Object
	private p_sortIt			'for the sorting
	private p_sortObj			'the sort object
	private p_menuStyle			'add your own style
	private p_menuKey			'checks request.querystring("CurrentMenuKey")
	private p_openKey			'the current Menu index
	private p_keyCounter		'counts the Menu index
	private p_skin				'skin control
	
	public defaultStylesheet	''[bool] - Use default styles: default=true?
	public closeable			''[bool] - Is it possible to close the MainMenu Points: default=true
	public mainPointAlign
	
	'Konstruktor => set the default values
	private sub Class_Initialize()
		set p_menuPoints 	= Server.createObject("Scripting.Dictionary")
		set p_sortObj	 	= new Sort
		p_menuKey			= request.querystring("CurrentMenuKey")
		p_openKey			= request.querystring("CurrentOpenKey")
		p_Counter 			= 0
		p_keyCounter		= 0
		p_sortIt  			= false
		p_menuStyle 		= empty
		p_skin				= 0
		defaultStylesheet 	= true
		closeable			= true
		mainPointAlign		= "center"
	end sub
	
	'Destruktor	=> close Dictionary
	private sub Class_Terminate()
		set p_menuPoints = nothing
		set p_sortObj	 = nothing
	end sub
	
	public property get sorting ''[bool] is sorting on ?
		sorting = p_sortIt
	end property
	public property let sorting(value) ''[bool] set if sorting is on
		p_sortIt = value
	end property
	
	public property let menuStyle(value) ''[string] add your own styles to the menu. It will be placed into the Table style="..."
		p_menuStyle = value
	end property
	
	public property let skin(value) ''[string] menu style
		Select Case UCase(value)
			Case "CLASSIC":		p_skin = 0
			Case "GREEN":		p_skin = 1
			Case "FLAT":		p_skin = 2
			Case Else:			p_skin = 0
		End Select
	end property
	
	'**************************************************************************************************************
	'* load skin
	'**************************************************************************************************************
	private function loadSkin()
		if p_skin = 1 then
			skinName = "green"
		elseif p_skin = 2 then
			skinName = "flat"
		elseif defaultStylesheet or (p_skin=0) then
			skinName = "classic"
		end if
		lib.page.loadStylesheetFile consts.gablibLocation & "class_menuXP/skins/" & skinName & ".css", empty
	end function
	
	'**************************************************************************************************************
	' draw the menu points
	'**************************************************************************************************************
	private sub drawMenuPoints()
		if p_menu.parent = "" then
			p_keyCounter = p_keyCounter + 1
			drawMainMenuPoint()
		elseif p_menu.parent = p_menuKey then
			openMenuPoint()
		end if
	end sub
	
	'**************************************************************************************************************
	' draw a main menu point
	'**************************************************************************************************************
	private sub drawMainMenuPoint()
		println("<tr><td><table class=""XPMenutable"" cellSpacing=0 cellPadding=0 width=""100%"">")
		println("<tr><td align=""" & MainPointAlign & """ nowrap class=""MainMenu"">")
		if not p_menu.link = empty then
			println("<a class=Main href=""" & p_menu.link &""" target="""& p_menu.target &""">")
		else
			if (CInt(p_openKey) = CInt(p_keyCounter)) And closeable then
				println("<a class=Main href="""& request.servervariables("Script_name") &""" target="""& p_menu.target &""">")
			else
				println("<a class=Main href=""?CurrentMenuKey=" & p_menu.title & "&CurrentOpenKey="& p_keyCounter &""" target="""& p_menu.target &""">")
			end if
		end if
		if not p_menu.image = empty then
			println("<img style=""VERTICAL-ALIGN: middle"" height=16 src="""& p_menu.image &""" width=16 border=0>&nbsp;&nbsp;")
		end if
		if p_menuKey = p_menu.title then
			activeOrNotClass = " class=mainMenuPointActive"
		else
			activeOrNotClass = empty
		end if
		println("<strong" & activeOrNotClass & ">" & p_menu.title & "</strong></A>")
		println("</td></tr></table></td></tr>")
	end sub
	
	'**************************************************************************************************************
	' draw a sub menu point
	'**************************************************************************************************************
	private sub openMenuPoint()
		println("<tr><td align=left valign=middle height=21>")
		println("<a class=""menu"" href="""& p_menu.link &""" target="""& p_menu.target &""">")
		if not p_menu.image = "" then
			println("&nbsp;&nbsp;<img align=absmiddle height=16 width=16 border=0 src="""& p_menu.image &""">&nbsp;&nbsp;")
		end if
		println(p_menu.title &"</a></td></tr>")	
	end sub
	
	'**************************************************************************************************************
	' draw the string
	'**************************************************************************************************************
	private sub println(myStr)
		str.writeln(myStr)
	end sub
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	This method draws the Menu
	'**************************************************************************************************************
	public sub draw()
		loadSkin()
		println("<table class=""XPMenutable"" style=""width:100%"& p_menuStyle &""" cellSpacing=0 cellPadding=0>")
		if p_sortIt then
			set p_Points = p_sortObj.sortdict(p_menuPoints, "parent, title", "Main")
			for i = 0 to p_Points.count - 1
				set p_menu = p_Points.Item(i)
				call drawMenuPoints()
			next
		else
			for each p_menu in p_menuPoints.Items
				call drawMenuPoints()
			next
		end if
		str.write("</table>")
	end sub
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Add`s a menu item.
	'' @PARAM:			- [MenuPoint] the MenuItem you want to add
	'**************************************************************************************************************
	public sub addMenuItem(myObject)
		p_menuPoints.Add p_Counter, myObject
		p_Counter = p_Counter + 1
	end sub

end class
lib.registerClass("menuXP")
%>