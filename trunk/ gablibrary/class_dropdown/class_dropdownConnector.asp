<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		DropdownConnector
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		2005-11-24 15:26
'' @CDESCRIPTION:	Connects two dropdowns together so you can move items from one to another, etc.
'' @REQUIRES:		-
'' @VERSION:		0.1
'' @FRIENDOF:		Dropdown
'' @TODO:			styles beim moven und sortieren &uuml;bernehmen. eigentlich alle attribute. (copy-object)
''					gehen n&auml;mlich verloren wenn man ein item verschiebt.

'**************************************************************************************************************
class DropdownConnector

	'private members
	private p_sort
	
	'public members
	public source			''[Dropdown]
	public target			''[Dropdown]
	public page				''[Generatepage] instance of the currentpage
	public allowMoveAll		''[bool] defualt = true
	public allowMoveSingle	''[bool] default = true
	public allowSwap		''[bool] default = true
	public captionSource	''[string] caption which will be displayed over the source dropdown
	public captionTarget	''[string] caption which will be displayed over the target dropdown
	
	public property let sort(value) ''[string] set the sorting. ASC = ascending, DESC = descending, empty = no sorting
		select case uCase(value)
			case "ASC"
				p_sort = 1
			case "DESC"
				p_sort = 2
			case else
				p_sort = 0
		end select
	end property
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		allowMoveAll = true
		allowMoveSingle = true
		allowSwap = true
		p_sort = 1
	end sub
	
	'**********************************************************************************************************
	'* destructor 
	'**********************************************************************************************************
	private sub class_terminate()
		set source = nothing
		set target = nothing
		set page = nothing
	end sub
	
	'**********************************************************************************************************
	'' @DESCRIPTION:	draws the panel including the connected dropdowns
	'**********************************************************************************************************
	public sub draw(size, width)
		load(me.page)
		with str
			.writeln("<table class=dropdownConnector cellpadding=0 cellspacing=0 width=" & width & ">")
			.writeln("<tr align=center><th><label for=" & source.ID & ">" & captionSource & "</label></th>")
			.writeln("<th>&nbsp;</th>")
			.writeln("<th><label for=" & target.ID & ">" & captionTarget & "</label></th></tr>")
			.writeln("<tr>")
			.writeln("<td width=50% >")
			addJSEventHandler(source)
			source.size = size
			source.draw()
			.writeln("</td>")
			.writeln("<td align=center>")
			drawPanel()
			.writeln("</td>")
			.writeln("<td width=50% >")
			addJSEventHandler(target)
			target.size = size
			target.draw()
			.writeln("</td>")
			.writeln("</tr>")
			.writeln("</table>")
		end with
	end sub
	
	'**********************************************************************************************************
	'' @DESCRIPTION:	draws a control-panel of the connector. buttons to move items...
	'**********************************************************************************************************
	public sub drawPanel()
		str.writeln("<div class=dropdownConnector>")
		if allowMoveAll then
			drawMoverButton "(new Dropdown('" & source.ID & "', " & p_sort & ")).moveAll('" & target.ID & "')", "icon_doublearrow_right.gif"
			str.write("<br>")
		end if
		if allowMoveSingle then
			drawMoverButton "(new Dropdown('" & source.ID & "', " & p_sort & ")).moveOptions('" & target.ID & "')", "icon_arrow_right.gif"
			str.write("<br>")
		end if
		if allowSwap then
			drawMoverButton "(new Dropdown('" & source.ID & "', " & p_sort & ")).swapOptions('" & target.ID & "')", "icon_arrow_leftright.gif"
			str.write("<br>")
		end if
		if allowMoveSingle then
			drawMoverButton "(new Dropdown('" & target.ID & "', " & p_sort & ")).moveOptions('" & source.ID & "')", "icon_arrow_left.gif"
			str.write("<br>")
		end if
		if allowMoveAll then
			drawMoverButton "(new Dropdown('" & target.ID & "', " & p_sort & ")).moveAll('" & source.ID & "')", "icon_doublearrow_left.gif"
		end if
		str.writeln("</div>")
	end sub
	
	'**********************************************************************************************************
	'* drawButton 
	'**********************************************************************************************************
	private sub drawMoverButton(onclick, imgFilename)
		str.write("<button onclick=""" & onclick & """>" & getImg(imgFilename) & "</button>")
	end sub
	
	'**********************************************************************************************************
	'* getImg 
	'**********************************************************************************************************
	private function getImg(filename)
		getImg = "<img src=" & source.controlLocation & "images/" & filename & " border=0>"
	end function
	
	'**********************************************************************************************************
	'* addJSEventHandler 
	'**********************************************************************************************************
	private sub addJSEventHandler(byRef DD)
		if allowMoveSingle then
			DD.attributes = DD.attributes & _
							" ondblclick=""(new Dropdown('" & DD.ID & "', " & p_sort & ")).moveOptions('" & getTargetID(DD) & "')""" & _
							" onkeyup=""if(event.keyCode==13)(new Dropdown('" & DD.ID & "', " & p_sort & ")).moveOptions('" & getTargetID(DD) & "')"""
		end if
	end sub
	
	'**********************************************************************************************************
	'' @DESCRIPTION:	loads all the needed javascript for moving, etc. and the stylesheet
	'' @PARAM:			currentPage [GeneratePage] instance of the current-page
	'**********************************************************************************************************
	public sub load(currentPage)
		currentPage.loadJavascriptFile(source.controlLocation & "dropdownConnector.js")
		currentPage.loadStylesheetFile source.controlLocation & "dropdownConnector.css", empty
	end sub
	
	'**********************************************************************************************************
	'* getTargetID 
	'**********************************************************************************************************
	private function getTargetID(sourceDD)
		if lCase(sourceDD.ID) = lCase(source.ID) then
			getTargetID = target.ID
		else
			getTargetID = source.ID
		end if
	end function

end class
%>