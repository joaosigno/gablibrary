<%
'**************************************************************************************************************

'' @CLASSTITLE:		TableLegend
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		2006-02-05 10:16
'' @CDESCRIPTION:	A legend for the drawtable
'' @VERSION:		0.1

'**************************************************************************************************************
class TableLegend

	'private members
	private items
	
	'public members
	public text				''[string] some free text for the legend
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	private sub class_initialize()
		set items = server.createObject("scripting.dictionary")
	end sub
	
	'**********************************************************************************************************
	'* destructor 
	'**********************************************************************************************************
	private sub class_terminate()
		set items = nothing
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	adds a new legend item
	'' @PARAM:			hint [string]: can be whether a cssclass-name or an image.
	''					if its a cssclass then a box will be rendered with the appereance of the cssclass
	''					else an image will be rendered
	'' @RETURN:			[type] 
	'**********************************************************************************************************
	public sub addItem(hint, description)
		items.add hint, description
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	gets the HTML as string
	'' @PARAM: 			legendCaption [string]: caption of the legend headline. 
	'**********************************************************************************************************
	public function toString(legendCaption)
		r = empty
		r = r & "<div id=drawTableLegendContainer>"
		r = r & "<div class=legendHL><a href=""javascript:dtToggleVisibility('drawTableLegend')"">" & legendCaption & "</a></div>"
		r = r & "<div id=drawTableLegend style=""display:none""><table cellspacing=0 cellpadding=3><tr><td id=dtlText colspan=2>" & text & "</td></tr>"
		for each key in items.keys
			if str.endsWith(uCase(key), ".GIF") then
				hint = "<img src=""" & key & """ border=0 align=absmiddle>"
			else
				hint = "<div class=""legendBox " & key & """ align=center>SAMPLE</div>"
			end if
			r = r & "<tr valign=top><td width=0% >" & hint & "<td></td><td>" & items(key) & "</td></tr>"
		next
		r = r & "</table></div></div>"
		toString = r
	end function

end class
%>