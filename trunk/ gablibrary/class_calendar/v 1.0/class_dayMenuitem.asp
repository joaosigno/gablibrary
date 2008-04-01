<%
'**************************************************************************************************************

'' @CLASSTITLE:		dayMenuitem
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		09.08.2004
'' @CDESCRIPTION:	Represents a menuitem for the daymenu. used in the calendar-control
'' @VERSION:		1.0

'**************************************************************************************************************

class dayMenuitem

	public caption					''[string] caption of the menuitem
	public toolTip					''[string] tooltip
	public onClick					''[string] what should be executed onclick
	public disabled					''[bool] disabled or not. default=false
	public hoverEffect				''[bool] enable/disable mousehover. default=true
	
	'Konstruktor => set the default values
	private sub Class_Initialize()
		caption						= empty
		toolTip						= empty
		onClick						= empty
		disabled					= empty
		hoverEffect					= true
	end sub
	
	'********************************************************************************************************
	'* @SDESCRIPTION:		draws the menuitem
	'********************************************************************************************************
	public sub draw()
		strDisabled = lib.iif(disabled, " disabled", empty)
		with str
			.writeln("<div" & strDisabled & " " &_
						"class=dayMenuItem " & getHoverEvents() &_
						"onClick=""" & onClick & """ " &_
						"title=""" & toolTip & """>")
			.writeln(caption)
			.writeln("</div>")
		end with
	end sub
	
	'********************************************************************************************************
	'* getHoverEvents 
	'********************************************************************************************************
	private function getHoverEvents()
		if hoverEffect then
			onMouseover = "changeCssClass('dayMenuItem dayMenuItemHover', this)"
			onMouseout = "changeCssClass('dayMenuItem', this)"
			getHoverEvents = " onmouseover=""" & onMouseover & """onmouseout=""" & onMouseout & """ "
		end if
	end function

end class
%>