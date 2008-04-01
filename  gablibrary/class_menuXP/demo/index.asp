<!--#include virtual="/gab_library/class_page/generatepage.asp"-->
<!--#include virtual="/gab_library/class_menuXP/menuXP.asp"-->
<%
dim menu
set page = new generatePage
with page
	.onlyWebDev = true
	.showFooter		= false
	.draw()
end with
set page = nothing

'******************************************************************************************
'* main - we start . DONT FORGET TO KILL ALL OBJECTS!! 
'******************************************************************************************
sub main()
	response.write "<TABLE border=0 cellpadding=0 cellspacing=0 width=100% height=100% style=""border-right:2px dotted #DBD8D1""><TR><TD valign=top>"
	set menu = new menuXP
	with menu
		.sorting = true
		.skin = "Classic"
		set Main_Info = new MenuPoint
		Main_Info.title	= "Diverse Information"
		.addMenuItem(Main_Info)
		
		set MenuItem = new MenuPoint
		with MenuItem
			.parent	= Main_Info.title
			.title	= "Intranet CSS-Concept"
			.link	= "/intranet_style/sample style page.asp"
			.target = "main"
		end with
		.addMenuItem(MenuItem)
		set MenuItem = nothing
		
		set MenuItem = new MenuPoint
		with MenuItem
			.parent	= Main_Info.title
			.title	= "Miscellaneous Informations"
			.link	= "../informations/"
			.target = "main"
		end with
		.addMenuItem(MenuItem)
		set MenuItem = nothing
		
		set MenuItem = new MenuPoint
		with MenuItem
			.parent	= Main_Info.title
			.title	= "IIS 6 Guide"
			.link	= "../references/eBook.Sams.-.Microsoft.IIS.6.Delta.Guide.chm"
			.target = "main"
		end with
		.addMenuItem(MenuItem)
		set MenuItem = nothing
		
		set MenuItem = new MenuPoint
		with MenuItem
			.parent	= Main_Info.title
			.title	= "Self HTML"
			.link	= "http://selfhtml.teamone.de/"
			.target = "_blank"
		end with
		menu.addMenuItem(MenuItem)
		set MenuItem = nothing
		set Main_Info = nothing
		
		set Main_Lists = new MenuPoint
		Main_Lists.title	= "Reporting Lists"
		.addMenuItem(Main_Lists)
		
		set MenuItem = new MenuPoint
		with MenuItem
			.parent	= Main_Lists.title
			.title	= "ToDo List"
		end with
		.addMenuItem(MenuItem)
		set MenuItem = nothing
		
		set MenuItem = new MenuPoint
		with MenuItem
			.parent	= Main_Lists.title
			.title	= "Error-log"
			.link 	= "/intranet_programmersHeaven/automated_errors/"
			.target = "main"
		end with
		.addMenuItem(MenuItem)
		set MenuItem = nothing
		
		set MenuItem = new MenuPoint
		with MenuItem
			.parent	= Main_Lists.title
			.title	= "Bug List"
		end with
		.addMenuItem(MenuItem)
		set MenuItem = nothing
		set Main_Lists = nothing
		
		.draw()
		end with
	set menu = nothing
	
	response.write "</TD></TR>"
	response.write "</TABLE>"

end sub
%>
