<!--#include virtual="/gab_Library/class_page/generatepage.asp"-->
<!--#include virtual="/gab_Library/class_calculator/calculator.asp"-->
<%
set page = new generatePage
with page
	.DBConnection 	= false
	.title 			= "Wyeth online calculator" & str.clone("&nbsp;", 100)
	.debugMode		= false
	.loginRequired	= false
	.frameSetter	= false
	.showFooter		= false
	.bodyAttribute	= "topmargin=0 lefmargin=0 onkeyup=""handleKeys();"""
	.devWarning		= false
	.isModalDialog	= true
	.draw()
end with
set page = nothing

'******************************************************************************************
'* main - we start . DONT FORGET TO KILL ALL OBJECTS!! 
'******************************************************************************************
sub main()
	str.writeln("<base target=_self>")
	set calci = new calculator
	with calci
		.JSTarget = request.queryString("JSTarget")
		.displayedValue = request.queryString("displayedValue")
		.commaStyle = request.queryString("commaStyle")
		.draw()
	end with
	set calci = nothing
end sub
%>