<!--#include virtual="/gab_Library/class_page/generatepage.asp"-->
<!--#include virtual="/gab_Library/class_datePicker/datePicker.asp"-->
<%
autoResize = false
if autoResize then
	addString = "syncHeight();"
else
	addString = empty
end if

set page = new generatePage
with page
	.DBConnection 	= false
	.title 			= "Please pick a date" & str.clone("&nbsp;", 100)
	.debugMode		= false
	.loginRequired	= false
	.frameSetter	= false
	.showFooter		= false
	.bodyAttribute	= "topmargin=0 lefmargin=0 onLoad=""init();" & addString & """ onkeyup=""handleKeys();"""
	.devWarning		= false
	.isModalDialog	= true
	
	if consts.isDevelopment then
		.pageEnterEffect = true
	end if
	.draw()
end with
set page = nothing

'******************************************************************************************
'* main - we start . DONT FORGET TO KILL ALL OBJECTS!! 
'******************************************************************************************
sub main()
	str.writeln("<base target=_self>")
	set cal = new datePicker
	with cal
		.autoResize = autoResize
		.selectedDate = request.queryString("selectedDate")
		.JSTarget = request.queryString("JSTarget")
		.maximumAllowedDate = request.queryString("max")
		.minimumAllowedDate = request.queryString("min")
		.draw()
	end with
end sub
%>