<!--#include virtual="/gab_library/class_page/generatepage.asp"-->
<%
set page = new generatePage
with page
	.onlyWebDev = true
	.draw()
end with
set page = nothing

'******************************************************************************************
'* main 
'******************************************************************************************
sub main()
	set logObj = new logger
	with logObj
		.splittingSize = 250
		.identification = "test"
		.onlyOneLogFile = true
		'.log("Seas du")
		
		'delete all logfiles
		'.delete()
	end with
end sub
%>