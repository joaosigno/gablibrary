<!--#include virtual="/gab_Library/class_page/generatepage.asp"-->
<%
set page = new generatePage
with page
	.DBConnection 	= false
	.title 			= "Emailtest"
	.debugMode		= false
	.draw()
end with
set page = nothing

'******************************************************************************************
'* main - we start . DONT FORGET TO KILL ALL OBJECTS!! 
'******************************************************************************************
sub main()
	set mailerObject = new customSendMail
	with mailerObject
		.subject = "class_customSendMail Test-email"
		.addRecipient "TO", "gabrukm@wyeth.com", "Michal"
		.body = "Test"
		.htmlEmail = true
		
		if .send() then
			str.writeln("Successfully sent!")
		else
			str.writeln("<strong>Error happend:</strong><br>" & .errormessage)
		end if
	end with
	set mailerObject = nothing
end sub
%>