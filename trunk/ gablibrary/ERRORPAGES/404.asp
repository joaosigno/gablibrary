<!--#include file="../class_page/generatepage.asp"-->
<!--#include file="../class_errorHandler/errorHandler.asp"-->
<!--#include file="../class_textTemplate/TextTemplate.asp"-->
<%
'******************************************************************************************
'* Creator: 	Michal Gabrukiewicz - gabru@gmx.at
'* Description: Displays a 404 error page.
'******************************************************************************************

set page = new GeneratePage
set page = nothing

requestedPage = request.ServerVariables("QUERY_STRING")
requestedPage = mid(requestedPage, inStr(requestedPage, ";") + 1)
referredPage = request.serverVariables("HTTP_REFERER")
host = "http://" & request.serverVariables("HTTP_HOST")

set aErr = new ErrorHandler
with aErr
	.errorDuring = "loading a requested page. The Page (" & requestedPage & ") cannot be found"
	.debuggingVar = requestedPage
	.alternativeText = "The page you are looking for might have been removed, had its name changed, or is temporarily unavailable." &_
						"Please try the following:" &_
						"<UL><LI>If you typed the page address in the Address bar, make sure that it is spelled correctly.</LI>" & _
						"<LI>Open the <a href=""" & host & """ target=_blank>" & host & "</a> home page, and then look for links to the information you want.</LI>" &_
						"<LI>Click the <a href=""javascript:history.back();"">Back</a> button to try another link.</LI></UL>"
	
	'if the page was not reffered then we dont send email and dont log the error
	'because it is not a real error. But if there is a referrer then we do everything.
	'updated: David Rankin
	'If someone else refers to a page on our server, we do not want to know about it, unless
	'the referer is from the same host
	if referredPage = empty or (InStr(referredPage, host) < 1 ) then
		.notifyViaMail = false
		.logging = false
	end if
	.generate()
end with
set aErr = nothing
%>