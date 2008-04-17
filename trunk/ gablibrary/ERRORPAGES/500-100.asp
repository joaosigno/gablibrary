<!--#include file="../class_page/generatePage.asp"-->
<!--#include file="../class_errorHandler/errorHandler.asp"-->
<!--#include file="../class_textTemplate/TextTemplate.asp"-->
<%
'******************************************************************************************
'* Creator: 	Michal Gabrukiewicz
'* Created on: 	2006-10-27 18:09
'* Description: Error page for all unexpected ASP errors (Internal Server error)
'******************************************************************************************

set page = new GeneratePage
set eHandler = new ErrorHandler
with eHandler
	set .errorObject = server.getLastError()
	.generate()
end with
set eHandler = nothing
set page = nothing
%>

