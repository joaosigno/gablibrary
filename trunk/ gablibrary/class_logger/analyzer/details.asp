<!--#include virtual="/gab_library/class_page/generatepage.asp"-->
<!--#include file="config.inc"-->
<%
set page = new generatePage
with page
	.DBConnection 	= false
	.title 			= "Gablib Log Analyzer"
	.contentSub		= "main"
	.debugMode		= false
	.loginRequired	= false
	.showFooter		= false
	.draw
end with
set page = nothing

'******************************************************************************************
'* main - we start . DONT FORGET TO KILL ALL OBJECTS!! 
'******************************************************************************************
sub main()
	call content()
end sub

sub readLogFile(myFile, ident, desc)
	if myFile = empty then exit sub
	
	desc = decode(desc)
	filePath = Server.MapPath(consts.logs_path & myFile)
	set fso = server.createObject("Scripting.FileSystemObject")
	set fileStream = fso.OpenTextFile(filePath, 1)
	
	do while not fileStream.AtEndOfStream
		line = fileStream.ReadLine()
		if ident = empty then
			if InStr(line, desc) > 1 then
				str.writeln("<div style=""margin-bottom:5px;"">" & line & "</div>")
			end if
		else
			if InStr(line, ident) > 1 and InStr(line, desc) > 1 then
				str.writeln("<div style=""margin-bottom:5px;"">" & AddHref(line) & "</div>")
			end if
		end if
	loop
	
	fileStream.close()
	set fileStream = nothing
	set fso = nothing
end sub

function AddHref(msg)
	tmp = Split(msg, "|")
	ipStr = Mid(msg, 2, Trim(Len(tmp(0))-1))
	newStr = "[ <label title=""" & ipStr & """>" & ipStr & "</label> | "
	
	'*TODO* read default name
	
	AddHref = newStr & tmp(1)
end function

'******************************************************************************************
'* decode
'******************************************************************************************
function decode(msg)
	tmp = replace(msg, "´", """")
	tmp = replace(tmp, "`", "'")
	decode = tmp
end function

'******************************************************************************************
'* content
'******************************************************************************************
sub content()
%>
<%= LOGGER_ADDITIONAL_STYLE %>
<base target="_self">
<link rel="stylesheet" type="text/css" href="style.css">

	<div class="HL">
		GabLIB Log Analyzer Detail View
	</div>

	<div class="infobox">
		<% call readLogFile(request("file"), request("id"), request("description")) %>
	</div>
	
	<div class="endline" style="vertical-align:bottom" align="right">
		<input type="button" value="Close" onclick="window.close()" class="button">
		<button onclick="location.reload()" class="button">Reload</button>
	</div>
<% end sub %>