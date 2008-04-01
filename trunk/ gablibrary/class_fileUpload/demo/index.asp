<!--#include virtual="/gab_Library/class_page/generatepage.asp"-->
<!--#include virtual="/gab_Library/class_fileUpload/fileUpload.asp"-->
<%
'******************************************************************************************
'* Creator: 	Michal Gabrukiewicz
'* Created on: 	2007-03-20 11:30
'* Description: ahxn
'* Input:		-
'******************************************************************************************

set fu = new FileUpload
set page = new GeneratePage
with page
	.onlyWebDev = true
	.draw()
end with
set page = nothing
set fu = nothing

'******************************************************************************************
'* main 
'******************************************************************************************
sub main()
	if page.ispostback() then
		with fu
			.defaultFilename = "wasis"
			.allowedExtensions = "gif,txt,sys"
			.uploadPath = "/dev/"
			.overwrite = false
			.filename = "file"
			.saveUnique = true
			str.write(.upload())
			str.write(.getErrorMsg)
		end with
	end if
	content()
end sub

'******************************************************************************************
'* content 
'******************************************************************************************
sub content() %>

	<form action="haxn.asp" method="post" enctype="multipart/form-data">
	
		<input type="File" name="file">
	
		<input type="Submit">
	
	</form>	

<% end sub %>