<!--#include virtual="/gab_Library/class_page/generatepage.asp"-->
<!--#include virtual="/gab_Library/class_documentHolder/documentHolder.asp"-->
<!--#include virtual="/gab_Library/class_fileUpload/fileUpload.asp"-->
<%
'******************************************************************************************
'* Creator: 	Michael Rebec
'* Created on: 	2006-09-29 15:59
'* Description: demo for the documentHolder
'* Input:		-
'******************************************************************************************

set docHolder = new documentHolder
set page = new GeneratePage
with page
	.onlyWebDev = true
	.draw()
end with
set page = nothing
set docHolder = nothing

'******************************************************************************************
'* main 
'******************************************************************************************
sub main()
	str.writeln(page.isPostBack())
	
	with docHolder
		.deleteCaption = "Delete current picture on save"
		.uploader.uploadPath = "/gab_library/class_documentHolder/demo/files/"
		.uploader.defaultFilename = "test"
		.uploader.maxFileSize = 50000
		.uploader.allowedExtensions = "txt, gif, jpg, css"
		'.value = session("demo_filename")
	end with
	
	if page.isPostBack() then
		if docHolder.perform() then
			str.writeln("uploaded.")
			session("demo_filename") = docHolder.value
		else
			str.writeln("error")
		end if
	end if
	
	
	'docHolder.value = uploader.uploadPath & session("demo_filename")
	content()
end sub

'******************************************************************************************
'* content 
'******************************************************************************************
sub content() %>

	<form name="frm" method="post" action="index.asp?" enctype="multipart/form-data">
	
		<table>
			<tr>
				<td>Upload:</td>
				<td><% docHolder.draw() %></td>
			</tr>
		</table>
		<input type="Hidden" name="dummy" value="field">
		<br><br>
		
		<input type="Submit" name="save" value="Save" class="button">
	
	</form>
	
<% end sub %>