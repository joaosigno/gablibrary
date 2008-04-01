<!--#include virtual="/gab_Library/class_page/generatepage.asp"-->
<!--#include virtual="/gab_Library/class_fileUpload/fileUpload.asp"-->
<%
'******************************************************************************************
'* Creator: 	Michal Gabrukiewicz
'* Created on: 	2006-09-25 12:02
'* Description: Demo for library
'* Input:		-
'******************************************************************************************

set page = new GeneratePage
with page
	.onlyWebDev = true
	.draw()
end with
set page = nothing

'******************************************************************************************
'* main 
'******************************************************************************************
sub main()
	
	set u = new FileUpload
	str.write(typename(lib.form))
	
	if page.isPostback() then
		'using the request.form via the lib.
		str.write(lib.RF("test"))
	end if
	
	content()
end sub

'******************************************************************************************
'* content 
'******************************************************************************************
sub content() %>

	<form method="post" enctype="multipart/form-data">
	
		<input type="Text" name="test" value="<%= lib.RFE("test") %>">
		
		<input type="Submit" name="send">
		
	</form>

<% end sub %>