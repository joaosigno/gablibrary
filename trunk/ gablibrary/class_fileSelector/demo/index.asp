<!--#include virtual="/gab_Library/class_page/generatepage.asp"-->
<!--#include file="../fileSelector.asp"-->
<%
set fs = new fileSelector
set page = new generatePage
with page
	.onlyWebDev = true
	.draw()
end with
set page = nothing
set fs = nothing

'******************************************************************************************
'* main - we start . DONT FORGET TO KILL ALL OBJECTS!! 
'******************************************************************************************
sub main()
	fs.sourcePath = "/gab_library/class_dropdown/"
	fs.multipleSelection = false
	'fs.showTreeLines = false
	'fs.showExtensionOfKnownTypes = true
	'fs.filesSelectable = false
	fs.foldersSelectable = true
	if page.isPostback() then
		fs.selected = fs.selected
	else
		fs.selected = array("/gab_Library/class_arrayList/arrayList.asp", "/gab_Library/class_calculator/index.asp", "/gab_Library/class_calculator/icons/icon_0.gif")
	end if
	fs.height = 500
	content()
end sub

'******************************************************************************************
'* content 
'******************************************************************************************
sub content()
%>
	<form action="index.asp" method="post" name="frm">
		<div style="padding:30;"><% fs.draw() %></div>
		<div><input type="Submit" name="submit"></div>
	</form>
	
<% end sub %>
