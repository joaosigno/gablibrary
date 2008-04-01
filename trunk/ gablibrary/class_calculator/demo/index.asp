<!--#include virtual="/gab_Library/class_page/generatepage.asp"-->
<%
set page = new generatePage
with page
	.onlyWebDev = true
	.draw()
end with
set page = nothing

'******************************************************************************************
'* main - we start 
'******************************************************************************************
sub main()
	content()
end sub

'******************************************************************************************
'* content 
'******************************************************************************************
sub content() %>

	<form name="frm">
		<input type="Text" name="val" style="width:200px;">
		<button onclick="openCenteredModal('/gab_Library/class_calculator/demo/calc.asp?JSTarget=frm.val&displayedValue=' + val.value, 295, 251, false)">calculate</button>
	</form>

<% end sub %>