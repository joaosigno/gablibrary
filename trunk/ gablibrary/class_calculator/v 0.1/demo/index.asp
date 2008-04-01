<!--#include virtual="/gab_Library/class_page/generatepage.asp"-->
<%
set page = new generatePage
with page
	.DBConnection 	= false
	.title 			= "Calculator Demo"
	.debugMode		= false
	.loginRequired	= false
	.draw()
end with
set page = nothing

'******************************************************************************************
'* main - we start 
'******************************************************************************************
sub main()
	call content()
end sub

'******************************************************************************************
'* content 
'******************************************************************************************
sub content()
%>

	<form name="frm">
		
		Custom: <input type="Text" name="cust" class="formField" style="width:200px;">
		<% call lib.custom.drawCalculatorButton("frm.cust", ",", "/gab_Library/class_calculator/calc_custom.asp") %>
		<button onclick="openCenteredModal('/gab_Library/class_calculator/calc_custom.asp?JSTarget=frm.cust&displayedValue=' + cust.value, 295, 251, false)">calculate</button>
		<br>
		
		
		Common: <input type="Text" name="val" class="formField" style="width:200px;">
		<% call lib.custom.drawCalculatorButton("frm.val", ",", empty) %>
		
		<br>
		
		Custom Comma: <input type="Text" name="comma" class="formField" style="width:200px;">
		<% call lib.custom.drawCalculatorButton("frm.comma", ":", empty) %>
	</form>

<% end sub %>