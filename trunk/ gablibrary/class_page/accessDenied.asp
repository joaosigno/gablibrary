<!--#include virtual="/gab_Library/class_page/generatepage.asp"-->
<%
set page = new generatePage
with page
	.DBConnection 	= false
	.debugMode		= false
	.loginRequired	= false
	.devWarning		= false
	.draw()
end with
set page = nothing

'******************************************************************************************
'* main - we start . DONT FORGET TO KILL ALL OBJECTS!! 
'******************************************************************************************
sub main()
	content()
end sub
%>

<% sub content() %>

	<h1>Access denied!</h1>
	<br>

	<div style="text-align:right;color:red;position:absolute;right:30px;font-size:7pt;font-weight:normal;">
		<img src="<%= consts.logo %>" alt="<%= consts.company_name %>" border="0">
	</div>
	
	<div class="error">
		Sorry, you don't have permission to access this page.<br>
		Please contact your local Administrator<br>
		if you have questions.
		<br>
	</div>
	
	<br><br><br>
	<div class="endline">
		<button class="button" onclick="parent.top.location.href = '<%= consts.domain %>'"><%= consts.domain %></button>
	</div>
	
<% end sub %>