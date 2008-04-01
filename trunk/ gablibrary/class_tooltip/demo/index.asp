<!--#include virtual="/gab_library/class_page/generatepage.asp"-->
<%
set page = new generatePage
with page
	.loadTooltips	= true
	.onlyWebDev = true
	.draw()
end with
set page = nothing

'******************************************************************************************
'* main - we start . DONT FORGET TO KILL ALL OBJECTS!! 
'******************************************************************************************
sub main()
	call content()
end sub

sub content()
%>	
	<form name="frm">
		<input type="button" style="position:absolute;top:20px;left:10px;" value="clickst du hier" <%= lib.tooltip("mein titel", "der text dazu<br>und noch etwas dazu") %> class="button">
		<input type="button" style="position:absolute;bottom:20px;left:10px;" value="clickst du hier" <%= lib.tooltip("mein titel", "der text dazu<br>und noch etwas dazu") %> class="button">
		<input type="button" style="position:absolute;top:20px;right:10px;" value="clickst du hier" <%= lib.tooltip("mein titel", "der text dazu<br>und noch etwas dazu") %> class="button">
		<input type="button" style="position:absolute;bottom:20px;right:10px;" value="clickst du hier" <%= lib.tooltip("mein titel", "der text dazu<br>und noch etwas dazu") %> class="button">
	</form>
	
<% end sub %>