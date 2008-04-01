<!--#include virtual="/gab_Library/class_page/generatepage.asp"-->
<%
set page = new generatePage
with page
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
%>

<% sub content() %>
	
	<script language="JavaScript">
	<!--
	function showCalendar(obj, mini, maxi) {
		window.showModalDialog('../index.asp?selectedDate=' + obj.value + '&JSTarget=frm.' + obj.name + '&min=' + mini + '&max=' + maxi, window, 'dialogHeight: 280px; dialogWidth: 300px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No; scroll: No');
	}
	//-->
	</script>
	
	<form name="frm">
		Minimum allowed:<br>
		<input type="Text" name="min" size="10" onclick="showCalendar(this, '', '');">
		<br>
		Maximum allowed:<br>
		<input type="Text" name="max" size="10" onclick="showCalendar(this, '', '');"><br>
		<br>
		Date with ranges above:<br>
		<input type="Text" name="date" size="10" onclick="showCalendar(this, min.value, max.value);"><br>
	</form>

<% end sub %>