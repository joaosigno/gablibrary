<!--#include virtual="/gab_Library/class_page/generatepage.asp"-->
<%
'******************************************************************************************
'* Creator: 	Michal Gabrukiewicz
'* Created on: 	2007-04-26 13:08
'* Description: demo
'* Input:		-
'******************************************************************************************

set page = new GeneratePage
with page
	.onlyWebDev = true
	.ajaxed = true
	.devWarning = false
	.draw()
end with
set page = nothing

'******************************************************************************************
'* sub  
'******************************************************************************************
sub init()
end sub

'******************************************************************************************
'* main 
'******************************************************************************************
sub main()
	set j = new json
	str.writeln(j.toJSON("sepp", array(1, 2, 3), true) & "<br>")
	str.writeln(j.toJSON("sepp", array(1, 2, 3), false) & "<br>")
	set j = nothing
	
	content()
end sub

'******************************************************************************************
'* sub  
'******************************************************************************************
sub callback(action)
	page.return lib.newDict(empty)
	'page.returnValue "yo", false
	'page.return false
end sub

'******************************************************************************************
'* content 
'******************************************************************************************
sub content() %>

	<script>
		function done(u) {
			alert(u.was)
			for (i = 0; i < u.length; i++) {
				$("l").options[$("l").options.length] = new Option(u[i].lastname, u[i].id_user, false, false);
			}
		}
	</script>
	<select id="l"></select>
	<button onclick="gablib.callback('do', done)">...</button>

<% end sub %>