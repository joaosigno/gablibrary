<!--#include virtual="/gab_Library/class_page/generatepage.asp"-->
<%
'******************************************************************************************
'* Creator: 	Michal Gabrukiewicz
'* Created on: 	2005-00-04 19:00
'* Description: maintenance-site
'******************************************************************************************

set page = new generatePage
with page
	.DBConnection = false
	.title = "Maintenance - Try again later"
	.showOnMaintenanceWork = true
	.showFooter = false
	.loginRequired = false
	.draw()
end with
set page = nothing

'******************************************************************************************
'* main 
'******************************************************************************************
sub main()
	content()
end sub

'******************************************************************************************
'* content 
'******************************************************************************************
sub content()
%>

	<table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%"> 
	<tr>
		<td align="center" valign="middle">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tr>
				<td class="hl3" align="center" style="line-height:18px;">
					<br>
					<IMG src="<%= consts.logo %>" border=0><br><br>
					<strong>Due to maintenance work the requested page<BR>
					is currently not available.</strong><br>
					Please try again later.
					<br><br>
					Thank you for your understanding!<br>
					If you have any questions please contact the IT-Department
					<br><br>
				</td>
			</tr>
			</table>
		</td>
	</tr>
	</table>

<% end sub %>