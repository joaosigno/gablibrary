<!--#include virtual="/gab_Library/class_page/generatepage.asp"-->
<!--#include virtual="/gab_Library/class_treeview/treeview.asp"-->
<%
'******************************************************************************************
'* Creator: 	Michal Gabrukiewicz
'* Created on: 	2005-09-22 11:06
'* Description: Demo-page for the treeview
'******************************************************************************************

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
	sql = "SELECT d.id_department, d.department, u.lastname ||' '|| u.firstname AS fullname, u.id_user " & _
			"FROM departments d, users u " & _
			"WHERE u.id_department = d.id_department ORDER BY d.id_department, u.lastname, u.firstname"
	set aTreeview = new Treeview
	with aTreeview
		set .datasource = lib.getRecordset(sql)
		.addGroup(.getNewGroup("id_department", "department"))
		.addGroup(.getNewGroup("id_user", "fullname"))
		.draw()
	end with
	set aTreeview = nothing
end sub
%>