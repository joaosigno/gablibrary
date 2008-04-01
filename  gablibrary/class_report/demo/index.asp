<!--#include virtual="/gab_Library/class_page/generatepage.asp"-->
<!--#include virtual="/gab_Library/class_report/report.asp"-->
<%
'******************************************************************************************
'* Creator: 	Michal Gabrukiewicz
'* Created on: 	2006-10-02 21:15
'* Description: demo of the REPORT
'* Input:		-
'******************************************************************************************

set r = new Report
set page = new GeneratePage
with page
	.onlyWebDev = true
	.draw()
end with
set page = nothing
set r = nothing

'******************************************************************************************
'* main 
'******************************************************************************************
sub main()
	with r
		.sql = "SELECT COUNT(*) AS logins FROM vUsers WHERE fk_department = {zone}"
		
		'rows
		set row = new ReportField
		with row
			.caption = "SUM Logins"
			.value = "logins"
		end with
		.addRow(row)
		
		'columns
		set RS = lib.getRecordset("SELECT * FROM department d INNER JOIN (SELECT DISTINCT fk_department FROM vUsers) a ON a.fk_department = d.id_department")
		while not RS.eof
			set col = new ReportField
			with col
				.border = "left right"
				.caption = RS("name")
				.param = RS("id_department")
			end with
			.addCol(col)
			RS.movenext()
		wend
	end with
	
	content()
end sub

'******************************************************************************************
'* content 
'******************************************************************************************
sub content() %>

	<% r.draw() %>

<% end sub %>