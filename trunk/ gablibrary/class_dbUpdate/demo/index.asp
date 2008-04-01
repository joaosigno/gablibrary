<!--#include virtual="/gab_Library/class_page/generatepage.asp"-->
<!--#include virtual="/gab_Library/class_dbUpdate/dbUpdate.asp"-->
<%
title 			= "myTitle"

set pg	 		= new generatePage
pg.DBConnection	= true
pg.onlyWebDev = true
pg.title		= title
pg.contentSub	= "main"
pg.draw

sub main
	if request.form.count > 0 then
	
	set db = new dbUpdate
	db.table = "myTable"				'DB Tablename
	db.pkName = "id"					'Name of the primaryfield in our DB
	db.pk = request.querystring("id")	'Number of record we want to update or delete
	
	
		select case request.form("action")
			case "neu"
				res = db.insert
			case "edit"
				res = db.update
			case "del"
				res = db.delete
		end select
		
		if res then
			response.write "erfolgreich!"
		end if
	end if
end sub

sub myform() %>

<FORM action="<%= request.servervariables("URL") %>" method=POST>
<TABLE border="0">
<TR>
	<TD>Test</TD>
	<TD><input type="text" name="test"></TD>
</TR>
<TR>
	<TD>Test 2</TD>
	<TD><input type="text" name="test2"></TD>
</TR>
<TR>
	<TD colspan="2">
		<BUTTON type="submit" name="action">neu</BUTTON>
		<BUTTON type="submit" name="action">edit</BUTTON>
		<BUTTON type="submit" name="action">del</BUTTON>
	</TD>
</TR>
</TABLE>
</FORM>

<% end sub %>