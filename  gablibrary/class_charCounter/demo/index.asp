<!--#include virtual="/gab_Library/class_page/generatepage.asp"-->
<!--#include virtual="/gab_Library/class_charCounter/charCounter.asp"-->
<%
title 			= "myTitle"

set pg	 		= new generatePage
pg.DBConnection	= true
pg.onlyWebDev = true
pg.title		= title
pg.contentSub	= "main"
pg.draw

sub main
	set textarea = new charCounterTextarea
	with textarea
		.name = "rcomment"
		.rows = 3
		.cols = 50
		.value = rcomment
		.attributes = "style='width:100%;'"
		.formName = "frm"
	end with
	
	'* AND NOW THE COUTNER ITSELF 
	set charCount = new charCounter
	with charCount
		.barLength	= "200"
		.maxChars	= 10
		.allocateControl textarea
	end with
	%>
	<FORM name="frm">
	<TABLE border="0" width="100%">
	<TR>
		<TD><%= charCount.drawControl %></TD>
	</TR>
	<TR>
		<TD><%= charCount.drawCounter %></TD>
	</TR>
	<TR>
		<TD><BR><BR></TD>
	</TR>
	<%
	'*****************************************************
	'* MAYBE ANOTHER ONE? 
	'*****************************************************
	set textarea2 = new charCounterTextarea
	with textarea2
		.name = "ccomment"
		.rows = 10
		.value = rcomment
		.attributes = "style='width:300;'"
		.formName = "frm"
	end with
	charCount.iAmNotAlone = true
	charCount.maxChars = 50
	charCount.barLength	= "400"
	charCount.allocateControl textarea2
	%>
	<TR>
		<TD><%= charCount.drawControl %></TD>
	</TR>
	<TR>
		<TD><%= charCount.drawCounter %></TD>
	</TR>
	<TR>
		<TD><BR><BR></TD>
	</TR>
	<%
	'*****************************************************
	'* A TEXTFIELD? WITH VALUE MORE THAN ALLOWED?
	'*****************************************************
	set textfield = new charCounterTextfield
	with textfield
		.name = "test"
		.value = "more than allowed!!!!!"
		.attributes = "size=10"
		.formName = "frm"
	end with
	charCount.iAmNotAlone = true
	charCount.maxChars = 7
	charCount.barLength	= "200"
	charCount.allocateControl textfield
	%>
	<TR>
		<TD><%= charCount.drawControl %></TD>
	</TR>
	<TR>
		<TD><%= charCount.drawCounter %></TD>
	</TR>
	</FORM>
	</TABLE>
	
	<%
end sub
%>