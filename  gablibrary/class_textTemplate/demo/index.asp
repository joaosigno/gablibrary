<!--#include virtual="/gab_Library/class_page/generatepage.asp"-->
<!--#include virtual="/gab_Library/class_textTemplate/textTemplate.asp"-->
<%
'******************************************************************************************
'* Creator: 	David Rankin
'* Created on: 	2006-09-06 13:10
'* Description: Demo for Text Parser
'* Input:		-
'******************************************************************************************
set template = new textTemplate
set page = new GeneratePage
with page
	.plain = true
	.onlyWebDev = true
	.draw()
end with
set page = nothing
set template = nothing

'******************************************************************************************
'* main 
'******************************************************************************************
sub main()
	
	with template
		.fileName = "/gab_Library/class_textTemplate/demo/demoTemplate.html"
		'case unsensitive..
		'.utf8 = false
		.addVariable "version", "0.2"
		.addVariable "MODIFIED", "06.09.2006"
		.addVariable "NAME", "David Rankin"
		set block = new TextTemplateBlock
		block.addItem(array("WEEKDAY", "Monday", "VALUE", vbMonday))
		block.addItem(array("WEEKDAY", "Tuesday", "VALUE", vbTuesday))
		block.addItem(array("WEEKDAY", "Friday", "VALUE", vbFriday))
		.addVariable "WEEKDAYS", block
		'.cleanParse = false
	end with
	
	content()
end sub

'******************************************************************************************
'* content 
'******************************************************************************************
sub content() %>
	
	<strong>ALL</strong>
	
	<p><%= template.returnString() %></p>
	
	<hr>
	
	<strong>FIRST LINE ONLY</strong>
	
	<p><%= template.getFirstLine()%></p>
	
	<hr>
	
	<strong>ALL BUT FIRST LINE</strong>
	
	<p><%= template.getAllButFirstLine() %></p>
	
<% end sub %>
