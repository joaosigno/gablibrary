<!--#include virtual="/gab_Library/class_page/generatepage.asp"-->
<!--#include virtual="/gab_Library/class_mathematics/mathematics.asp"-->
<%
'******************************************************************************************
'* Creator: 	Michael Rebec
'* Created on: 	2006-10-24 15:42
'* Description: -
'* Input:		-
'******************************************************************************************

set math = new mathematics
set page = new GeneratePage
with page
	.onlyWebDev = true
	.draw()
end with
set page = nothing
set math = nothing

'******************************************************************************************
'* main 
'******************************************************************************************
sub main()
	content()
end sub

'******************************************************************************************
'* content 
'******************************************************************************************
sub content() %>

	PI: <%= math.PI() %><br>
	This year in roman letters: <%= math.roman(cInt(year(now))) %>

<% end sub %>