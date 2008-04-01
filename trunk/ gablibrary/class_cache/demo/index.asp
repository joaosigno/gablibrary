<!--#include virtual="/gab_Library/class_page/generatepage.asp"-->
<!--#include virtual="/gab_Library/class_cache/cache.asp"-->
<%
'******************************************************************************************
'* Creator: 	Michal Gabrukiewicz
'* Created on: 	2006-11-10 17:12
'* Description: demo for the cache
'* Input:		-
'******************************************************************************************

set page = new GeneratePage
with page
	.onlyWebDev = true
	.loginRequired = false
	.draw()
end with
set page = nothing

'******************************************************************************************
'* main 
'******************************************************************************************
sub main()
	set c = new Cache
	
	'set the interval to 10 seconds.
	c.interval = "s"
	c.intervalValue = 10
	c.name = "cacheDemo"
	
	stored = c.getItem("test")
	if stored <> "" then
		str.write("test value from cache = " & stored)
	else
		c.store "test", "michal cached."
	end if
	
	content()
end sub

'******************************************************************************************
'* content 
'******************************************************************************************
sub content() %>

	<div>refresh the page sometimes to see the effect.</div>

<% end sub %>