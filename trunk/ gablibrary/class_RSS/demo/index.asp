<!--#include virtual="/gab_Library/class_page/generatepage.asp"-->
<!--#include virtual="/gab_Library/class_RSS/RSS.asp"-->
<!--#include virtual="/gab_Library/class_cache/cache.asp"-->
<%
'******************************************************************************************
'* Creator: 	Michal Gabrukiewicz
'* Created on: 	2006-11-10 17:09
'* Description: demo for the RSS
'* Input:		-
'******************************************************************************************

set r = new RSS
set page = new GeneratePage
with page
	.loginRequired = false
	.onlyWebDev = true
	.draw()
end with
set page = nothing
set r = nothing

'******************************************************************************************
'* main 
'******************************************************************************************
sub main()
	r.url = "http://rss.orf.at/futurezone.xml"
	'thats the way how to use the cache
	r.setCache "s", 30
	
	'create RSS 2.0 feed
	set r2 = new RSS
	with r2
		.title = "New Feed"
		.link = "http://devinfo.doco.com/gab_Library/class_RSS/demo/rss2.xml"
		.description = "Some feed description here"
		.publishedDate = now()
		set it = new RSSItem
		it.title = "Item 1"
		it.description = "A little descriptiones with <em>html</em>."
		it.publishedDate = dateAdd("d", -2, now())
		it.author = "michal"
		.addItem(it)
		
		set it = new RSSItem
		it.title = "Item 3"
		it.publishedDate = dateAdd("d", -20, now())
		it.category = "ASP"
		it.description = "more moreomeoemeo adkasjdl askdjaskl däö ü A little descriptiones with <em>html</em>."
		it.author = "michalski"
		.addItem(it)
		
		.generate "RSS2.0", "rss2.xml"
		if .failed then str.write("failed to generate.")
	end with
	
	content()
	
end sub

'******************************************************************************************
'* content 
'******************************************************************************************
sub content() %>
	
	<% r.load() %>
	
	<% if not r.failed then %>
		<% for each it in r.items.items %>
			<div><%= it.title %></div>
		<% next %>
	<% end if %>
	
	<%= r.draw("rss.xsl") %>
	
	<% 'after calling draw() or load() you can access the failed property to check if everything was fine... %>
	<% if r.failed then str.write("could not get it... some error ") %>
	
	<h2>Read the above created</h2>
	
	<% set r3 = new RSS : r3.setCache "s", 30 %>
	<% r3.url = "http://localhost/gab_Library/class_RSS/demo/rss2.xml" %>
	
	<% r3.load() %>
	
	language: <%= r3.language %><br>
	title: <%= r3.title %><br>
	link: <%= r3.link %><br>
	dat: <%= r3.publishedDate %><br>
	<% for each it in r3.items.items %>
		<div>[<%= it.category %>] <%= it.title %> (<%= it.publishedDate %>) - <%= it.description %></div>
	<% next %>
	
	<h2>Other RSS 2.0</h2>
	
	<% set r3 = new RSS : r3.setCache "s", 30 %>
	<% r3.url = "http://www.webdevbros.net/feed/" %>
	
	<% r3.load() %>
	
	language: <%= r3.language %><br>
	title: <%= r3.title %><br>
	link: <%= r3.link %><br>
	dat: <%= r3.publishedDate %><br>
	<% for each it in r3.items.items %>
		<div style="padding:20px">[<%= it.category %>] <%= it.author %>: <a href="<%= it.link %>"><%= it.title %></a> (<%= it.publishedDate %>) - <%= it.description %></div>
	<% next %>
	
	<h2>Atom</h2>
	
	<% set r3 = new RSS : r3.setCache "s", 30 %>
	<% r3.url = "http://www.webdevbros.net/feed/atom" %>
	
	<% r3.load() %>
	
	language: <%= r3.language %><br>
	title: <%= r3.title %><br>
	link: <%= r3.link %><br>
	dat: <%= r3.publishedDate %><br>
	<% for each it in r3.items.items %>
		<div style="padding:20px;">[<%= it.category %>] <%= it.author %>: <a href="<%= it.link %>"><%= it.title %></a> (<%= it.publishedDate %>) - <%= it.description %></div>
	<% next %>
	
	<h2>RDF</h2>
	
	<% set r3 = new RSS : r3.setCache "s", 30 %>
	<% r3.url = "http://www.webdevbros.net/feed/rdf" %>
	
	<% r3.load() %>
	
	language: <%= r3.language %><br>
	title: <%= r3.title %><br>
	description: <%= r3.description %><br>
	link: <%= r3.link %><br>
	dat: <%= r3.publishedDate %><br>
	<% for each it in r3.items.items %>
		<div style="padding:20px;">[<%= it.category %>] <%= it.author %>: <a href="<%= it.link %>"><%= it.title %></a> (<%= it.publishedDate %>) - <%= it.description %></div>
	<% next %>
		
<% end sub %>