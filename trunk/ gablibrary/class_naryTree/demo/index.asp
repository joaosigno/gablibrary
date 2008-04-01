<!--#include virtual="/gab_Library/class_page/generatepage.asp"-->
<!--#include virtual="/gab_Library/class_naryTree/naryTree.asp"-->
<%
'******************************************************************************************
'* Creator: 	David Rankin
'* Created on: 	2006-06-14 11:51
'* Description: treenodetest
'* Input:		-
'******************************************************************************************

set page = new GeneratePage
with page
	.onlyWebDev = true
	.draw()
end with
set page = nothing

'******************************************************************************************
'* main 
'******************************************************************************************
sub main()
	set aTree = new NaryTree
	with aTree
		set nodes = server.createObjecT("scripting.dictionary")
		nodes.add "josie", "gabru"
		nodes.add "reb", "gabru"
		nodes.add "stehle", "hesel"
		nodes.add "tom", "mies"
		nodes.add "schatz", "mies"
		nodes.add "hesel", empty
		nodes.add "gabru", "stehle"
		nodes.add "max", "schatz"
		nodes.add "david", "stehle"
		nodes.add "mies", "stehle"
		
		.addNodes(nodes)
		
		'set n = aTree.find("stehlec")
		'if typename(n) <> "Nothing" then
			'str.write("<table border=1>")
			'str.write("<tr align=center><td colspan=" & n.childs.count & ">" & n.value & "</td></tr>")
			'str.write("<tr align=center>")
			'for each child in n.childs.items
				'str.write("<td>" & child.value & "</td>")
			'next
			'str.write("</tr>")
			'str.write("</table>")
		'end if
		'set n = nothing
		
		.getChildren("gabru")
	end with
	
	set aTree  = nothing
end sub

'******************************************************************************************
'* content 
'******************************************************************************************
sub content() %>
	
	

<% end sub %>