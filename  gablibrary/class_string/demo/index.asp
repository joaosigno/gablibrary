<!--#include virtual="/gab_library/class_page/generatepage.asp"-->
<%
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
	
	str.write("Regular expression check: " & str.matching("w", ".+", true))
	
	str.write(str.clone("<br>", 10))
	
	str.write(str.toInt("x", -1) = -1)
	str.write(lib.getFromQS("id", 0))
	str.write("<br>")
	
	test = str.toCharArray("check this out")
	str.write("<strong>str.toCharArray</strong>: ")
	for i = 0 to ubound(test)
		str.write(test(i) & " ")
	next 
	
	str.write("<strong>str.arrayToString()</strong>: " & test1)
	
	str.write("<BR><BR>")
	test1 = str.arrayToString(test, empty)
	response.write "<strong>str.arrayToString()</strong>: " & test1
	
	response.write "<BR><BR>"
	response.write "<strong>str.startsWith()</strong>: " & str.startsWith(test1, "ch")
	
	response.write "<BR><BR>"
	response.write "<strong>str.endWith()</strong>: " & str.endsWith(test1, "out")
	
	response.write "<BR><BR>"
	response.write "<strong>str.clone()</strong>: " & str.clone("abc", 10)
	
	response.write "<BR><BR>"
	response.write "<strong>str.trimStart()</strong>: " & str.trimStart(test1, 3)
	
	response.write "<BR><BR>"
	response.write "<strong>str.trimEnd()</strong>: " & str.trimEnd(test1, 2)
	
	response.write "<BR><BR>"
	response.write "<strong>str.swapCase()</strong>: " & str.swapCase("HiHiHi")
	
	response.write "<BR><BR>"
	response.write "<strong>str.isAlphabetic()</strong>: " & str.isAlphabetic("!")
	
	response.write "<BR><BR>"
	response.write "<strong>str.capitalize()</strong>: " & str.capitalize("clara fehler")
	
	response.write "<BR><BR>"
	response.write "<strong>str.format()</strong>: " & str.format("hier {0} und hier {1} und da {2}", array("eins", 2, 3))
	
	response.write "<BR><BR>"
	response.write "<strong>str.splitValue()</strong>: " & str.splitValue("was geht hier", " ", 1)
	
	response.write "<BR><BR>"
	response.write "<strong>str.shorten()</strong>: " & str.shorten("ganz viel text viel viel text", 10, "...")
	
	response.write "<BR><BR>"
	response.write "<strong>str.divide()</strong>: " & str.arrayToString(str.divide("01234567890", 20), "-")
	
	response.write "<BR><BR>"
	response.write "<strong>str.stripTags()</strong>: " & str.stripTags("<strong>ohne Tags</strong>")
	
	response.write "<BR><BR>"
	response.write "<strong>str.trimComplete()</strong>: " & str.trimComplete(vbcrlf & vbcrlf & "   hey  " & vbcrlf)
	
	content()
end sub

'******************************************************************************************
'* main 
'******************************************************************************************
sub content() %>

	<style>
		#content, #content * {
			font-family:courier;
		}
	</style>
	
	<div id="content">
		
		<div>
			<strong>padleft:</strong> <%= str.padLeft(292, 10, 0) %>
		</div>
		<div>
			<strong>padRight:</strong> <%= str.padRight(292, 10, 0) %>
		</div>
		
	</div>

<% end sub %>