<!--#include virtual="/gab_library/class_page/generatepage.asp"-->
<!--#include virtual="/gab_library/class_encryption/class_gabLibCrypt.asp"-->
<%
set page = new generatePage
with page
	.DBConnection 	= false
	.title 			= "Wyeth Intranet"
	.contentSub		= "main"
	.debugMode		= false
	.loginRequired 	= true
	.devWarning		= false
	.draw
end with
set page = nothing

'******************************************************************************************
'* main - we start . DONT FORGET TO KILL ALL OBJECTS!! 
'******************************************************************************************
sub main()
	str.writeln("SessionID: " & Session.SessionID & "<br><br>")
	
	str.writeln(str.clone("*", 50) & "<br>")
	str.writeln("* >>" & decode("code") & "<<<br>")
	str.writeln(str.clone("*", 50) & "<br><br>")
	call content()
	call drawTable()
end sub

'******************************************************************************************
'* content
'******************************************************************************************
sub content()
%>
	<a href="crypt.asp?code=<%=enc("false", false)%>">Use Normal encryption</a>&nbsp;&nbsp;|&nbsp;&nbsp;
	<a href="crypt.asp?code=<%=enc("true", true)%>">Use Session ecnryption</a>&nbsp;&nbsp;|&nbsp;&nbsp;
	<a href="crypt.asp?code=<%=enc("this is a LARGER text with whitespaces @ (""without SID"")", true)%>">Larger Text (with SID)</a>&nbsp;&nbsp;|&nbsp;&nbsp;
	<a href="crypt.asp?code=<%=enc("this is a LARGER text with whitespaces  & ('with SID')", false)%>">Larger Text (without SID)</a>&nbsp;&nbsp;|&nbsp;&nbsp;
	<a href="crypt.asp">Nothing</a>
<%
end sub

'******************************************************************************************
'* decode
'******************************************************************************************
function decode(field)
	set cr	= new gabLibCrypt
	tmp 	= request.querystring(field)
	decode 	= cr.deCrypt(tmp)
	set cr 	= nothing
end function

'******************************************************************************************
'* enc
'******************************************************************************************
function enc(msg, flag)
	set cr 			= new gabLibCrypt
	cr.useSessionID = flag
	enc 			= cr.enCrypt(msg)
	set cr 			= nothing
end function

'******************************************************************************************
'* drawTable
'******************************************************************************************
sub drawTable()
	for i = 33 to 255
		b = b & chr(i)
		'response.write(chr(i))
	next
	
	set cr = new gabLibCrypt
		x = cr.enCrypt(b)
		y = cr.deCrypt(x)
	set cr = nothing
	
	set cr = new gabLibCrypt
		cr.useSessionID = true
		sx = cr.enCrypt(b)
		sy = cr.deCrypt(sx)
	set cr = nothing
%>
	<br>
	<br>
	<br>
	<br>
	En/Decryption Table (handle with care - could be used to decode gablibCrypt !!)
	<table border=1 cellspacing=0 cellpadding=2>
		<tr>
			<th>ASCII</th>
			<th>Ausgang</th>
			<th>Dec N</th>
			<th>Dec S</th>
			<th>Cod N</th>
			<th>Cod S</th>
		</tr>
		<% j = 1 : for i = 1 to 224 %>
			<tr>
				<td align=center><%= i+32 %></td>
				<td align=center><%= mid(b, i, 1) %></td>
				<td align=center><%= mid(y, i, 1)%></td>
				<td align=center><%= mid(sy, i, 1)%></td>
				<td align=center><%= mid(x, j, 2) %></td>
				<td align=center><%= mid(sx, j, 2) %></td>
			</tr>
		<% j = j + 2 : next %>
	</table>
<%
end sub
%>