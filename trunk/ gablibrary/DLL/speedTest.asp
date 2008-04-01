<!--#include virtual="/gab_Library/class_page/generatepage.asp"-->
<%
const AMOUNT = 30000
const ADDY = "hallo<BR>"
set pg = new GeneratePage
pg.draw()

sub main()
	what = request.querystring("what")
	select case what
		case 1
			call adding
		case 2
			call rp
		case 3
			call sb
	end select
	%>
	<UL>
		<LI><a href="speedTest.asp?what=1">Simple adding</a></LI>
		<LI><a href="speedTest.asp?what=2">Response.Write</a></LI>
		<LI><a href="speedTest.asp?what=3">StringBuilder</a></LI>
	</UL>
	<%
end sub

sub adding
	for i = 0 to AMOUNT
		st = st & ADDY
	next
	response.write st
end sub

sub sb
	Set output = Server.CreateObject("StringBuilderVB.StringBuilder")
	output.Init 20000, 7500
	for i = 0 to AMOUNT
		output.append ADDY
	next
	response.write output.tostring
end sub

sub rp
	for i = 0 to AMOUNT
		response.write ADDY
	next
end sub
%>