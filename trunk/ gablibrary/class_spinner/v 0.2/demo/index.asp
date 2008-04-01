<!--#include virtual="/gab_library/class_page/generatepage.asp"-->
<!--#include virtual="/gab_library/class_spinner/spinner.asp"-->
<%
set page = new generatePage
with page
	.DBConnection 	= false
	.title 			= "Wyeth Intranet"
	.contentSub		= "main"
	.debugMode		= false
	.loginRequired	= false
	.draw
end with
set page = nothing

'******************************************************************************************
'* main - we start . DONT FORGET TO KILL ALL OBJECTS!! 
'******************************************************************************************
sub main()
set spin = new spinner
%>
	<br>
	<br>
	<div align="center">
<%
	'spin.skin 			= SKIN_TWO
	spin.minimum		= 1
	spin.maximum		= 100
	spin.width			= 30
	call spin.draw("participants", 1)
	str.writeln("<br>")
	
	spin.Orientation = ORIENTATION_LEFTRIGHT
	spin.width		= 100
	spin.minimum	= 0
	spin.maximum	= 10
	spin.step		= 1
   	call spin.draw("test1", "10")
	str.writeln("<br>")
	
	spin.width		= 100
	spin.minimum	= 0
	spin.maximum	= 10
	spin.step		= 1
	spin.looping	= false
	call spin.draw("asdgk", "5")
	str.writeln("<br>")
	
	spin.Orientation = ORIENTATION_TOPDOWN
	spin.minimum 	= 0
	spin.maximum 	= 100
	spin.step		= 10
	spin.looping	= true
	call spin.draw("test2", "0")
	str.writeln("<br>")
	
	spin.minimum 	= 0
	spin.maximum 	= 10
	spin.step		= "0.001"
	spin.decimalPlaces = 3
	spin.looping	= false
	call spin.draw("test3", "5.000")
	str.writeln("<br>")
set spin = nothing

set spin = new spinner
myArray = Array("Monday", "Tuesday", "Wednsday", "Thursday", "Friday", "Saturday", "Sunday")
	
	spin.firstInstance 	= false
	spin.ControlType	= CONTROL_CUSTOM
	spin.ControlItems	= myArray
	spin.minimum		= 0
	spin.maximum		= 6
	spin.readonly		= true
	spin.looping		= true
	str.writeln("<br>")
	call spin.draw("daytest", "Thursday")
	str.writeln("<br>")
	
	spin.looping		= false
	call spin.draw("daytest1", "Thursday")
	str.writeln("<br>")
	
	spin.Orientation	= ORIENTATION_LEFTRIGHT
	spin.onchange		= "daytest1.value = daytest2.value;"
	call spin.draw("daytest2", "Thursday")
	str.writeln("<br>")
set spin = nothing


set spin = new spinner
set spin1 = new spinner
	
	spin.firstInstance	= false
	spin.minimum		= 0
	spin.maximum		= 23
	spin.integerPlaces	= 2
	spin.looping		= true
	spin.width			= 50
	
	spin1.firstInstance	= false
	spin1.minimum		= 0
	spin1.maximum		= 45
	spin1.step			= 15
	spin1.integerPlaces	= 2
	spin1.looping		= false
	spin1.width			= 50
		%>
	<br><br>
		<table>
			<tr>
				<td><% call spin.draw("hour", hour(now)) %></td>
				<td>&nbsp;&nbsp;&nbsp;</td>
				<td><% call spin1.draw("minute", "00") %></td>
			</tr>
		</table>
	</div> 
<%
end sub
%>