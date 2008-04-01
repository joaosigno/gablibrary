<!--#include virtual="/gab_library/class_page/generatepage.asp"-->
<!--#include virtual="/gab_library/class_sort/sort.asp"-->
<%
set page = new generatePage
with page
	.DBConnection 	= false
	.onlyWebDev = true
	.draw
end with
set page = nothing

'******************************************************************************************
'* main - we start . DONT FORGET TO KILL ALL OBJECTS!! 
'******************************************************************************************
sub main()

'*** Shufflesort ***
ar = Array("da", "sagte", "die", "katze", "einfach", "servas")
set sorted = new sort
sorted.sourceType 	= TYPE_SHUFFLE
sorted.source		= ar
ar1					= sorted.sort()
set sorted = nothing
	
for j = 0 to ubound(ar1)
	response.write(ar1(j) & "&nbsp;&nbsp;")
next

	response.write "<BR>"

'*** Columnsort ***
ar = Array("A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P")
col = 4
set sorted = new sort
sorted.sourceType 	= TYPE_COLUMN
sorted.column		= col
sorted.source		= ar
ar1					= sorted.sort()
set sorted = nothing
	
for j = 0 to ubound(ar1)
	if j mod col = 0 then response.write("<BR>")
	response.write(ar1(j) & "&nbsp;&nbsp;")
next

%>
<br>
<br>
<br>
<TABLE width="100%">
	<TR>
		<TD align="left">
		<TABLE>
		<TR>
			<TD width="30%"><strong>normaler Bubblesort</strong><BR><BR></TD>
			<TD width="30%"><strong>2dimensionaler Bubblesort</strong><BR><BR></TD>
		</TR>
			<TD>
		<%
		'*** bubbleSort ***
				arrstrbez = Array("WebMaster/in","WebPublisher/in",_
							 "PC-Supporter/in", "Netzwerkspezialist/in", "Applikationsentwickler/in")
				arrintnumbers = Array(6,5,4,3,2,1)
				
				Response.write "<strong>unsortiert</strong><BR>"
				For each arrelement in arrstrbez
				  	Response.Write(arrelement & "<br />" & vbCrLf)
				Next
				
				Response.write "<BR><BR><strong>sortiert</strong><BR>"
				
				set sorted = new sort
				sorted.sourceType 	= TYPE_ARRAY
				sorted.source		= arrstrbez
				myNewArray 			= sorted.sort()
				set sorted = nothing
				For each arrelement in myNewArray
				  	Response.Write(arrelement & "<br />" & vbCrLf)
				Next
		%>
			</TD>
			<TD>
		<%
		'*** bubbleSort2 ***
		Dim arrvar(3,4)
		arrvar(0,0) = 3
		arrvar(1,0) = 2
		arrvar(2,0) = 1
		
		arrvar(0,1) = "A"
		arrvar(1,1) = "D"
		arrvar(2,1) = "C"
		
		arrvar(0,2) = DateValue("1.7.61")
		arrvar(1,2) = DateValue("20.6.59")
		arrvar(2,2) = DateValue("1.1.98")
		
		arrvar(0,3) = "**"
		arrvar(1,3) = "***"
		arrvar(2,3) = "*"
		
		Response.write "<strong>unsortiert</strong><BR>"
		call bubbleDemo(arrvar, empty)
		
		Response.write "<BR><BR><strong>sortiert nach spalte 3</strong><BR>"
		set sorted = new sort
		sorted.sourceType 	= TYPE_ARRAY2
		sorted.sortOrder	= SORT_ASC
		sorted.source		= arrvar
		sorted.column		= 3
		myNewArray 			= sorted.sort()
		set sorted = nothing
		
		call bubbleDemo(myNewArray, 3)
		%>
			</TD>
		</TR>
		</TABLE>
		</TD>
	</TR>
</TABLE>
<%
end sub

Sub bubbleDemo(ByVal arrvar, intcounter)
	Response.Write("<table border=""1"" cellpadding=""5"" cellspacing=""0"">" & vbCrLf)
	For intcounter = 0 to UBound(arrvar,1) - 1
	  Response.Write("  <tr>" & vbCrLf)
	  For intinnercounter = 0 to UBound(arrvar,2) - 1
	    Response.Write("    <td>" & arrvar(intcounter,intinnercounter) & _
	       "</td>" & vbCrLf)
	  Next    
	  Response.Write("  </tr>" & vbCrLf)
	Next
	Response.Write("</table>" & vbCrLf)
End Sub
		
%>