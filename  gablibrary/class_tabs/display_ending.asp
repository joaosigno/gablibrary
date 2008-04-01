	
	<td width="100%" class="tab_end_line">
		<%
		if not addProcedure = empty then
			execute(addProcedure)
		else
			str.write("&nbsp;")
		end if
		%>
	</td>
</tr>
</table>