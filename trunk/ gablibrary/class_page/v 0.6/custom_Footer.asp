<div class="notForPrint">
	<div class="footerLine"></div>
	<div class="footer">
		<span class="mainCol">.:: <%= consts.company_name %></span> Application - 
		<span title="Click & Reload the page:<%= vbcrlf & request.serverVariables("SCRIPT_NAME") %>">
			<a href="#" onclick="window.status='';location.reload();">loaded in <%= pageLoadTime %>s</a>
		</span>
		- <a href="<%= showURL %>" target="_blank"><%= showURL %></A> ::.
	</div>
	
	<% if consts.isDevelopment() then %>
		<div title="Loads the current page on the LIVE-server">
			<a href="http://<%= consts.liveServerName & request.serverVariables("SCRIPT_NAME") & "?" &  request.queryString %>" style="color:#ddd;font-size:7pt;" target="_blank">
				load@live
			</a>
		</div>
	<% end if %>
</div>