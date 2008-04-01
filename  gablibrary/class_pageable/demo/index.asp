<!--#include virtual="/gab_Library/class_pageable/pageable.asp"-->
<%
'liefert die aktuelle Seitennummer
function getCurrentPage()
	getCurrentPage = 1
	if request("page") <> "" then
		getCurrentPage = request("page")
	end if
end function

'zeichnet einen button
sub drawButton(value, pageNr, cssClass, disabled, color, fontsize)
	with response
		if disabled then dis = " disabled"
		onClick = "window.location.href='index.asp?page=" & pageNr
		.write("<button type=button class=""btn " & cssClass & """ style=""color:#" & color & ";font-size:" & fontsize & "px"" onclick=""" & onClick & "'""" & dis & ">")
		.write(value)
		.write("</button>")
	end with
end sub

'zeichnet unsere seiten-bar
sub drawPagingBar()
	drawButton "&lt;&lt;", paging.firstPage, "prevnext", paging.isOnFirstPage(), "000", 12
	drawButton "&lt;", paging.currentPage - 1, "prevnext", not paging.hasPreviousPage(), "000", 12
	if paging.hasPreviousBlock() then response.write("....")
	
	con = 0
	for i = 0 to uBound(paging.pages)
		if paging.currentpage = paging.page(i) then con = i
		if i < cInt(paging.numberOfPages / 2) then
			leftSide = paging.currentpage - paging.page(lBound(paging.pages)) +1
			h = hex(255-((255 / leftSide) * i))
			'f = cint((20 / leftSide)) * i
		else
			rightSide = paging.page(uBound(paging.pages)) - paging.currentpage
			h = hex((255 / rightSide) * (i - con))
			'f = 20 - ((cint(20 / rightSide) * (i - con)))
		end if
		drawButton paging.page(i), paging.page(i), "common", false, h & h & h, f
	next
	
	if paging.hasNextBlock() then response.write("....")
	drawButton "&gt;", paging.currentPage + 1, "prevnext", not paging.hasNextPage(), "000", 12
	drawButton "&gt;&gt;", paging.lastPage, "prevnext", paging.isOnLastPage(), "000", 12
end sub

'wir verwenden ein array bef&uuml;llt mit daten
dim data(100)
for i = 0 to uBound(data)
	randomize()
	data(i) = chr(255 * (rnd()))
next

'pageable wird instanziert und mit den notigen werten versorgt
set paging = new pageable
with paging
	.currentPage = getCurrentPage()
	.recordCount = uBound(data) + 1
	.recordsPerPage = 3
	.numberOfPages = 20
	.perform()
end with
%>

<style>
	.btn { width:30; border:0; background-color:white; }
	.common { color:black;width:20 }
	.prevnext { font-weight:bold; }
</style>

<form>
	<div align="center"><% drawPagingBar() %></div>
	<br>
	<div>
		<% for i = paging.dataStartPosition to paging.dataEndPosition %>
			<div align="center">Data<%= i - 1%>: <%= data(i - 1) %></div>
		<% next %>
	</div>
</form>

<% set paging = nothing %>