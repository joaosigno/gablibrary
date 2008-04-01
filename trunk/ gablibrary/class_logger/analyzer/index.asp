<!--#include virtual="/gab_library/class_page/generatepage.asp"-->
<!--#include file="config.inc"-->
<%
set dict = server.createobject("scripting.dictionary")
dim m_fileSize
const m_version = "0.4"

set page = new generatePage
with page
	.DBConnection 	= true
	.title 			= "Log Analyzer"
	.contentSub		= "main"
	.loginRequired	= true
	.debugMode		= false
	.draw()
end with
set page = nothing
set dict = nothing

'******************************************************************************************
'* main - we start . DONT FORGET TO KILL ALL OBJECTS!! 
'******************************************************************************************
sub main()
	if page.IsPostBack() then
		call readLogFile(request.form("files"))
	end if
	
	call content()
end sub

'******************************************************************************************
'* drawStat
'******************************************************************************************
sub drawStat()
	num = dict.count
	if num = 0 then exit sub
	dim maxCol : maxCol = 255
	dim width : width = 200
	dim maxCount : maxCount = 0
	dim ident : ident = empty
	dim desc : desc = empty
	dim myCol(3)
	
	if request("options") = "1" then
		str.writeln("<fieldset class=""fieldset""><legend class=""alHeader"">Overview</legend><div class=""alContent content"">")
		str.writeln("<table cellspacing=10 cellpadding=0 width=""100%"">")
		str.writeln("<tr><td colspan=2><div class=""hl2"">Total entries: " & getMaximumCount(empty))
		str.writeln("&nbsp;&nbsp;&nbsp;FileSize: " & Round(m_fileSize/1000, 2) & " KB</div></td></tr>")
	else
		str.writeln("<fieldset class=""fieldset""><legend class=""alHeader"">Detailed View</legend><div class=""alContent content"">")
		str.writeln("<table cellspacing=10 cellpadding=0 width=""100%"">")
	end if
	
	total = 0
	totalDays = 0
	totalWithoutWeekend = 0
	totalWorkdays = 0
	
	for each k in dict.keys
		str.writeln("<tr>")
		
		maxCount = getMaximumCount(k)
		mathWidth = round((width / 100) * (dict(k) / maxCount * 100))
		myCol(1) = cInt(rnd() * maxCol) : myCol(2) = CInt(rnd() * maxCol) : myCol(3) = cInt(rnd() * maxCol)
		
		if request("options") = "1" then
			desc = k
		else
			tmpString = Split(k, "**")
			if not tmpString(0) = ident then
				ident = tmpString(0)
				str.writeln("<tr><td colspan=2><div class=""hl2"">" & getWeekDay(ident) & " - Total: " & maxCount & "</div></td></tr>")
			end if
			desc = tmpString(1)
		end if
		
		str.writeln("<td nowrap>")
		str.writeln("<span class=""barEnd""></span>")
		str.writeln("<span style=""width:" & mathWidth & "px;"" class=""barUsed""></span>")
		str.writeln("<span class=""barUsedEnd""></span>")
		str.writeln("<span style=""width:" & (width - mathWidth) & "px;"" class=""bar""></span>")
		str.writeln("<span class=""barEnd""></span>")
		str.writeln("<span class=""percent"">" & dict(k) & " ("& Round(dict(k)/maxCount*100) &"%)</span><br><br>")
		str.writeln("</td><td valign=""top"" width=""100%"">")
		str.writeln("<span class=""description"">" & getDescription(desc) & "</span>")
		str.writeln("<a href=""javascript:showDetails('" & ident & "', '" & recodeDescription(desc) & "','" & request("files") & "');"">details...</a></td></tr>")
		
		total = total + maxCount
		totalDays = totalDays + 1
		
		if not weekday(ident) = vbSunday and not weekday(ident) = vbSaturday then
			totalWithoutWeekend = totalWithoutWeekend + maxCount
			totalWorkdays = totalWorkdays + 1
		end if
	next
	
	totalWeekend = totalDays - totalWorkdays
	
	avgPerDay = 0
	if totalDays > 0 then
		avgPerDay = round(total / totalDays, 2)
	end if
	
	avgPerWorkdays = 0
	if totalWorkdays > 0 then
		avgPerWorkdays = round(totalWithoutWeekend / totalWorkdays, 2)
	end if
	
	avgPerWeekend = 0
	if totalWeekend > 0 then
		avgPerWeekend = (total - totalWithoutWeekend) / totalWeekend
	end if
	
	str.writeln("</fieldset>")
	str.writeln("<u>Avg entries per day:</u> " & avgPerDay & "&nbsp;&nbsp;&nbsp;&nbsp;<u>Avg entries per workday:</u> " & avgPerWorkdays)
	str.writeln("<br><u>Avg entries per weekend-day:</u> " & avgPerWeekend)
	str.writeln("</table></div>")
end sub

'******************************************************************************************
'* getDescription
'******************************************************************************************
function getDescription(msg)
	if Len(msg) > LOGGER_SPLIT_LENGTH then
		getDescription = Left(msg, LOGGER_SPLIT_LENGTH) & "..."
	else
		getDescription = msg
	end if
end function

'******************************************************************************************
'* recodeDescription
'******************************************************************************************
function recodeDescription(msg)
	tmp = replace(msg, """", "´")
	tmp = replace(tmp, "'", "`")
	recodeDescription = tmp
end function

'******************************************************************************************
'* getWeekDay
'******************************************************************************************
function getWeekDay(id)
	tmp = weekdayname(weekday(CDate(id)))
	
	if date = CDate(id) then
		if weekday(CDate(id)) = vbSunday or weekday(CDate(id)) = vbSaturday then
			getWeekDay = "<span style=""color:red"">" & id & " (" & LOGGER_DAY_TODAY & ")</span>"
		else
			getWeekDay = "<span>" & id & " (" & LOGGER_DAY_TODAY & ")</span>"
		end if
		exit function
	elseif date = CDate(id)+1 then
		if weekday(CDate(id)) = vbSunday or weekday(CDate(id)) = vbSaturday then
			getWeekDay = "<span style=""color:red"">" & id & " (" & LOGGER_DAY_YESTERDAY & ")</span>"
		else
			getWeekDay = "<span>" & id & " (" & LOGGER_DAY_YESTERDAY & ")</span>"
		end if
		exit function
	end if
	if weekday(CDate(id)) = vbSunday or weekday(CDate(id)) = vbSaturday then
		getWeekDay = "<span style=""color:red"">" & id & " (" & tmp & ")</span>"
	else
		getWeekDay = "<span>" & id & " (" & tmp & ")</span>"
	end if
	
end function

'******************************************************************************************
'* getMaxCount
'******************************************************************************************
function getMaximumCount(ident)
	dim tmp
	if InStr(ident, "**") > 1 then
		tmpString = Split(ident, "**")
		tmp = tmpString(0)
		
		for each k in dict.keys
			if str.startsWith(k, tmp) then
				getMaximumCount = getMaximumCount + dict(k)
			end if
		next
	else
		for each k in dict.keys
			getMaximumCount = getMaximumCount + dict(k)
		next
	end if
end function

'******************************************************************************************
'* analyze
'******************************************************************************************
private sub analyze(msg)
	if InStr(msg, "[") < 1 then exit sub
	dim flag: 	flag = false
	
	start = InStr(msg, "|") + 1
	count = InStr(msg, "]") - start
	dateString = Split(Trim(Mid(msg, start, count)), " ")
	
	intTab = InStr(msg, vbTab)
	count = InStr(intTab, msg, " ")
	
	if request("options") = "1" then
		identString = Trim(Mid(msg, count))
	else
		identString = dateString(0) & "**" & Trim(Mid(msg, count))
	end if
	
	for each k in dict.Keys
		if k = identString then
			dict(k) = dict(k) + 1
			flag = true
			exit for
		end if
	next
	
	if not flag then
		dict.add identString, 1
	end if
end sub

'******************************************************************************************
'* parseLogDirectory
'******************************************************************************************
function parseLogDirectory()
	counter = 0
	dim query : query = empty
	set objFSO = server.createObject("Scripting.FileSystemObject")
	set folder = objFSO.getFolder(server.mapPath(consts.logs_path))
	
	for each file in folder.files
		filename = file.name
		if str.endsWith(ucase(filename), ".TXT") then
			counter = counter + 1
			query = query & filename & ":"
		end if
	next
	
	if not counter = 0 then
		query = str.trimEnd(query, 1)
		call drawDropDown(query, query, "files")
	else
		str.writeln("No files found !!")
	end if
	
	set objFSO = nothing
	set folder = nothing
	parseLogDirectory = counter
end function

'******************************************************************************************
'* drawOptions
'******************************************************************************************
sub drawOptions()
	call drawDropDown("Overview:Details", "1:2", "options")
end sub

'******************************************************************************************
'* drawDropDown
'******************************************************************************************
sub drawDropDown(query, pk, name)
	set dd = new createDropdown
	with dd
		.sqlQuery		= query
		.pk				= pk
		.name			= name
		.idToMatch		= request(name)
		.draw
	end with
	set dd = nothing
end sub

'******************************************************************************************
'* readLogFile
'******************************************************************************************
sub readLogFile(myFile)
	if myFile = empty then exit sub
	
	filePath = Server.MapPath(consts.logs_path & myFile)
	set fso = server.createObject("Scripting.FileSystemObject")
	set fileInfo = fso.GetFile(filePath)
		m_fileSize = fileInfo.size
	set fileInfo = nothing
	set fileStream = fso.OpenTextFile(filePath, 1)
	
	do while not fileStream.AtEndOfStream
		call analyze(fileStream.ReadLine())
	loop
	
	fileStream.close()
	set fileStream = nothing
	set fso = nothing
end sub

sub content()
%>
<%= LOGGER_ADDITIONAL_STYLE %>
<link rel="stylesheet" type="text/css" href="style.css">
<script language="javascript" src="javascript.js"></script>
<form name="frm" method="post" action="index.asp">
	<fieldset class="alFieldset">
		<legend class="alHeader">
			GabLIB Log Analyzer v<%=m_version%>
		</legend>
		
		<div class="alContent content">
			<table border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td nowrap>Select File:</td>
				<td nowrap><% parseLogDirectory() %></td>
				<td><% drawOptions() %></td>
				<td width="100%"><input type="Submit" value="Analyze File" class="button_common"></td>
			</tr>
			</table>
		
		</div>
		<div>
			<% drawStat() %>
		</div>
	</fieldset>
</form>
<%
end sub
%>