<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		debuggingConsole
'' @CREATOR:		Original: Microsoft - http://support.microsoft.com/default.aspx?scid=KB;EN-US;q288965
''					Formatting & extras: Mike - http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=7475&lngWId=4
''					Modified: Hunter Beanland - http://www.geocities.com/hbeanland/
''					Some modifications: Michal Gabrukiewicz - gabru @ grafix.at
'' @CREATEDON:		03.09.2003
'' @CDESCRIPTION:	It shows a complete debugging console like e.g. the console in aps.net. You will see all sessions-variables,
''					all cookies, info about database-connection, querystring-collection and form-collection and a lot more.
''					Also add variables to the panel. If you will turn off the debuggingMode they wont appear anymore.
'' @VERSION:		1.2

'**************************************************************************************************************
class debuggingConsole

	private dbg_Enabled
	private dbg_Show
	private dbg_RequestTime
	private dbg_FinishTime
	private dbg_Data
	private dbg_DB_Data
	private dbg_AllVars
	private dbg_Show_default
	private DivSets(2)
    
	'Construktor => set the default values
	Private Sub Class_Initialize()
		dbg_RequestTime = Now()
		dbg_AllVars = false
		Set dbg_Data = createObject("Scripting.Dictionary") 
		DivSets(0) = "<TR><TD style='cursor:pointer;' onclick=""javascript:if (document.getElementById('data#sectname#').style.display=='none'){document.getElementById('data#sectname#').style.display='block';}else{document.getElementById('data#sectname#').style.display='none';}""><DIV id=sect#sectname# style=""font-weight:bold;cursor:pointer;background:#7EA5D7;color:white;padding-left:4;padding-right:4;padding-bottom:2;"">|#title#|  <DIV id=data#sectname# style=""cursor:text;display:none;background:#FFFFFF;padding-left:8;"" onclick=""window.event.cancelBubble = true;"">|#data#|  </DIV>|</DIV>|"
		DivSets(1) = "<TR><TD><DIV id=sect#sectname# style=""font-weight:bold;cursor:pointer;background:#7EA5D7;color:white;padding-left:4;padding-right:4;padding-bottom:2;"" onclick=""javascript:if (document.getElementById('data#sectname#').style.display=='none'){document.getElementById('data#sectname#').style.display='block';}else{document.getElementById('data#sectname#').style.display='none';}"">|#title#|  <DIV id=data#sectname# style=""cursor:text;display:block;background:#FFFFFF;padding-left:8;"" onclick=""window.event.cancelBubble = true;"">|#data#|  </DIV>|</DIV>|"
		DivSets(2) = "<TR><TD><DIV id=sect#sectname# style=""background:#7EA5D7;color:lightsteelblue;padding-left:4;padding-right:4;padding-bottom:2;"">|#title#|  <DIV id=data#sectname# style=""display:none;background:lightsteelblue;padding-left:8"">|#data#|  </DIV>|</DIV>|"
		dbg_Show_default = "0,0,0,0,0,0,0,0,0,0,0,0"
	End Sub
	
	public property let enabled(bNewValue) ''[bool] Sets "enabled" to true or false
		dbg_Enabled = bNewValue
	End Property
	public property get enabled	''[bool] Gets the "enabled" value
		Enabled = dbg_Enabled
	End Property
	
	public property let show(bNewValue) ''[string] Sets the debugging panel. Where each digit in the string represents a debug information pane in order (11 of them). 1=open, 0=closed
		dbg_Show = bNewValue
	End Property
	public property get show ''[string] Gets the debugging panel.
		Show = dbg_Show
	End Property
	
	public property let allvars(bNewValue) ''[bool] Sets wheather all variables will be displayed or not. true/false
		dbg_AllVars = bNewValue
	End Property
	public property get allvars	''[bool] Gets if all variables will be displayed.
		AllVars = dbg_AllVars
	End Property
	
	'******************************************************************************************************************
	''@SDESCRIPTION:	Adds a variable to the debug-informations. 
	''@PARAM:			label [string]: Description of the variable
	''@PARAM:			output [variable]: The variable itself
	'******************************************************************************************************************
	public sub print(label, output)
		If dbg_Enabled Then
			if err.number > 0 then 
				call dbg_Data.Add(ValidLabel(label), "!!! Error: " & err.number & " " &  err.Description)
				err.Clear
			else
				uniqueID = ValidLabel(label)
				call dbg_Data.Add(uniqueID, output)
			end if
		End If
	End Sub
  	
	'******************************************************************************************************************
	'* ValidLabel 
	'******************************************************************************************************************
	Private Function ValidLabel(byval label)
		dim i, lbl
		i = 0
		lbl = label
		do
			if not dbg_Data.Exists(lbl) then exit do
			i = i + 1
			lbl = label & "(" & i & ")"
		loop until i = i
		
		ValidLabel = lbl
	End Function
 	
	'******************************************************************************************************************
	'* PrintCookiesInfo 
	'******************************************************************************************************************
	Private Sub PrintCookiesInfo(byval DivSetNo)
		dim tbl, cookie, key, tmp
		For Each cookie in Request.Cookies
			If Not Request.Cookies(cookie).HasKeys Then
				tbl = AddRow(tbl, cookie, Request.Cookies(cookie))    
			Else
				For Each key in Request.Cookies(cookie)
					tbl = AddRow(tbl, cookie & "(" & key & ")", Request.Cookies(cookie)(key))    
				Next
			End If
		Next
		
		tbl = MakeTable(tbl)
		if Request.Cookies.count <= 0 then DivSetNo = 2
		tmp = replace(replace(replace(DivSets(DivSetNo),"#sectname#","COOKIES"),"#title#","COOKIES"),"#data#",tbl)
		Response.Write replace(tmp,"|", vbcrlf)
	end sub
  	
	'******************************************************************************************************************
	'* PrintSummaryInfo 
	'******************************************************************************************************************
	Private Sub PrintSummaryInfo(byval DivSetNo)
		dim tmp, tbl
		tbl = AddRow(tbl, "Time of Request",dbg_RequestTime)
		tbl = AddRow(tbl, "Elapsed Time", DateDiff("s", dbg_RequestTime, dbg_FinishTime) & " seconds")
		tbl = AddRow(tbl, "Request Type", Request.ServerVariables("REQUEST_METHOD"))
		tbl = AddRow(tbl, "Status Code", Response.Status)
		tbl = AddRow(tbl, "Script Engine", ScriptEngine & " " & ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion & "." & ScriptEngineBuildVersion)
		tbl = MakeTable(tbl)
		tmp = replace(replace(replace(DivSets(DivSetNo),"#sectname#","SUMMARY"),"#title#","SUMMARY INFO"),"#data#",tbl)
		Response.Write replace(tmp,"|", vbcrlf)
	End Sub
	
	'******************************************************************************************************************
	''@SDESCRIPTION:	Adds the Database-connection object to the debug-instance. To display Database-information
	''@PARAM:			oSQLDB [object]: connection-object
	'******************************************************************************************************************
	public sub grabDatabaseInfo(byval oSQLDB)
		dbg_DB_Data = AddRow(dbg_DB_Data, "ADO Ver",oSQLDB.Version)
		dbg_DB_Data = AddRow(dbg_DB_Data, "OLEDB Ver",oSQLDB.Properties("OLE DB Version"))
		dbg_DB_Data = AddRow(dbg_DB_Data, "DBMS",oSQLDB.Properties("DBMS Name") & " Ver: " & oSQLDB.Properties("DBMS Version"))
		dbg_DB_Data = AddRow(dbg_DB_Data, "Provider",oSQLDB.Properties("Provider Name") & " Ver: " & oSQLDB.Properties("Provider Version"))
	End Sub
	
	'******************************************************************************************************************
	'* PrintDatabaseInfo 
	'******************************************************************************************************************
	Private Sub PrintDatabaseInfo(byval DivSetNo)
		dim tbl
		tbl = MakeTable(dbg_DB_Data)
		tbl = replace(replace(replace(DivSets(DivSetNo),"#sectname#","DATABASE"),"#title#","DATABASE INFO"),"#data#",tbl)
		str.write(replace(tbl,"|", vbcrlf))
	End Sub
	
	'******************************************************************************************************************
	'* PrintCollection 
	'******************************************************************************************************************
	Private Sub PrintCollection(Byval Name, ByVal Collection, ByVal DivSetNo, ByVal ExtraInfo)
		Dim vItem, tbl, Temp
		For Each vItem In Collection
			if isobject(Collection(vItem)) and Name <> "SERVER VARIABLES" and Name <> "QUERYSTRING" and Name <> "FORM" then
				if lCase(typename(Collection(vItem))) = "validateable" then
					set coll = Collection(vItem).getInvalidData()
					output = "{Validateable}"
					for each key in coll
						output = output & "<br>" & key & " = " & str.HTMLEncode(coll(key))
					next
					tbl = AddRow(tbl, vItem, output)
				else
					tbl = AddRow(tbl, vItem, "{object}")
				end if
			elseif isnull(Collection(vItem)) then
				tbl = AddRow(tbl, vItem, "{null}")
			elseif isarray(Collection(vItem)) then
				tbl = AddRow(tbl, vItem, "{array}")
			else
				if dbg_AllVars then
					tbl = AddRow(tbl, "<nobr>" & vItem & "</nobr>", server.HTMLEncode(Collection(vItem)))
				elseif (Name = "SERVER VARIABLES" and vItem <> "ALL_HTTP" and vItem <> "ALL_RAW") or Name <> "SERVER VARIABLES" then
					if Collection(vItem) <> "" then
						tbl = AddRow(tbl, vItem, server.HTMLEncode(Collection(vItem))) ' & " {" & TypeName(Collection(vItem)) & "}")
					else
						tbl = AddRow(tbl, vItem, "...")
					end if
				end if
			end if
		Next
		if ExtraInfo <> "" then tbl = tbl & "<TR><TD COLSPAN=2><HR></TR>" & ExtraInfo
		tbl = MakeTable(tbl)
		if Collection.count <= 0 then DivSetNo = 2
		tbl = replace(replace(DivSets(DivSetNo),"#title#",Name),"#data#",tbl)
		tbl = replace(tbl,"#sectname#",replace(Name," ",""))
		str.write(replace(tbl,"|", vbcrlf))
	End Sub
  	
	'******************************************************************************************************************
	'* AddRow 
	'******************************************************************************************************************
	Private Function AddRow(byval t, byval var, byval val)
		t = t & "|<TR valign=top>|<TD>|" & var & "|<TD>= " & val & "|</TR>"
		AddRow = t
	End Function
	
	'******************************************************************************************************************
	'* MakeTable 
	'******************************************************************************************************************
	Private Function MakeTable(byval tdata)
		tdata = "|<table border=0 style=""font-size:10pt;font-weight:normal;"">" + tdata + "</Table>|"
		MakeTable = tdata
	End Function
	
	'******************************************************************************************************************
	''@SDESCRIPTION:	Draws the Debug-panel
	'******************************************************************************************************************
	public sub draw()
		If dbg_Enabled Then
			dbg_FinishTime = Now()
			
			Dim DivSet, x
			DivSet = split(dbg_Show_default,",")
			dbg_Show = split(dbg_Show,",")
			
			For x = 0 to ubound(dbg_Show)
				divSet(x) = dbg_Show(x)
			Next
			
			str.write("<BR><Table width=100% cellspacing=0 border=0 style=""font-family:arial;font-size:9pt;font-weight:normal;""><TR><TD><DIV style=""background:#005A9E;color:white;padding:4;font-size:12pt;font-weight:bold;"">Debugging-console:</DIV>")
			PrintSummaryInfo(divSet(0))
			PrintCollection "VARIABLES", dbg_Data, divSet(1), ""
			PrintCollection "QUERYSTRING", Request.QueryString(), divSet(2), ""
			PrintCollection "FORM", lib.Form, divSet(3), ""
			PrintCookiesInfo(divSet(4))
			PrintCollection "SESSION", Session.Contents(),divSet(5), AddRow(AddRow(AddRow("","Locale ID", Session.LCID & " (&H" & Hex(Session.LCID) & ")"),"Code Page", Session.CodePage), "Session ID", Session.SessionID)
			PrintCollection "APPLICATION", Application.Contents(), divSet(6), ""
			PrintCollection "SERVER VARIABLES", Request.ServerVariables(), divSet(7), AddRow("","Timeout",Server.ScriptTimeout)
			PrintDatabaseInfo(divSet(8))
			PrintCollection "SESSION STATIC OBJECTS", Session.StaticObjects(),divSet(9), ""
			PrintCollection "APPLICATION STATIC OBJECTS", Application.StaticObjects(),divSet(10), ""
			PrintCollection "REGISTERED CLASSES", lib.registeredClasses, divSet(11), ""
			str.write("</Table>")
		End If
	End Sub
	
	'Destructor
	Private Sub Class_Terminate()
		Set dbg_Data = Nothing
	End Sub

End Class
%>
