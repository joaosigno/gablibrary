<!--#include file="const.asp"-->
<!--#include file="class_columnCommon.asp"-->
<!--#include file="class_columnRadioButton.asp"-->
<!--#include file="class_columnFilter.asp"-->
<%
'**************************************************************************************************************

'' @CLASSTITLE:		Drawtable
'' @CREATOR:		Michal Gabrukiewicz - gabru@gmx.at
'' @CREATEDON:		11.06.2003
'' @CDESCRIPTION:	It generates a table for a SQL-query with all your needs. Sorting, Deleting, Paging, etc.
''					It is known as Datagrid from the "big-brother" asp.net. But even better ;) This class is
''					very complex and so it should handle every problem. There are also a lot of workarounds
''					for some common problems. e.g. You can change properties of the table-object on runtime
''					by using the displayFunction from a field. Dont forget to dim the object then.
'' @VERSION:		0.99

'**************************************************************************************************************

class drawtable

	private allowFastDelete			'true = shows delete button directly in the table. for fast deleting
	private version					'version of the table class
	private absolutePage 			'needed for paging
	private showAllRecs				'needed for paging
	private sqlCommands				'dictonary-object for sql commands
	private filterToStore			'internal use
	private fullsearchtext			'
	private excelvariable			'Var to store the whole excel content in
	private excelheader				'Var to store the excel header
	private	bigForTime				'if TimeDebuging is ON, this variable is needed
	private oSB						'StringBuilder Object for excel export needed
	private allTableColumns			'Dictionary Object to store the Table-Columns
	private allFilters				'Dictionary Object to store the Filters
	private triggerValue			'If you need an auto-value for something
	private objRsFieldlinkId		'We store the FieldlinkID in an object. used to be faster!
	private output					'We store the whole output in a StringBuilder Object.
	private radioButtonsAmount		'Here we store how many radiobutton-columns we have
	private headervariable			'This variable stores the header
	
	public rs						''Recordset Object	
	public commonsort				''leave it blank then no ORDERBY will be added
	public isEditable				''false = disable record-editing
	public allowAdding				''false = dont allow to add a new record
	public fieldlinkID 				''primary key of the table. e.g. used for edit-form.
	public sqlquery					''The sql-query for the whole table. Dont forget to include every field you will need later. Dont use ORDER BY. Use "commonsort"-property instead.
	public title					''Title for the Table
	public addurl					''addurl is the url of the edit and add-form for this table.
	public isSortable				''false = column-sorting will be disabled
	public table					''This is the table from our view. its empty if there is more than 1 table
	public showFilterBar			''true/false - enable/disable filterbar
	public headersPerRow			''every which row show the headers
	public target					''The target for all links
	public recsPerPage				''How many records per page do you want today?
	public paging					''true/false - allow paging or not
	public debuging					''true/false - some debug parameters will be displayed if true
	public timeDebuging				''true/false - where do we loose the time?
	public database					''for some sql commands we need to know which DB.
	public printing					''true/false - loads the print.css and shows a print link
	public tablewidth				''the width of the table
	public hovereffect				''mouseover effect of table
	public fullsearch				''searchfield for searching the full table
	public printTxt					''Your own Text or IMAGE for the prinitng button.
	public excelExport				''true/false - You want to allow Excel exports?
	public addTxt					''Your own Text or IMAGE for the Add new recordset button.
	public excelTxt					''Your own Text or IMAGE for the Export to Excel button.
	public defaultStylesheet		''you need a default stylesheet?
	public allowHtml				''should HTML Tags be trimmed? ATTENTION! SETTING IT "FALSE" WILL CAUSE SPEED REDUCTION!!! default: true
	public restoreFilterOnSearch	''true/false - do you want to remember the seted filter when using the fulltextsearch?
	public stringBuilderDLL			''true/false - is the stringBuilder-DLL available? StringBuilderDLL makes the table extremly fast!
	public extendButtonBar			''Function-name to execute in the button-bar on drawing. Function must return the output! e.g. You need an extra button in the button-bar
	public tableCellpadding			''[int] The cellpadding of the table
	
	'Konstruktor => set the default values
	private sub Class_Initialize()
		version					= "Drawtable class - ver 0.99"
		database	 			= DEFAULTDATABASE
		triggerValue			= 0
		radioButtonsAmount		= 0
		set sqlCommands 		= Server.createObject("Scripting.Dictionary")
		set allTableColumns		= Server.createObject("Scripting.Dictionary")
		set allFilters			= Server.createObject("Scripting.Dictionary")
		commonsort 				= empty
		isEditable 				= true
		allowAdding 			= true
		fieldlinkID				= empty
		sqlquery 				= empty
		addurl 					= empty
		title 					= empty
		isSortable				= true
		allowFastDelete			= false
		timeDebuging			= false
		table					= empty
		showFilterBar			= false
		headersPerRow			= HEADERPERROWS
		recsPerPage				= RECORDSPERPAGE
		paging					= false
		target					= empty
		absolutePage 			= Request.form("actualPageNumber")
		showAllRecs				= false
		debuging 				= false
		printing				= true
		tablewidth				= "100%"
		hovereffect				= true
		printTxt				= TXTSTRPRINT
		excelExport				= false
		addTxt					= TXTADDNEWRECORD
		excelTxt				= TXTEXPORTEXCEL
		defaultStylesheet		= false
		allowHTML				= true
		restoreFilterOnSearch	= false
		excelvariable			= "" 
		excelheader				= ""
		bigForTime				= 0
		stringBuilderDLL		= true
		headervariable			= empty
		extendButtonBar			= empty
		tableCellpadding		= 3
		
		call InitializeSessionObject
	end sub
	
	'**************************************************************************************************************
	' getAbsolutePage
	'**************************************************************************************************************
	private sub getAbsolutePage()
		absolutePage 	= CInt(getSessionObject("pagenumber"))
	end sub
	
	'**************************************************************************************************************
	' getFilterValue
	'**************************************************************************************************************
	private function getFilterValue(filterValue)
		if Request.Form.Count > 0 then
			getFilterValue = Request(filterValue)
		else
			getFilterValue = getSessionObject(filterValue)
		end if
	end function
	
	'**************************************************************************************************************
	' getSortValue
	'**************************************************************************************************************
	private function getSortValue()
		if Request.Form.Count > 0 then
			getSortValue = Request("sortValue") 
		else
			mySort = getSessionObject("commonsort")
			if len(mySort) <= 1 then
				getSortValue = ""
			else
				commonSort		= mySort
				getSortValue	= mySort
			end if
		end if
	end function
	
	'**************************************************************************************************************
	' getSearchValue
	'**************************************************************************************************************
	private function getSearchValue()
		if Request.Form.Count > 0 then
			getSearchValue = request.form("fullsearchtext")
		else
			mySearch = getSessionObject("searchvalue")
			if len(mySearch) <= 1 then
				getSearchValue = ""
			else
				getSearchValue = mySearch
			end if
		end if
	end function
	
	'**************************************************************************************************************
	' getAllFilters
	'**************************************************************************************************************
	private function getAllFilters()
		if (checkSessionObject) and not (Request.Form.Count > 0) then
			set tableObject = Session("tableObject")
			if tableObject.Item("pageurl") = Request.ServerVariables("URL") then
				set dictTemp = Server.CreateObject("Scripting.Dictionary")
				for each key in tableObject.Keys
					if instr(lCase(key), "fltrfield_") then
						dictTemp.Add key, tableObject.Item(key)
					end if
'response.write key & " = " & tableObject.Item(key) & "<BR>"
				next
				set getAllFilters = dictTemp
			else
				set getAllFilters = Request.Form
			end if
			set tableObject = nothing
			set dictTemp 	= nothing
		else
			set getAllFilters = Request.Form
		end if
	end function
	
	'**************************************************************************************************************
	' getSessionObject
	'**************************************************************************************************************
	private function getSessionObject(objValue)
		if checkSessionObject then
			set tableObject = Session("tableObject")
			if tableObject.Item("pageurl") = Request.ServerVariables("URL") then
				getSessionObject = tableObject.Item(objValue)
			end if
			set tableObject = nothing
		else
			getSessionObject = 0
		end if
	end function
	
	'**************************************************************************************************************
	' checkSessionObject
	'**************************************************************************************************************
	private function checkSessionObject()
		if IsObject(Session("tableObject")) then
			checkSessionObject = true
		else
			checkSessionObject = false
		end if
	end function
	
	'**************************************************************************************************************
	' InitializeSessionObject
	'**************************************************************************************************************
	private sub InitializeSessionObject()
		if Request.Form.Count > 0 then
			set tableSessionObject = Server.CreateObject("Scripting.Dictionary")
			
			tableSessionObject.Add "pageurl", Request.ServerVariables("URL")
			tableSessionObject.Add "pagenumber", Request("actualPageNumber")
			tableSessionObject.Add "commonsort", Request("sortValue")
			tableSessionObject.Add "searchvalue", Request("fullsearchtext")
			
			for each field in Request.Form
				if (instr(lCase(field), "fltrfield_")) then
					tableSessionObject.Add field, Request(field)
				end if
			next
			
			set Session("tableObject") 	= tableSessionObject
			set tableSessionObject 		= nothing
		else
			call getAbsolutePage
		end if
	end sub
	
	'**************************************************************************************************************
	' init_StringBuilder
	'**************************************************************************************************************
	private function init_StringBuilder()
		if stringBuilderDLL then
			if excelExport then
				'we use our own stringBuilder DLL to create our String
				Set oSB = Server.CreateObject("StringBuilderVB.StringBuilder")
				'initialize the buffer with size and growth factor
				oSB.Init 40000, 7500
			end if
			Set output = Server.CreateObject("StringBuilderVB.StringBuilder")
			output.Init 40000, 7500
			Set headervariable = Server.CreateObject("StringBuilderVB.StringBuilder")
			headervariable.Init 5000,1000
		end if
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Add filter method. This is an old one. You should use addNewFilter
	'' @DESCRIPTION:	It creates the filter object itself but after you called the method. 
	'**************************************************************************************************************
	public sub addFilter(fieldName,showDropdownSQL,primaryKey,displayField,defaultSelect,fieldToMatch)
		set filterObj = new columnFilter
		with filterObj
			.fieldname = fieldname
			.showDropdownSQL = showDropdownSQL
			.primaryKey = primaryKey
			.displayField = displayField
			.defaultSelect = defaultSelect
			.fieldToMatch = fieldToMatch
		end with
		addNewFilter(filterObj)
		set filterObj = nothing
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Adds a filter to a special column.
	'' @DESCRIPTION:	It generates a dropdown field or a common text field for the column.
	''					If you need a filter for readio-button-fields just name the filter the same so the class
	''					will know that this fields go together.
	'' @PARAM:			- colObj: filter-object from class_columnFilter
	'******************************************************************************************************************
	public function addNewFilter(colObj)
		allFilters.add trigger(),colObj
	end function
	
	'**************************************************************************************************************
	' getSqlCommands 
	'**************************************************************************************************************
	private function getSqlCommands(databaseName)
		select case databaseName
			case "oracle"
				sqlCommands.Add "ucase","upper"
			case "mssql"
				sqlCommands.Add "ucase","ucase"
		end select
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Set the fast delete-function
	'' @PARAM:			- val: true/false
	'******************************************************************************************************************
	public function set_allowFastDelete(val) 'We need this method cause we have to check if the TABLE VIEW is only from 1 table
		if val then
			if not table = empty then
				allowFastDelete = true
			else
				allowFastDelete = false
			end if
		else
			allowFastDelete = false
		end if
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Add field method. This is an old one. You should use addColumn-method
	'' @DESCRIPTION:	It creates the field object itself but after you called the method. 
	'**************************************************************************************************************
	public function addfield(fieldname,displaystring,displayfunction,tdAttributes)
		set column = new columnCommon
		with column
			.fieldname = fieldname
			.displaystring = displaystring
			.displayfunction = displayfunction
			.tdAttributes = tdAttributes
		end with
		addColumn(column)
		set column = nothing
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Adds a column to your table.
	'' @PARAM:			- colObj: needs a column-instance
	'******************************************************************************************************************
	public sub addColumn(colObj)
		allTableColumns.add trigger(),colObj
		'we check if its a radiobutton and increase the radiobutton amount.
		'so later we can check if radiobuttons are available in this table
		if typename(colObj) = "columnRadioButton" then
			radioButtonsAmount = radioButtonsAmount + 1
		end if
	end sub
	
	'**************************************************************************************************************
	'* trigger 
	'**************************************************************************************************************
	private function trigger()
		triggerValue = triggerValue + 1
		trigger = triggerValue
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Add a radio Button column. Old method! Use addColumn with radiobutton-instance
	'******************************************************************************************************************
	public function addRadioBtn(fieldname,displaystring,matchID,tdAttributes)
		set radioButton = new columnRadiobutton
		with radioButton
			.fieldname = fieldname
			.displaystring = displaystring
			.value = matchID
			.tdAttributes = tdAttributes
		end with
		addColumn(radioButton)
		set radioButton = nothing
	end function
	
	'**************************************************************************************************************
	' delete_record 
	'**************************************************************************************************************
	private function delete_record()
		if allowFastDelete then
			delid = empty
			'we get the delID from the hiddenfield
			delid = request.form("fastdeleteID")
			if not delid = empty then 'if there is a delid then we delete
				delsql = "DELETE FROM " & table & " WHERE " & fieldlinkID & " = " & delid
				lib.getrecordset(delsql)
			end if
		end if
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Draws the whole table.
	'******************************************************************************************************************
	public function draw()
		
		'FIRST WE INIT THE STRINGBUILDER
		init_StringBuilder()
		
		'WE EXECUTE DELETE FUNCTION
		delete_record()
		
		call TableHeader(title,addUrl)
		call genTable(sqlQuery,fieldLinkId,addUrl,commonSort)
		
		'IF WE HAVE STRINGBUILDER THEN WE SHOW THE OUTPUT
		if stringBuilderDLL then
			response.write output.toString
			set output = nothing
			set headervariable = nothing
			set oSB = nothing
		end if
		
		set allFilters = nothing
		set allTableColumns = nothing
		
		if err <> 0 then
			set myErroro = new errorHandler
			with myErroro
				.errorDuring = "drawing table (drawtable-class)"
				.debuggingVar = helpVar
				.errorObject(err)
				.draw
			end with
			set myErroro = nothing
			response.end
		end if
	end function
	
	'**************************************************************************************************************
	' TableHeader
	'**************************************************************************************************************
	private sub TableHeader(title,addUrl)
		%><!--#include file="js.asp"--><%
		'We load the default css
		if defaultStylesheet then
			addToOutput ("<link rel=stylesheet type=text/css href='" & CLASSLOCATION & "standard.css'>")
		end if
		if printing then
			printButton = "&nbsp;<BUTTON onclick=""javascript:window.print();"" class=button>" & printTxt & "</BUTTON>"
			addToOutput ("<LINK type=""text/css"" href=""" & CLASSLOCATION & "print.css"" rel=""stylesheet"" media=""print"">") 'css for printing
		else
			printButton = empty
		end if
		addToOutput ("<TABLE border=0 cellpadding=0 cellspacing=0 width=100% >" & vbcrlf)
		addToOutput ("<TR><TD><FORM name=myVeryHiddenForm method=post action=""" & request.serverVariables("URL") & "?" & lib.getAllFromQueryStringBut("fastDeleteID") & """>")
		addToOutput ("<TABLE border=0 cellpadding=2 cellspacing=0 width=" & tablewidth & " >" & vbcrlf)
		addToOutput ("<TR class=headlineBG>" & vbcrlf)
		if not title = empty then
			addToOutput ("	<TD class=headline>" & title & "</TD>" & vbcrlf)
		end if
		addToOutput ("	<TD align=right>" & vbcrlf)
		addToOutput ("		<span class=noprint>" & vbcrlf)
		addToOutput ("		<TABLE border=0 cellpadding=0 cellspacing=0>" & vbcrlf)
		addToOutput ("		<TR>" & vbcrlf)
		
		'our fulltextsearch field
		if fullsearch then
			addToOutput ("<TD>" & vbcrlf)
			call displayFullSearchField()
			addToOutput ("</TD>" & vbcrlf)
		end if
		'excelexport button.
		displayexcelexport
		'print button
		if not printButton = empty then
			addToOutput ("<TD>" & printButton & "</TD>")
		end if
		'add button
		if not addUrl = "" and allowAdding then
			myLink = getAddLink(addUrl)
			addToOutput ("<TD>&nbsp;" & myLink & "</TD>" & vbcrlf)
		end if
		
		'additional buttons. Function/sub will be executed
		if not extendButtonBar = empty then
			execute("additionalFunction = " & extendButtonBar)
			addToOutput ("<TD>&nbsp;" & additionalFunction & "</TD>" & vbcrlf)
		end if
		
		addToOutput ("		</TR>" & vbcrlf)
		addToOutput ("		</TABLE>" & vbcrlf)
		addToOutput ("		</span>" & vbcrlf)
		addToOutput ("	</TD>" & vbcrlf)
		addToOutput ("</TR>" & vbcrlf)
		addToOutput ("</TABLE>" & vbcrlf)
		addToOutput ("</TD>" & vbcrlf)
		addToOutput ("</TR>" & vbcrlf)
		addToOutput ("</TABLE>" & vbcrlf)
	end sub

	'******************************************************************************************************************
	' HERE WE CREATE THE HEADLINES FOR EVERY TABLE. IF commonSort = empty THEN SORTING IS DISABLED!
	'******************************************************************************************************************
	private sub showHeader(myFields,commonSort)
		myAhref = empty 'it stays empty if we dont want orderlinks
		addToHeaderVariable ("<TR align=left>" & vbcrlf)
		
		mySort = request.form("sortValue")
		if mySort = "" then mySort = commonSort end if
		
		addToExcelHeader(vbcrlf & "<TR>")
		
		for each col in allTableColumns.items
			showCurrentSort = empty
			'if table is sorted by this col we set sort DESC to enable sorting other way!
			if instr(mySort,"DESC") then
				tmpArr = split(mySort,"XXYYZZ")
				if tmpArr(0) = col.fieldName then
					showCurrentSort = "&nbsp;&nbsp;\/"
				end if
			end if
			if mySort = col.fieldName then
				showCurrentSort = "&nbsp;&nbsp;/\"
				sortby = col.fieldName & "XXYYZZ" & "DESC" 'we cannot give a Space with querystring so we save it as + and replace it later
			else
				sortby = col.fieldName
			end if
			
			'if we allow sorting then we make the link
			if isSortable then
				myAhref = "<a class='head' href=""javascript:document.myVeryHiddenForm.sortValue.value='" & sortby & "';document.myVeryHiddenForm.submit();"">"
			end if
			
			addToHeaderVariable ("<TD class='head' " & col.headerTdAttributes & " nowrap><label title=""" & TOOLTIP_TXT_SORTBY & " " & col.displayString & """>" & myAhref & col.displayString & "</a></label><nobr><span class=currentSort>" & showCurrentSort & "</span></nobr></TD>" & vbcrlf)
			addToExcelHeader(vbcrlf & "<TD><B>" & col.displayString & "</B></TD>" & vbcrlf)
		next
		
		addToExcelHeader("</TR>" & vbcrlf)
		if allowFastDelete OR showFilterBar then
			addToHeaderVariable ("<TD class='head'>&nbsp;</TD>" & vbcrlf)
		end if
		addToHeaderVariable ("</TR>" & vbcrlf)
	end sub
	
	'*************************************************************************************************************
	' filterBar 
	'*************************************************************************************************************
	private sub filterBar(myFieldArr)
		
		if not allFilters.count = 0 AND showFilterBar then 'we dont show anything if there are no filters defined
			addToOutput ("<TR class=filterBG>")
			
			filterToStore = empty 'we need this to remember the filterfield if there will be an update of rbform
			for each col in allTableColumns.items
				matched = "no"
				sql = ""
				for each fil in allFilters.items
					if col.fieldName = fil.fieldName then
						matched 	= fil.fieldName
						sql 		= fil.showDropdownSQL
						myPK		= fil.primaryKey
						showName	= fil.displayField
						attributes	= fil.tdAttributes
						myVal 		= getFilterValue("fltrField_" & matched) 'request.form("fltrField_" & matched)
						'if the form was submitted we select the value of the dropdown
						if fil.description = empty then
							commonValue = TXTCOMMONFILTER
						else
							commonValue = fil.description
						end if
						if myVal	= empty then
							myVal	= fil.defaultSelect
						end if
					end if
				next
				
				'check if the field is a Radiobutton field, so we need a colspan
				if radioButtonsAmount = 0 then
					colspan = empty
				else
					if not lastFilter = matched then
						if typename(col) = "columnRadioButton" then
							colspan = " colspan=" & radioButtonsAmount
						else
							colspan = empty
						end if
					else
						colspan = empty
					end if
				end if
				
				if not matched = "no" AND not lastFilter = matched then
					addToOutput ("<TD" & colspan & " " & attributes & "><span class=noprint>")
					if sql = "" then 'we make a common text filter field
						tooltipText = replace(TOOLTIP_TXT_FILTER_TEXTFIELD,"{fieldname}",col.displayString)
						addToOutput ("<LABEL title=""" & tooltipText & """><input type=text value='" & myVal & "' class=formField name=fltrField_" & matched & " " & onChangeSubmit() & " ></LABEL>" & vbcrlf)
					else 'we must create a dropdown
						set dd = new createDropdown
						
						'now we build the dropdown
						with dd
							.addDisplayField (showName)
							.pk 				= myPK
							.sqlquery			= sql
							.name				= "fltrField_" & matched
							.idToMatch			= myVal
							.commonTxt			= commonValue
							.commonTxtVal		= "XXXYYYZZZXXX"
							.onAttribute		= onChangeSubmit()
							addToOutput(.getAsString)
						end with
						set dd = nothing
					end if
					'we store the field
					filterToStore = filterToStore & "<input type=hidden value='" & myVal & "' class=formField name=fltrField_" & matched & ">" & vbcrlf
					addToOutput ("</span></TD>" & vbcrlf)
					lastFilter = matched
				else
					if not lastFilter = matched then
						addToOutput ("<TD" & colspan & " " & attributes & "></TD>" & vbcrlf)
					end if
				end if
			next
			addToOutput ("<TD align=center nowrap valign=top>" & vbcrlf)
			addToOutput ("		<input type=submit name=go class=button_common style='width:0px;height:0px;'><span class=noprint><label title=""" & TOOLTIP_TXT_FILTERRESET & """><button style='height:12px;cursor:hand;' onclick=""javascript:restoreMyFilter();myVeryHiddenForm.submit();"" name=restore_my_filter class=button_common>^</button></label></span>" & vbcrlf)
			addToOutput ("</TD>" & vbcrlf)
			addToOutput ("</TR>" & vbcrlf)
		end if
	end sub
	
	'******************************************************************************************************************
	' HERE WE GET AND ADD NEW RECORDSET LINK 
	'******************************************************************************************************************
	private function getAddLink(addUrl)
		if not target = empty then
			myTarget = "parent." & target & "."
		else
			myTarget = empty
		end if
		getAddlink = "<BUTTON class=button onclick=""javascript:" & myTarget & "location.href='" & addUrl & "'"">" & addTxt & "</BUTTON>"
	end function
	
	'******************************************************************************************************************
	' THATS THE BASE PROCEDURE WHICH GENERATES THE WHOLE TABLE WITH THE RECORDS 
	'******************************************************************************************************************
	private sub genTable(sqlQuery,fieldLinkId,addUrl,commonSort)
		'first we get our sqlcommands. depend on databse.
		getSqlCommands(database)
		
		if not addUrl = "" and allowAdding then
			myLink = getAddLink(addUrl)
		end if
		
		sql = sqlQuery
		
		if showFilterBar then
			'if there is a where in our sql statement then its okay. else we add a where 1=1 ourself
			if instr(ucase(sql),"WHERE") = 0 then
				sql = sql & " WHERE 1=1"
			end if
		end if
		
		'mysort = request.form("sortValue")
		mysort = getSortValue()  'get the sortValue to know what column to sort
		if not mysort = "" then 'add the sort
			if instr(mysort,"XXYYZZ") then
				sort = replace(mySort,"XXYYZZ",") ")
				sort1 = left(mysort,(len(mysort)-10))
			else
				sort = mySort & ")"
				sort1 = mysort
			end if
			
			ucased = sqlCommands.Item("ucase")
			for each col in allTableColumns.items
				if cstr(sort1) = cstr(col.fieldname) AND col.isNumber then
					ucased = empty
				end if
			next
			sql = sql & " ORDER BY " & ucased & "(" & sort
		else
			'********************************************
			' WE HAVE A COMMONSORT 
			'********************************************
			if not commonSort = "" then
				'if we have more fields to sort
				select case database
					case "oracle"
						if instr(commonSort,",") then
							mySorts = split(commonSort,",")
							for o = 0 to ubound(mySorts)
								if instr(ucase(mySorts(o)),"DESC") = 0 then
									sOrd = sOrd & sqlCommands.Item("ucase") & "(" & mySorts(o) & "),"
								else
									sOrd = sOrd & sqlCommands.Item("ucase") & "(" & left(mySorts(o),(len(mySorts(o))-5)) & ") DESC,"
								end if
							next
							sql = sql & " ORDER BY " & left(sOrd,(len(sOrd)-1))
						else
							if instr(ucase(commonSort),"DESC") then
								sql = sql & " ORDER BY " & sqlCommands.Item("ucase") & "(" & left(commonSort,(len(commonSort)-5)) & ") DESC"
							else
								sql = sql & " ORDER BY " & sqlCommands.Item("ucase") & "(" & commonSort & ")"
							end if
						end if
					case "mssql"
						sql = sql & " ORDER BY " & commonSort
				end select
			end if
		end if
		
		'**********************************************************************************************************************
		' USE OUR FILTER 
		'**********************************************************************************************************************
		if showFilterBar then
			splittedSQL			= split(ucase(sql),"WHERE") 'we split the sql to add the filter condition. we make sql ucase so we be sure to find the where
			atLeastOneCriteria	= false
			
			'WE CHECK IF TERE ARE SOME FILTER SELECTED
			if Request.Form.Count > 0 then
				set myFilters = request.form
			else
				set myFilters = getAllFilters
			end if
				for each fild in myFilters 'for each fild in request.form
				myFieldName = empty
				if instr(fild,"fltrField_") and not getFilterValue(fild) = empty and not getFilterValue(fild) = "XXXYYYZZZXXX" then
					atLeastOneCriteria = true
					myFieldName = right(fild,(len(fild)-10)) 'we cut the fltrField_ thing to get the name
					for each fil in allFilters.items
						if cstr(myFieldName) = cstr(fil.fieldName) then
							'We check if there is a special fieldname to compare with. if not we take the one which we displayed
							if not fil.fieldToMatch = empty then
								matchField = fil.fieldToMatch
							else
								matchField = fil.fieldName
							end if
							
							'searchKeyword = ltrim(request.form(fild)) 'the value we search for
							searchKeyword = ltrim(getFilterValue(fild))
							operatorFound = false
							
							'we go through our array with allowed filter operators
							for p = 0 to ubound(ALLOWED_FILTER_OPERATORS)
								if left(searchKeyword,len(ALLOWED_FILTER_OPERATORS(p))) = ALLOWED_FILTER_OPERATORS(p) then
									operator = ALLOWED_FILTER_OPERATORS(p)
									searchKeyword = trim(right(searchKeyword,(len(searchKeyword)-len(operator))))
									operatorFound = true
									exit for
								end if
							next
							
							'If we have a custom comma-style we need to replace it
							if not fil.commaStyle = empty then
								tmp = replace(searchKeyword,fil.commaStyle,".")
								searchKeyword = tmp
							end if
							
							if operatorFound then 'we check if there was an operator found. if yes then we execute the sql with the operator
								splittedSQL(1) = " (" & matchField & ") " & operator & " ('" & searchKeyword & "') AND " & splittedSQL(1)
							else 'we execute the sql with pre-defined statements. = and LIKE
								'we have no sql so its a common text field. so we need a like
								if fil.showDropdownSQL = empty then
									splittedSQL(1) = " " & sqlCommands.Item("ucase") & "(" & matchField & ") LIKE " & sqlCommands.Item("ucase") & "('%" & searchKeyword & "%') AND " & splittedSQL(1)
								else
									splittedSQL(1) = " (" & matchField & " = '" & searchKeyword & "') AND " & splittedSQL(1)
								end if
							end if
						end if
					next
				end if
			next
			
			set fil = nothing
			
			'WE CHECK IF THERE ARE DEFAULT VALUES FOR A FILTER
			if not atLeastOneCriteria then
				for each fil in allFilters.items
					if not cstr(fil.defaultSelect) = empty and request.form("fltrField_" & fil.fieldName) = empty then
						atLeastOneCriteria = true
						'We check if there is a special fieldname to compare with. if not we take the one which we displayed
						if not fil.fieldToMatch = empty then
							matchField = fil.fieldToMatch
						else
							matchField = fil.fieldName
						end if
						'we have no sql so its a common text field. so we need a like
						if fil.showDropdownSQL = empty then
							splittedSQL(1) = " UPPER(" & matchField & ") LIKE UPPER('%" & fil.defaultSelect & "%') AND " & splittedSQL(1)
						else
							splittedSQL(1) = " " & matchField & " = '" & fil.defaultSelect & "' AND " & splittedSQL(1)
						end if
					end if
				next
			end if
			
			if atLeastOneCriteria then
				sql = splittedSQL(0) & " WHERE " & splittedSQL(1)
				'if the sql has more WHERE statements than one we add all. The is always set after the first WHERE
				if ubound(splittedSQL) > 1 then
					for ij = 2 to ubound(splittedSQL)
						sql = sql & " WHERE " & splittedSQL(ij)
					next
				end if
			end if
		end if
		
		if debuging then
			addToOutput ("<strong>SQL-Query: </strong>" & sql & "<BR>")
		end if
		
		'*******************************************************************
		' PAGING 
		'*******************************************************************
		
		if paging then
			if timedebuging then 
				startRecordSet = timer()
			end if
			
			set rs = lib.getUnlockedRecordset(sql)
			
			if timedebuging then
				EndRecordSet = timer()
				addToOutput ("unlocked recordset received in : " & EndRecordSet - startRecordSet & "<br>")
			end if
			
			fullsearchroutine(rs)
			
			if not rs.eof then
				if absolutePage = empty then
					absolutePage = 1
				end if
				
				rs.PageSize 	= recsPerPage
				rs.CacheSize 	= recsPerPage
				
				if absolutePage = 0 then
					showAllRecs = true
					rs.absolutePage = 1
				end if
				
				If absolutePage = "" or not isNumeric(absolutePage) or cint(absolutePage) > cint(rs.PageCount) then
					showAllRecs 	= false
					absolutePage 	= 1
					rs.absolutePage = absolutePage
				else
					if not absolutePage = 0 then
						rs.absolutePage = absolutePage
					end if
				end if
			end if
		else
			set rs = lib.getUnlockedRecordset(sql)
			fullsearchroutine(rs)
		end if
		
		set sqlCommands 	= nothing
		counter 			= 0
		sumCount			= 0
		
		addToOutput ("<TABLE width=" & tablewidth & " cellspacing=1 cellpadding=" & tableCellpadding & " border=0 class=""tableClass"">" & vbcrlf)
		'show header first time. if nothing found also headers will be shown
		call showHeader(myFields,commonSort)
		if stringBuilderDll then
			addToOutput (headerVariable.toString)
		else
			addToOutput (headervariable)
		end if
		
		'if debuging then we make the hiddenfields visible
		if debuging then
			hiddenYesNo 	= "text"
			debugTXTsort 	= "Sorting value: "
			debugTXTpnr		= "Actual Pagenumber: "
			debugTXTfdel	= "Fast Delete ID: "
			debugTXTffll	= "Fulltextsearch: "
		else
			hiddenYesNo 	= "hidden"
			debugTXTsort 	= empty
			debugTXTpnr		= empty
			debugTXTfdel	= empty
		end if
		
		addToOutput (debugTXTsort & "<INPUT type=" & hiddenYesNo & " name=sortValue value='" & mysort & "'>") 'we write the current sort as value.
		addToOutput (debugTXTpnr & "<INPUT type=" & hiddenYesNo & " name=actualPageNumber value='" & absolutePage & "'>") 'we save the current page number to
		
		'show us the cool filter bar
		call filterBar(myFields)
		
		if allowFastDelete then
			addToOutput (debugTXTfdel & "<INPUT type=" & hiddenYesNo & " name=fastDeleteID value=''>") 'if we want to delete a record with fastDelete then we write it in here
		end if
		
		if not rs.eof then
			
			'**********************************************************************************************************
			' WE GOT THROUGH THE RECORDS 
			'**********************************************************************************************************
			
			if timedebuging then
				bigWhileStart = timer()
			end if
			
			set objRsFieldlinkId = rs(fieldlinkid) 'our primary-key in an object. should be faster
			
			while not rs.eof
				if (paging AND counter < recsPerPage) OR not paging OR showAllRecs then
					
					addToExcelVariable(vbcrlf & "<tr>" & vbcrlf)
					
					'SHOW US THE HEADERS. EVERY X LINE.
					if not counter = 0 AND counter mod headersPerRow = 0 then
						if stringBuilderDll then
							addToOutput (headerVariable.toString)
						else
							addToOutput (headervariable)
						end if
						'call showHeader(myFields,commonSort)
					end if
					
					if counter = 0 then
						addToOutput ("</FORM>" & vbcrlf)
						addToOutput ("<FORM name=rbfrm method=post action=""" & request.servervariables("URL") & "?" & lib.getallFromQuerystringBut("fastDeleteID") & """>" & vbcrlf)
						addToOutput ("<INPUT type=hidden name=strMatrix size=60 value=''>" & vbcrlf)
						addToOutput ("<INPUT type=hidden name=save value=true>" & vbcrlf)
						'her we have a hidden input text to allow the programmer to store something
						addToOutput ("<INPUT type=hidden name=valueToStore value=''>" & vbcrlf)
						'here we save some values if form will be updated. to remind the settings
						addToOutput ("<INPUT type=hidden name=sortValue value='" & mysort & "'>" & vbcrlf)
						addToOutput ("<INPUT type=hidden name=actualPageNumber value='" & absolutePage & "'>" & vbcrlf)
						addToOutput (filterToStore) 'here we write our stored filter fields
					end if
					
					mouseover = lib.showMouseover(counter,hovereffect)
					
					addToOutput ("<TR " & mouseover & " id=tr_" & objRsFieldlinkId & ">" & vbcrlf)
					'we display the id of the field if debuging is on
					if debuging then
						myDebugID = "(id: " & objRsFieldlinkId & ")&nbsp;"
					end if
					
					if timedebuging then
						bigForStart = timer()
					end if
					
					'GO THROUGH ALL COLUMNS
					showTableData(myDebugID)
					
					if timedebuging then
						bigForEnd = timer()
						bigForTime = bigForTime + (bigForEnd - bigForStart)
						bigForEnd = 0
						bigForStart = 0
					end if
					
					'****************************************************************************
					' FAST DELETE BUTTON 
					'****************************************************************************
					if allowFastDelete OR showFilterBar then
						if allowFastDelete then
								addToOutput ("<TD align='right' valign=middle width=15><span class=noprint>" & vbcrlf)
								addToOutput ("<label title=""" & TOOLTIP_TXT_FASTDELETE & """><button name='XXXYYYZZZ' onclick=""javascript:yesNoFastDelete(")
								addToOutput (objRsFieldlinkId)
								addToOutput (",'" & TXTFASTDELETEQUESTION & "')"" class=button_common>" & TXTFASTDELETEBUTTON & "</button></label>")
								addToOutput ("</span></TD>" & vbcrlf)
						elseif showFilterBar then
							addToOutput ("<TD width=15></TD>" & vbcrlf)
						end if
					end if
					addToOutput ("</TR>" & vbcrlf)
					counter = counter + 1
					
					addToExcelVariable("</tr>" & vbcrlf)
					
				end if
				sumCount = sumCount + 1
				rs.movenext	
			wend
			
			if timedebuging then
				bigWhileEnd = timer()
				addToOutput ("big for needed: " & bigForTime & "<br>")
				addToOutput ("big while needed: " & bigWhileEnd - bigWhileStart & "<br>")
			end if
		
		end if
		addToOutput ("</form>" & vbcrlf)
		addToExcelVariable(vbcrlf & "</table>" & vbcrlf)
		
		if paging then
			if not counter = 0 then
				pageCount 	= rs.PageCount
				'we calculate the sum of all records if paging is displayed
				if not absolutePage = 1 AND not showAllRecs then
					sumRecs	= sumCount + (recsPerPage * (absolutePage - 1))
				else
					if showAllRecs then
						sumRecs = counter
					else
						sumRecs = sumCount
					end if
				end if
			else
				absolutePage = 1
			end if
		else
			sumRecs 	= 0
			pageCount 	= 0
		end if
		set rs = nothing
		call TableFooter(counter,myLink,pageCount,sumRecs)
		
		call createExcelForm()
		set objRsFieldlinkId = nothing
	end sub
	
	'******************************************************************************************************************
	' PAGING BAR 
	'******************************************************************************************************************
	private sub PagingBar(pageCount,counter)
		'show me the prev link
		if not absolutePage = 1 and not counter = 0 and not showAllRecs then
			addToOutput ("<label title=""" & TOOLTIP_TXT_PREVPAGE & """><a href=""javascript:document.myVeryHiddenForm.actualPageNumber.value='" & absolutePage - 1 & "';document.myVeryHiddenForm.submit();"" class=pagingPrevNext>" & TXTPAGINGPREV & "</a></label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
		end if
		
		'show me the all link
		if not showAllRecs then
			addToOutput ("<label title=""" & TOOLTIP_TXT_SHOWALL & """><a href=""javascript:document.myVeryHiddenForm.actualPageNumber.value='0';document.myVeryHiddenForm.submit();"" class=pagingLink>" & TXTPAGINGALL & "</a></label>")
		else
			addToOutput ("<span class=pagingCurrentPage>" & TXTPAGINGALL & "</span>")
		end if
		addToOutput ("&nbsp;&nbsp;&nbsp;&nbsp;")
		
		'show me the page links
		for intPageCounter = 1 to pageCount
			'dont link actual page
			if CInt(intPageCounter) = CInt(absolutePage) then
			    addToOutput ("<span class=pagingCurrentPage>" & intPageCounter & TXTPAGINGSEPERATOR & "</span>")
		    else
				if abs(intPageCounter-absolutePage) < PAGING_AMOUNT_OF_NUMBERS then
		        	addToOutput ("<a href=""javascript:document.myVeryHiddenForm.actualPageNumber.value='" & intPageCounter & "';document.myVeryHiddenForm.submit();"" class=pagingLink>" & intPageCounter & TXTPAGINGSEPERATOR & "</a>")
				end if
			end if
		next
		
		'show me the next link
		if not cint(absolutePage) = cint(pageCount) and not counter = 0 and not showAllRecs then
			addToOutput ("&nbsp;&nbsp;&nbsp;&nbsp;<label title=""" & TOOLTIP_TXT_NEXTPAGE & """><a href=""javascript:document.myVeryHiddenForm.actualPageNumber.value='" & absolutePage + 1 & "';document.myVeryHiddenForm.submit();"" class=pagingPrevNext>" & TXTPAGINGNEXT & "</a></label>")
		end if
	end sub
	
	'******************************************************************************************************************
	' OUR TABLE FOOTER. INCL. NUMBER OF RECORDS AND ADD NEW RECORD LINK 
	'******************************************************************************************************************
	private sub TableFooter(counter,myLink,pageCount,sumRecs)
		if counter > 1 OR counter = 0 then records = " " & TXTRECORDSPLURAL else records = " " & TXTRECORDSSINGULAR end if 'plural or singular!
		if not sumRecs = 0 then sumRecsTXT = " / " & sumRecs else sumRecsTXT = empty end if
		if counter = 0 then
			'we make a fake record line to let the design be the same
			addToOutput ("<TR>" & vbcrlf)
			for each col in allTableColumns.items
				addToOutput ("<TD " & col.tdAttributes & "></TD>" & vbcrlf)
			next
			if allowFastDelete then
				addToOutput ("<TD width=15></TD>" & vbcrlf)
			end if
			addToOutput ("</TR>" & vbcrlf)
			'and now the real message that there is nothing
			addToOutput ("<TR><TD align=center colspan=30><div class=noRecsFound>" & TXTNORECSAVAILABLE & "</div></TD></TR>" & vbcrlf)
		end if
		addToOutput ("</TABLE>" & vbcrlf)
		addToOutput ("	<TABLE border=0 width=100% cellpadding=" & tableCellpadding & " cellspacing=0 class=tableFooterBG><TR>" & vbcrlf)
		addToOutput ("		<TD nowrap valign=top><span class=recordsDisplayed>" & counter & sumRecsTXT & " " & records & ".</span></TD>" & vbcrlf)
		
		if paging then
			addToOutput ("<TD width=100% align=right valign=top><span class=noprint>")
			call PagingBar(pageCount,counter)
			addToOutput ("&nbsp;&nbsp;</span></TD>" & vbcrlf)
		end if
		
		addToOutput ("	</TR></TABLE>" & vbcrlf)
	end sub
	
	'******************************************************************************************************************
	' FULLTEXT SEARCH
	'******************************************************************************************************************
	private sub displayFullSearchField()
		if restoreFilterOnSearch then
			mySubmit = " onclick=""javascript:myVeryHiddenForm.submit();"""
			onkeyPress = empty
		else
			mySubmit = " onclick=""javascript:restoreMyFilter();myVeryHiddenForm.submit();"""
			onkeyPress = " onkeypress=""javascript:checkChar(event);"""
		end if
		addToOutput ("<TABLE border=0 cellpadding=0 cellspacing=0 width='100%'>" & vbcrlf)
		addToOutput ("<TR>" & vbcrlf)
		addToOutput ("	<TD valign=middle nowrap>" & TXTSEARCHALL & "</TD>" & vbcrlf)
		addToOutput ("	<TD valign=middle width=100% >&nbsp;<LABEL title=""" & TOOLTIP_TXT_SEARCHALL & """><input class=formField type=text name=fullsearchtext value=""" & request.form("fullsearchtext") & """ " & onkeyPress & " style=""width:100px;""></LABEL></TD>" & vbcrlf)
		addToOutput ("	<TD valign=middle nowrap>&nbsp;<input type=button name=button class=button value=search" & mySubmit & "></TD>" & vbcrlf)
		addToOutput ("</TR>" & vbcrlf)
		addToOutput ("</TABLE>" & vbcrlf)
		addToOutput ("<SCRIPT language=JavaScript>myVeryHiddenForm.fullsearchtext.focus();</SCRIPT>")
	end sub
	
	'******************************************************************************************************************
	' FULLSEARCHROUTINE 
	'******************************************************************************************************************
	private sub fullsearchroutine(rs)
		'fullsearchtext = request.form("fullsearchtext")
		fullsearchtext = getSearchValue()
		if not rs.eof then
			if fullsearch and fullsearchtext <> "" then
				filterstring = ""
				for each x in rs.fields
					if (x.type = 200) or (x.type = 201) or (x.type = 202) or (x.type = 203) then 'removed " or (x.type = 135)" because date-fields make troubles
						if filterstring = "" then 
							filterstring = x.name & " like '*" & fullsearchtext & "*'"
						else
							filterstring = filterstring & " or " & x.name & " like '*" & fullsearchtext & "*'"	
						end if
					end if
				next
				if debuging then
					response.write filterstring
				end if
				rs.filter = filterstring
			end if
		end if
	end sub

	'******************************************************************************************************************
	' DISPLAY EXCEL EXPORT
	'******************************************************************************************************************
	private sub displayExcelExport()
		if excelexport then
			addToOutput ("<TD>")
			addToOutput ("&nbsp;<button onClick=""document.forms['xlform'].submit();"" class='button' name='exportto'>" & excelTxt & "</button>" & vbcrlf)
			addToOutput ("</TD>")
		end if
	end sub

	'******************************************************************************************************************
	' CREATE EXCELFORM
	'******************************************************************************************************************
	private sub createExcelForm()
		if excelexport then
			if stringBuilderDLL then
				excelContent = "<TABLE border=1>" & vbcrlf & excelheader & osb.toString() & vbcrlf & "</TABLE>"
			else
				excelContent = "<TABLE border=1>" & vbcrlf & excelheader & excelvariable & vbcrlf & "</TABLE>"
			end if
			addToOutput ("<form name='xlform' method='post' action='/gab_Library/class_excelexport/displayexcel.asp'>" & vbcrlf)
			
			dim xltextLength : xltextLength = len(excelContent)
			dim stepValue : stepValue		= 100000
			
			dim midstring
			
			dim i : i = 1
			
			while i < xltextLength
				midstring = trim(mid(excelContent,i,stepValue))
				addToOutput ("	<input type='hidden' name='xltext"&i&"' value='"& midstring &"'>" & vbcrlf)
				i = i + stepValue
			wend
			addToOutput ("</form>" & vbcrlf)
		end if
	end sub
	
	'******************************************************************************************************************
	' DO EXCEL EXPORT
	'******************************************************************************************************************
	private sub addToExcelVariable(str)
		if excelexport then
			if stringBuilderDLL then
				oSB.Append str
			else
				excelvariable = excelvariable & str
			end if
		end if
	end sub
	
	private sub addToExcelHeader(string)
		if excelexport then	
			'only insert the first line as header, if a </TR>-Tag is existing, we already added a header
			if instr(ucase(excelheader),"</TR>") = 0 then
				excelheader = excelheader & string
			end if
		end if
	end sub
	
	'******************************************************************************************************************
	' ONCHANGE SUBMIT 
	'******************************************************************************************************************
	private function onChangeSubmit()
			onChangeSubmit = " onChange = 'document.myVeryHiddenForm.submit();'"
	end function
	
	'******************************************************************************************************************
	' showTableData 
	'******************************************************************************************************************
	private sub showTableData(myDebugID)
		for each col in allTableColumns.items
			if col.colType = "radiobutton" then
				if cint(col.value) = cint(rs(col.fieldname)) then
					checked = " checked"
					addToExcelVariable("<td>X</td>")
				else
					checked = ""
					addToExcelVariable("<td></td>")
				end if
				addToOutput("	<TD align=center " & col.tdAttributes & " ><input type=""radio"" name=rb_" & objRsFieldlinkId & " value=" & col.value & checked & " onclick=javascript:changeBG(rb_" & objRsFieldlinkId & "," & objRsFieldlinkId & ");" & "></TD>" & vbcrlf)
			else
				'WE EXECUTE THE DISPLAYFUNCTION IF NEEDED
				if col.displayFunction <> "" then
					if not allowHTML then
						displayfunction = "field0 = " & col.displayFunction & "(""" & lib.TrimHTML(rs(col.fieldName)) & """)"
					else
						displayfunction = "field0 = " & col.displayFunction & "(""" & rs(col.fieldName) & """)"
					end if
					execute(displayfunction)
				else
					if not allowHTML then
						field0 = lib.TrimHTML(rs(col.fieldName))
					else
						field0 = rs(col.fieldName)
					end if
				end if
				
				if isEditable then
					if instr(addUrl,"?") then qsPrefix = "&" else qsPrefix = "?" end if 'we check if the addurl has already some Querystring parameters
					if not target = empty then myTarget = " parent." & target & "." else myTarget = empty end if
					editUrlBegin = "<a class=""navklein"">"
					editUrlEnd = "</a>"
					tdRowJavascript = " onclick=""" & myTarget & "location.href='" & addUrl & qsPrefix & fieldLinkid & "=" & objRsFieldlinkId & "'"" style='cursor:hand;'"
				else
					editUrlBegin = empty
					editUrlEnd = empty
					tdRowJavascript = empty
				end if
				
				addToOutput("	<TD " & col.tdAttributes & tdRowJavascript & ">" & myDebugID & editUrlBegin & field0 & editUrlEnd & "</TD>" & vbcrlf)
				addToExcelvariable("<TD>" & field0 & "</TD>" & vbcrlf)
			end if
		next
	end sub
	
	'******************************************************************************************************************
	' addToOutput 
	'******************************************************************************************************************
	private function addToOutput(str)
		if stringBuilderDLL then
			output.append str
		else
			response.write str
		end if
	end function
	
	'******************************************************************************************************************
	' addToHeaderVariable 
	'******************************************************************************************************************
	private function addToHeaderVariable(str)
		if stringBuilderDLL then
			headerVariable.append str
		else
			headerVariable = headerVariable & str
		end if
	end function
	
end class
%>
