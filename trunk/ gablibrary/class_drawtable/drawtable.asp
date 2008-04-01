<!--#include virtual="/gab_LibraryConfig/_drawtable.asp"-->
<!--#include file="config.asp"-->
<!--#include file="class_tableRow.asp"-->
<!--#include file="class_tableLegend.asp"-->
<!--#include file="class_columnCommon.asp"-->
<!--#include file="class_columnRadioButton.asp"-->
<!--#include file="class_columnFilter.asp"-->
<!--#include file="class_tableSession.asp"-->
<!--#include virtual="/gab_Library/class_excelExporter/excelExporter.asp"-->
<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		Drawtable
'' @CREATOR:		Michal Gabrukiewicz - gabru @ grafix.at
'' @CREATEDON:		11.06.2003
'' @CDESCRIPTION:	It generates a table for a SQL-query with all your needs. Sorting, Deleting, Paging, etc.
''					It is known as Datagrid from the "big-brother" asp.net. But even better ;) This class is
''					very complex and so it should handle every problem. There are also a lot of workarounds
''					for some common problems. e.g. You can change properties of the table-object on runtime
''					by using the displayFunction from a field. Dont forget to dim the object then.
'' @POSTFIX:		table
'' @VERSION:		1.3

'**************************************************************************************************************

class Drawtable

	'private members
	private absolutePage 			'needed for paging
	private showAllRecs				'needed for paging
	private sqlCommands				'dictonary-object for sql commands
	private filterToStore			
	private fullsearchtext			
	private	bigForTime				'if TimeDebuging is ON, this variable is needed
	private allFilters				'Dictionary Object to store the Filters
	private triggerValue			'If you need an auto-value for something
	private currentRecordID			'ID of the current-record
	private output					'We store the whole output in a StringBuilder Object.
	private radioButtonsAmount		'Here we store how many radiobutton-columns we have
	private headervariable			'This variable stores the header
	private tableSession			'the Table Session Object
	private loggerIdentification	'the identification for Logger-class
	private fltrField_, xlsExporter
	private p_fastDelete, p_fastDeleteMessage, excelIcon
	private p_height, row1Color, row2Color, rowColorHover, printIcon
	
	private property get fastDeleteMessage()
		fastDeleteMessage = empty
		if p_fastDeleteMessage <> empty then
			fastDeleteMessage = replace(p_fastDeleteMessage, "'", "\'") & "\n\n"
		end if
	end property
	
	'public members
	public SQLQuery					''[string] The sql-query for the whole table. Dont forget to include every field you will need later. Dont use ORDER BY. Use "commonsort"-property instead.
	public table					''[string] Tablename to which the data belongs to. important e.g. for fastdeleteing.
	public PK						''[string] name of the column which holds the primary key. e.g. a column which uniquely identifies each record
	public RS						''[Recordset] Recordset Object with all the data
	public allTableColumns			''[Dictionary] Collection where your columns are stored. it can be usefull to access it on runtime
	public commonSort				''[string] how should the data be ordered by default? syntax is like in the ORDER BY-clause
	public isEditable				''[bool] are the records clickable? default = true
	public allowAdding				''[bool] should adding a new record be allowed
	public title					''[string] Caption for the table. will be displayed at the top
	public formURL					''[string] the url you want to send the user to when clicking on a record or 
									''clicking the "add record" button. {0} = will be replaced by ID of the record. e.g. news.asp?id={0}
									''use addurlJS if you want to execute a javascript instead of calling an url
	public addurlJS					''[string] javascript which should be executed when clicking on a row or on the "add" buutton.
									''e.g. alert('{0}') (placeholder is replaced by the value of the PK)
	public isSortable				''[bool] enable/disable sorting. default = true
	public showFilterBar			''[bool] enable/disable filterbar
	public headersPerRow			''[int] every how many rows do you want to display the headers again. 0 = just one header.
	public target					''[string] The target for all links
	public recsPerPage				''[int] How many records per page do you want to display?
	public paging					''[bool] allow paging or not
	public debuging					''[bool] some debug parameters will be displayed if true
	public timeDebuging				''[bool] where do we loose the time?
	public database					''[string] for some sql commands we need to know which DB.
	public printing					''[bool] loads the print.css and shows a print link
	public tablewidth				''[string] the width of the table. default = 100%
	public hovereffect				''[string] enable hovereffect on mouseover of a row? default = True
	public fullsearch				''[bool] enables the fullsearch which lets you searching over the whole table.
	public printTxt					''[string] Your own Text or IMAGE for the prinitng button.
	public excelExport				''[bool] You want to allow Excel exports?
	public addTxt					''[string] Your own Text or IMAGE for the Add new recordset button.
	public excelTxt					''[string] Your own Text or IMAGE for the Export to Excel button.
	public allowHtml				''[bool] Allow HTML-Tags in tabledata? ATTENTION! SETTING IT "FALSE" WILL CAUSE SPEED REDUCTION!!! default: true
	public restoreFilterOnSearch	''[bool] do you want to remember the seted filter when using the fulltextsearch?
	public stringBuilderDLL			''[bool] - is the stringBuilder-DLL available? StringBuilderDLL makes the table extremly fast!
	public extendButtonBar			''[string] Name of the function you want to execute in the buttonbar. Function must return the output! e.g. You need an extra button in the button-bar
	public tableCellpadding			''[int] Cellpadding of the table. default = 3
	public tableCellspacing			''[int] Cellspacing of the table. default = 1
	public logging					''[bool] logging on/off. if on then actions will be logged (only actions which are implemented till now.)
	public showSelectAll			''[bool] turn on the selectall feature: possibility to select all radiobuttons of a column. link is displayed in header. default = false
	public autoDelete				''[bool] turns on/off the auto delete of a record if you click on the fastdelete-button. usefull if you want to handle the delete yourself.
									''if you turn it off you will get the record id by request.form("fastdeleteID") after postback.
	public isInModal				''[bool] indicates if the table is drawn in a modal-dialog. e.g. Important for excel-export because it wont work in modal
	public filterConditionPosition	''[int] after which "WHERE" statement of the sql-query should the filter be placed? default = 1
									''e.g. SELECT * FROM t WHERE a IN (SELECT * FROM t1 WHERE b = c) by default the filter will be placed
									''and executed after the first WHERE. in some cases (advanced use) you need the filter to be executed after
									''the second, third, etc. WHERE. this property lets you change that.
	public onRowCreated				''[string] name of the sub you want to execute when a row is created. it is raised before it will be drawn.
									''eventargument is an instance of ROW.
	public legend					''[TableLegend] a (optional) legend for the table.
	public cssLocation				''[string] location of the stylesheet file. by default it is taken from the config.asp. 
	public addurl					''[string] OBSOLETE! use formURL instead!
	public defaultStylesheet		''[bool] OBSOLETE! use the setting in const.asp to set the styles.
	public showFooter				''[bool] should the footer of the table be shown (records per page, paging, etc). default = true
	public showColumnHeaders		''[bool] should the headers of the columns be shown? default = true
	public showTitleBar				''[bool] should the titlebar (title, fullsearch, excelexport, etc.) be shown. default = true
	public enableResetFilter		''[bool] should the reset-filter possibility? default = true
	public autoGenerateColumns		''[bool] auto generate the columns? default = true. if using the addColumn once then its disabled.
									''the caption is called the same as the fieldname just capitalized (first letter uppered)
									''the PK is excluded.
	public pageNumbersAmount		''[int] amount of page numbers which should be displayed in the paging bar
	
	public default property get height ''[int] gets the height
		height = p_height
	end property
	
	public property let height(value) ''[int] height in pixels of the table in order to let the headers stay on top. 0 = no height
		p_height = value
		'we disable the headersperrow if we have a height, because then the headers are on top
		if value > 0 then headersPerRow = 0
	end property
	
	public property let fastDeleteMessage(value) ''[string] sets an additional message for the fastdelete confirm-box.
		p_fastDeleteMessage = value
	end property
	
	public property get getCurrentRecordID ''[string] gets the unique-id for the current record. 
		set getCurrentRecordID = currentRecordID
	end property
	
	public property let fastDelete(value) ''[bool] enables fast-deleting. tablename must be set before!
		p_fastDelete = false
		if value then
			if table <> empty then p_fastDelete = true
		end if
	end property
	
	public property get fastDelete ''[bool] gets the fast-deleting setting
		fastDelete = p_fastDelete
	end property
	
	public property let fieldlinkID(val) ''[string] OBSOLETE! use PK instead. name of the primary key-column of the table.
		PK = val
	end property
	
	public property get fieldlinkID ''[string] OBSOLETE! use PK instead
		fieldlinkID = PK
	end property
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub Class_Initialize()
		set tableSession = nothing
		database	 			= lib.init(GL_DT_DATABASE, "mssql")
		triggerValue			= 0
		radioButtonsAmount		= 0
		set sqlCommands 		= Server.createObject("Scripting.Dictionary")
		set allTableColumns		= Server.createObject("Scripting.Dictionary")
		set allFilters			= Server.createObject("Scripting.Dictionary")
		commonsort 				= empty
		isEditable 				= true
		allowAdding 			= true
		formURL					= empty
		sqlquery 				= empty
		addurl 					= empty
		addUrlJS				= empty
		title 					= empty
		isSortable				= true
		p_fastDelete			= false
		timeDebuging			= false
		table					= empty
		headersPerRow			= lib.init(GL_DT_HEADERPERROWS, 30)
		recsPerPage				= lib.init(GL_DT_RECORDSPERPAGE, 50)
		pageNumbersAmount		= lib.init(GL_DT_PAGENUMBERSAMOUNT, 10)
		paging					= false
		target					= empty
		absolutePage 			= 0
		showAllRecs				= false
		debuging 				= false
		printing				= true
		fullsearch				= true
		tablewidth				= "100%"
		hovereffect				= true
		printTxt				= TXTSTRPRINT
		excelExport				= false
		addTxt					= TXTADDNEWRECORD
		excelTxt				= TXTEXPORTEXCEL
		defaultStylesheet		= false
		allowHTML				= true
		restoreFilterOnSearch	= false
		bigForTime				= 0
		stringBuilderDLL		= true
		headervariable			= empty
		extendButtonBar			= empty
		tableCellpadding		= 3
		tableCellspacing		= 0
		logging					= true
		loggerIdentification	= "drawtableLogs"
		fltrField_				= "fltrField_"
		p_fastDeleteMessage		= empty
		showSelectAll			= false
		autoDelete				= true
		isInModal				= false
		set xlsExporter			= new ExcelExporter
		filterConditionPosition	= 1
		onRowCreated			= empty
		set legend				= nothing
		p_height				= 0
		cssLocation				= lib.init(GL_DT_CSSLOCATION, DT_CLASSLOCATION & "standard.css")
		showFooter 				= true
		showColumnHeaders		= true
		showTitleBar			= true
		enableResetFilter		= true
		autoGenerateColumns		= true
		showFilterBar			= false
		PK						= empty
		row1Color = lib.init(GL_DT_ROW_COLOR_1, "#ffffff")
		row2Color = lib.init(GL_DT_ROW_COLOR_2, "#eeeeee")
		rowColorHover = lib.init(GL_DT_ROW_COLOR_HOVER, "#FEFDE0")
		printIcon = lib.init(GL_DT_PRINTICON, consts.STDAPP("icons/print.gif"))
		excelIcon = lib.init(GL_DT_EXCELICON, consts.STDAPP("icons/file_xls.gif"))
	end sub
	
	'**********************************************************************************************************
	'* destructor 
	'**********************************************************************************************************
	private sub class_terminate()
		set xlsExporter = nothing
		set sqlCommands = nothing
		set allTableColumns = nothing
		set allFilters = nothing
		set tableSession = nothing
		set legend = nothing
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Draws the whole table.
	'******************************************************************************************************************
	public function draw()
		if PK = empty then lib.error("ICE-8999: PK is needed (name of the column with the primary key).")
		
		set tableSession = new TableSessionObject
		tableSession.InitializeSessionObject()
		absolutePage = tableSession.getAbsolutePage()
		
		init_StringBuilder()
		deleteRecord()
		tableHeader()
		genTable()
		
		'if we have stringbuilder then we show the output
		if stringBuilderDLL then
			str.write(output.toString)
			set output = nothing
			set headervariable = nothing
		end if
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Adds a filter to a special column.
	'' @DESCRIPTION:	It generates a dropdown field or a common text field for the column.
	''					If you need a filter for readio-button-fields just name the filter the same so the class
	''					will know that this fields go together.
	'' @PARAM:			- colObj: filter-object from class_columnFilter
	'******************************************************************************************************************
	public function addNewFilter(colObj)
		allFilters.add trigger(), colObj
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	checks if the form has been submitted. use it if you want work with radiobuttons, etc.
	'' @RETURN:			[bool] true if yes
	'**************************************************************************************************************
	public function dataSubmitted()
		dataSubmitted = (request.form("save") <> "")
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	returns an array with all the IDs of the records which have been changed
	'' @DESCRIPTION:	changed refers to clicking on a radiobutton in a row
	''					arrayfield also includes the value which has been set. e.g. ID 10 to value 2 => 10:2
	'' @RETURN:			[array] array with IDs (PK's) and value separated by ':' which have been changed.
	'**************************************************************************************************************
	public function getChangedRows()
		getChangedRows = split(trim(replace(request.form("strMatrix"), ",", " ")), " ")
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	OBSOLETE! Add filter method. This is an old one. You should use addNewFilter
	'' @DESCRIPTION:	It creates the filter object itself but after you called the method. 
	'**************************************************************************************************************
	public sub addFilter(fieldName, showDropdownSQL, primaryKey, displayField, defaultSelect, fieldToMatch)
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
	'' @SDESCRIPTION:	OBSOLETE! use fastDelete-property instead
	'' @DESCRIPTION:	should fast delete be enabled?
	'' @PARAM:			val [bool]: enable/disable
	'******************************************************************************************************************
	public sub set_allowFastDelete(val)
		fastDelete = val
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Adds a column to your table.
	'' @DESCRIPTION:	autogeneration of the columns will be 
	'' @PARAM:			colObj [column]: needs a column-instance
	'******************************************************************************************************************
	public sub addColumn(colObj)
		'if a column is added then we disable the autogenerate columns
		autoGenerateColumns = false
		allTableColumns.add trigger(), colObj
		'we check if its a radiobutton and increase the radiobutton amount.
		'so later we can check if radiobuttons are available in this table
		if typename(colObj) = "columnRadioButton" then
			radioButtonsAmount = radioButtonsAmount + 1
		end if
		if colObj.isFiltered then
			set aFilter = new ColumnFilter
			aFilter.fieldname = colObj.fieldname
			addNewFilter(aFilter)
		end if
	end sub
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Check if the Session Object for the filters exists.
	'' @DESCRIPTION:	If you use a dropdown filter and your defaultSelect property isn't empty then use this
	''					function to obey the bug occuring when you select e.g. - Filter - in the dropdown (RPIS)
	'' @RETURN:			[bool]: true if object exists
	'**************************************************************************************************************
	public function existsTableSession()
		if tableSession.checkSessionObject() then
			if tableSession.getSessionObject("pageurl") = "" then
				existsTableSession = false
			else
				existsTableSession = true
			end if
		else
			existsTableSession = false
		end if
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	OBSOLETE! Add field method. This is an old one. You should use addColumn-method
	'' @DESCRIPTION:	It creates the field object itself but after you called the method. 
	'**************************************************************************************************************
	public function addfield(fieldname, displaystring, displayfunction, tdAttributes)
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
	
	'**************************************************************************************************************
	' getSqlCommands 
	'**************************************************************************************************************
	private function getSqlCommands(databaseName)
		select case databaseName
			case "oracle", "mssql"
				sqlCommands.add "ucase", "upper"
			case "access"
				sqlCommands.add "ucase", "ucase"
		end select
	end function
	
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
	' deleteRecord 
	'**************************************************************************************************************
	private function deleteRecord()
		if fastDelete and autoDelete then
			delid = empty
			'we get the delID from the hiddenfield
			delid = request.form("fastdeleteID")
			if not delid = empty then 'if there is a delid then we delete
				if logging then
					call lib.logAndForget(loggerIdentification, "ID" & delid & " deleted from """ & title & """ (" & table & ")")
				end if
				delsql = "DELETE FROM " & table & " WHERE " & me.PK & " = " & delid
				lib.getrecordset(delsql)
			end if
		end if
	end function
	
	'**************************************************************************************************************
	' init_StringBuilder
	'**************************************************************************************************************
	private function init_StringBuilder()
		if stringBuilderDLL then
			Set output = Server.CreateObject("StringBuilderVB.StringBuilder")
			output.Init 40000, 7500
			Set headervariable = Server.CreateObject("StringBuilderVB.StringBuilder")
			headervariable.Init 5000,1000
		end if
	end function
	
	'**************************************************************************************************************
	' tableHeader
	'**************************************************************************************************************
	private sub tableHeader()
		%><!--#include file="js.asp"--><%
		'We load the default css
		addToOutput ("<link rel=""stylesheet"" type=""text/css"" href=""" & DT_CLASSLOCATION & "drawtable.css"">")
		addToOutput ("<link rel=""stylesheet"" type=""text/css"" href=""" & cssLocation & """>")
		
		printButton = empty
		if printing then
			printTxtNew = printTxt
			if printIcon <> "" then printTxtNew = "<img src=""" & printIcon & """ align=absmiddle border=0>&nbsp;" & printTxtNew
			printButton = "&nbsp;<button onclick=""window.print();return false;"" type=button class=button>" & printTxtNew & "</button>"
		end if
		
		addToOutput ("<form name=myVeryHiddenForm method=post action=""" & request.serverVariables("URL") & "?" & lib.getAllFromQueryStringBut("fastDeleteID") & """ style=""display:inline;"">")
		addToOutput ("<table border=0 cellpadding=0 cellspacing=0 width=100% >")
		addToOutput ("<tr><td>")
		addToOutput ("<table border=0 cellpadding=2 cellspacing=0 width=""" & tablewidth & """>")
		
		if showTitleBar then
			addToOutput ("<tr class=headlineBG>")
			if not title = empty then addToOutput ("<td class=headline>" & title & "</td>")
			addToOutput ("<td align=right>")
			addToOutput ("<span class=noprint>")
			addToOutput ("<table border=0 cellpadding=0 cellspacing=0><tr>")
			
			'our fulltextsearch field
			if fullsearch then
				addToOutput ("<td>")
				displayFullSearchField()
				addToOutput ("</td>")
			end if
			
			'excelexport button.
			displayexcelexport()
			
			'print button
			if not printButton = empty then addToOutput ("<td>" & printButton & "</td>")
			
			'add button
			myLink = getAddLink()
			if myLink <> "" then addToOutput ("<td>&nbsp;" & myLink & "</td>")
			
			'additional buttons. Function/sub will be executed
			if not extendButtonBar = empty then
				execute("additionalFunction = " & extendButtonBar)
				addToOutput ("<td>&nbsp;" & additionalFunction & "</td>")
			end if
			
			addToOutput ("</tr>")
			addToOutput ("</table>")
			addToOutput ("</span>")
			addToOutput ("</td>")
			addToOutput ("</tr>")
		end if
		addToOutput ("</table>")
		addToOutput ("</td>")
		addToOutput ("</tr>")
		addToOutput ("</table>")
	end sub
	
	'******************************************************************************************************************
	'* printHeader
	'* here we create the headlines for every table. if commonsort = empty then sorting is disabled!
	'******************************************************************************************************************
	private sub printHeader(myFields, commonSort)
		if trim(title) <> "" then xlsExporter.addOutput("<table><tr><td><b>" & title & "</b></td></tr></table>")
		xlsExporter.addOutput("<table border=1>")
		
		if not showColumnHeaders then exit sub
		
		myAhref = empty 'it stays empty if we dont want orderlinks
		addToHeaderVariable ("<tr align=left>")
		
		mySort = request.form("sortValue")
		if mySort = "" then mySort = commonSort end if
		
		xlsExporter.addOutput("<tr>")
		
		for each col in allTableColumns.items
			
			selectAllString = empty
			if showSelectAll then
				if col.colType = "radiobutton" then
					if not col.disabled then
						selectAllString = "<br><span title=""" & TOOLTIP_TXT_SELECTALL & """ class=""hand selectAll"" onclick=""selectAllRadioButtons('" & col.value & "')"">" & TXTSELECTALL & "</span>"
					end if
				end if
			end if
			
			showCurrentSort = empty
			'if table is sorted by this col we set sort DESC to enable sorting other way!
			if instr(mySort,"DESC") then
				tmpArr = split(mySort,"XXYYZZ")
				if tmpArr(0) = col.fieldName then
					showCurrentSort = "&nbsp;&nbsp;<img align=absmiddle src=" & consts.STDAPP("icons/sortdesc.gif") & " border=0>"
				end if
			end if
			if mySort = col.fieldName then
				showCurrentSort = "&nbsp;&nbsp;<img align=absmiddle src=" & consts.STDAPP("icons/sortasc.gif") & " border=0>"
				sortby = col.fieldName & "XXYYZZ" & "DESC" 'we cannot give a Space with querystring so we save it as + and replace it later
			else
				sortby = col.fieldName
			end if
			
			'if we allow sorting then we make the link
			if isSortable then
				myAhref = "<a title=""" & TOOLTIP_TXT_SORTBY & " " & col.displayString & """ class='head' href=""javascript:document.myVeryHiddenForm.sortValue.value='" & sortby & "';submitTableData();"">" & col.displayString & "</a>"
			else
				myAhref = col.displayString
			end if
			
			addToHeaderVariable("<td class=head " & col.headerTdAttributes & " nowrap><nobr>" & myAhref & showCurrentSort & "</nobr>" & selectAllString & "</td>")
			xlsExporter.addOutput("<td><b>" & col.displayString & "</b></td>")
		next
		
		xlsExporter.addOutput("</tr>")
		if fastDelete or showFilterBar then
			addToHeaderVariable ("<td class='head'>&nbsp;</td>")
		end if
		addToHeaderVariable ("</tr>")
	end sub
	
	'*************************************************************************************************************
	' filterBar 
	'*************************************************************************************************************
	private sub filterBar(myFieldArr)
		
		if not allFilters.count = 0 and showFilterBar then 'we dont show anything if there are no filters defined
			addToOutput ("<tr class=filterBG>")
			
			filterToStore = empty 'we need this to remember the filterfield if there will be an update of rbform
			for each col in allTableColumns.items
				matched = "no"
				sql = ""
				autoSplit = true
				for each fil in allFilters.items
					if uCase(col.fieldName) = uCase(fil.fieldName) then
						matched = fil.fieldName
						sql = fil.showDropdownSQL
						myPK = fil.primaryKey
						autoSplit = fil.dropdownAutosplit
						showName = fil.displayField
						attributes = fil.tdAttributes
						inputAttributes = fil.inputAttributes
						myVal = tableSession.getFilterValue(fltrField_ & matched)
						'if the form was submitted we select the value of the dropdown
						if fil.description = empty then
							commonValue = TXTCOMMONFILTER
						else
							commonValue = fil.description
						end if
						if myVal = empty and request.form.count = 0 then myVal = fil.defaultSelect
					end if
				next
				
				'check if the field is a Radiobutton field, so we need a colspan
				if radioButtonsAmount = 0 then
					colspan = empty
				else
					if not lastFilter = matched then
						if col.colType = "radiobutton" then
							colspan = " colspan=" & radioButtonsAmount
						else
							colspan = empty
						end if
					else
						colspan = empty
					end if
				end if
				
				if not matched = "no" and not lastFilter = matched then
					addToOutput ("<td" & colspan & " " & attributes & "><span class=noprint>")
					if sql = "" then 'we make a common text filter field
						tooltipText = replace(TOOLTIP_TXT_FILTER_TEXTFIELD,"{fieldname}", col.displayString)
						addToOutput ("<label title=""" & tooltipText & """><input type=text value=""" & str.HTMLEncode(myVal) & """ name=" & fltrField_ & matched & " " & onChangeSubmit() & " " & inputAttributes & "></label>")
					else 'we must create a dropdown
						'now we build the dropdown
						set dd = new createDropdown
						with dd
							.addDisplayField (showName)
							.pk = myPK
							.sqlquery = sql
							.enableAutosplit = autoSplit
							.name = fltrField_ & matched
							.idToMatch = myVal
							.commonTxt = commonValue
							.commonTxtVal = "XXXYYYZZZXXX"
							.onAttribute = onChangeSubmit()
							addToOutput(.getAsString)
						end with
						set dd = nothing
					end if
					
					'we store the field
					filterToStore = filterToStore & hiddenFieldString(fltrField_ & matched, myVal) & vbcrlf
					addToOutput("</span></td>")
					lastFilter = matched
				else
					if not lastFilter = matched then
						addToOutput("<td" & colspan & " " & attributes & "></td>")
					end if
				end if
			next
			addToOutput("<td align=center nowrap valign=top>")
			if enableResetFilter then
				addToOutput("<input type=submit name=go class=button_common style='width:0px;height:0px;'><span class=noprint><button type=button title=""" & TOOLTIP_TXT_FILTERRESET & """ style='height:12px;cursor:pointer;' onclick=""restoreMyFilter();myVeryHiddenForm.submit();"" name=restore_my_filter class=button_common>^</button></span>")
			end if
			addToOutput("</td>")
			addToOutput("</tr>")
		end if
	end sub
	
	'******************************************************************************************************************
	'* getAddLink
	'******************************************************************************************************************
	private function getAddLink()
		getAddLink = ""
		'only return when adding possible
		if not ((addUrl <> "" or formURL <> "" or addURLJS <> "") and allowAdding) then exit function
		
		if target <> empty then
			myTarget = "parent." & target & "."
		else
			myTarget = "window."
		end if
		
		if formURL <> "" then
			myURL = str.format(formURL, array(""))
		else
			'addurl is obsolete
			myURL = addUrl
		end if
		
		if addUrlJS <> "" then
			getAddlink = "<button type=button class=button onclick=""" & str.format(addUrlJS, array("")) & """>" & addTxt & "</button>"
		else
			getAddlink = "<button type=button class=button onclick=""" & myTarget & "location.href='" & myURL & "'"">" & addTxt & "</button>"
		end if
	end function
	
	'******************************************************************************************************************
	'* getValidOperatorFrom 
	'* checks if there is a valid operator in a given string.
	'* returns an array. first field is the found operator. 2nd field is the string without the operator
	'* if array-length is -1 then nothing was found.
	'******************************************************************************************************************
	private function getValidOperatorFrom(strToCheck)
		returnValue = array()
		'we go through our array with allowed filter operators
		for p = 0 to ubound(ALLOWED_FILTER_OPERATORS)
			if left(strToCheck, len(ALLOWED_FILTER_OPERATORS(p))) = ALLOWED_FILTER_OPERATORS(p) then
				redim preserve returnValue(1)
				returnValue(0) = ALLOWED_FILTER_OPERATORS(p)
				returnValue(1) = trim(right(strToCheck, (len(strToCheck) - len(returnValue(0)))))
				exit for
			end if
		next
		getValidOperatorFrom = returnValue
	end function
	
	'******************************************************************************************************************
	'* genTable 
	'******************************************************************************************************************
	private sub genTable()
		'first we get our sqlcommands. depend on databse.
		getSqlCommands(database)
		
		sql = sqlQuery
		
		if showFilterBar then
			'if there is a where in our sql statement then its okay. else we add a where 1=1 ourself
			if instr(ucase(sql), "WHERE") = 0 then sql = sql & " WHERE 1=1"
		end if
		
		'get the sortValue to know what column to sort
		mysort = tableSession.getSortValue(allFilters, isSortable)
		
		if len(mySort) > 1 then commonSort = mySort
		
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
				if cstr(sort1) = cstr(col.fieldname) AND col.isNumber then ucased = empty
			next
			sql = sql & " ORDER BY " & ucased & "(" & str.SQLSafe(sort)
		else 'we have a commonsort 
			if not commonSort = "" then
				'if we have more fields to sort
				select case database
					case "oracle"
						if instr(commonSort, ",") then
							mySorts = split(commonSort, ",")
							for o = 0 to ubound(mySorts)
								if instr(ucase(mySorts(o))," DESC") = 0 then
									sOrd = sOrd & sqlCommands.Item("ucase") & "(" & mySorts(o) & "),"
								else
									sOrd = sOrd & sqlCommands.Item("ucase") & "(" & left(mySorts(o),(len(mySorts(o))-5)) & ") DESC,"
								end if
							next
							sql = sql & " ORDER BY " & str.SQLSafe(left(sOrd,(len(sOrd)-1)))
						else
							if instr(ucase(commonSort)," DESC") then
								sql = sql & " ORDER BY " & sqlCommands.Item("ucase") & "(" & str.SQLSafe(left(commonSort,(len(commonSort)-5))) & ") DESC"
							else
								sql = sql & " ORDER BY " & sqlCommands.Item("ucase") & "(" & str.SQLSafe(commonSort) & ")"
							end if
						end if
					case "mssql"
						sql = sql & " ORDER BY " & str.SQLSafe(commonSort)
				end select
			end if
		end if
		
		'use the FILTER
		if showFilterBar then
			'we split the sql to add the filter condition. we make sql ucase so we be sure to find the where
			splittedSQL = split(UCase(sql), "WHERE")
			atLeastOneCriteria = false
			
			'we check if tere are some filter selected
			if request.form.count > 0 then
				set myFilters = request.form
			else
				set myFilters = tableSession.getAllFilters
			end if
			
			for each fild in myFilters
				myFieldName = empty
				if instr(fild,fltrField_) and not tableSession.getFilterValue(fild) = empty and not tableSession.getFilterValue(fild) = "XXXYYYZZZXXX" then
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
							
							searchKeyword = ltrim(tableSession.getFilterValue(fild)) 'the value we search for
							
							'If we have a custom comma-style we need to replace it
							if not fil.commaStyle = empty then
								tmp = replace(searchKeyword, fil.commaStyle, ".")
								searchKeyword = tmp
							end if
							
							'we check if there was an operator found. if yes then we execute the sql with the operator
							operatorArray = getValidOperatorFrom(searchKeyword)
							if uBound(operatorArray) > -1 then
								splittedSQL(filterConditionPosition) = " (" & matchField & ") " & operatorArray(0) & " ('" & str.SQLSafe(operatorArray(1)) & "') AND " & splittedSQL(filterConditionPosition)
							else 'we execute the sql with pre-defined statements. = and LIKE
								'we have no sql so its a common text field. so we need a like
								if fil.showDropdownSQL = empty then
									splittedSQL(filterConditionPosition) = " " & sqlCommands.Item("ucase") & "(" & matchField & ") LIKE " & sqlCommands.Item("ucase") & "('%" & str.SQLSafe(searchKeyword) & "%') AND " & splittedSQL(filterConditionPosition)
								else
									splittedSQL(filterConditionPosition) = " (" & matchField & " = '" & str.SQLSafe(searchKeyword) & "') AND " & splittedSQL(filterConditionPosition)
								end if
							end if
						end if
					next
				end if
			next
			
			set fil = nothing
			
			'we check if there are default values for a filter
			if not atLeastOneCriteria then
				for each fil in allFilters.items
					if not cstr(fil.defaultSelect) = empty and request.form.count = 0 and tableSession.getFilterValue(fltrField_ & fil.fieldName) = empty then
						atLeastOneCriteria = true
						'We check if there is a special fieldname to compare with. if not we take the one which we displayed
						if not fil.fieldToMatch = empty then
							matchField = fil.fieldToMatch
						else
							matchField = fil.fieldName
						end if
						'we have no sql so its a common text field. so we need a like
						if fil.showDropdownSQL = empty then
							operatorArray = getValidOperatorFrom(fil.defaultSelect)
							if uBound(operatorArray) > -1 then
								splittedSQL(filterConditionPosition) = " (" & matchField & ") " & operatorArray(0) & " ('" & str.SQLSafe(operatorArray(1)) & "') AND " & splittedSQL(filterConditionPosition)
							else
								splittedSQL(filterConditionPosition) = " UPPER(" & matchField & ") LIKE UPPER('%" & str.SQLSafe(fil.defaultSelect) & "%') AND " & splittedSQL(filterConditionPosition)
							end if
						else
							splittedSQL(filterConditionPosition) = " " & matchField & " = '" & str.SQLSafe(fil.defaultSelect) & "' AND " & splittedSQL(filterConditionPosition)
						end if
					end if
				next
			end if
			
			'this is for the addressbook-bug; in the addressbook we have one dropdown-filter(country), at the employees (which has the same
			'tablesession object, because it has the same url) we have two dropdowns (country and department)
			'if you activate the country filter in the addressbook, the department dropdown is not available in the session object ->
			'if you change the view to the employees, the department dropdown will receive the default value
			if existsTableSession and atLeastOneCriteria then
				for each fil in allFilters.items
					if not cstr(fil.defaultSelect) = empty and tableSession.getFilterValue(fltrField_ & fil.fieldName) = empty then
					if not fil.fieldToMatch = empty then
							matchField = fil.fieldToMatch
						else
							matchField = fil.fieldName
						end if
						if not fil.showDropdownSQL = empty then
							splittedSQL(filterConditionPosition) = " " & matchField & " = '" & str.SQLSafe(fil.defaultSelect) & "' AND " & splittedSQL(filterConditionPosition)
						end if
					end if
				next
			end if
			
			if atLeastOneCriteria then
				sql = empty
				for ij = 0 to uBound(splittedSQL)
					if ij > 0 then sql = sql & " WHERE "
					sql = sql & " " & splittedSQL(ij)
				next
			end if
		end if
		
		'load the data
		if debuging then addToOutput ("<strong>SQL-Query: </strong>" & sql & "<br>")
		startRecordSet = timer()
		set me.RS = lib.getUnlockedRecordset(sql)
		EndRecordSet = timer()
		if timedebuging then addToOutput ("unlocked recordset received in : " & EndRecordSet - startRecordSet & "<br>")
		if autoGenerateColumns then
			for each RSField in RS.fields
				'we exclude the primary key column
				if uCase(RSField.name) <> uCase(me.PK) then
					set column = new ColumnCommon
					with column
						.fieldName = RSField.name
						.displayString = str.capitalize(.fieldName)
						.isNumber = (RSField.type = 5 or RSField.type = 6 or RSField.type = 3)
					end with
					addColumn(column)
					set column = nothing
				end if
			next
		end if
		
		'check the paging
		if paging then
			fullsearchroutine(RS)
			
			if not RS.eof then
				if absolutePage = empty then absolutePage = 1
				
				RS.PageSize = recsPerPage
				RS.CacheSize = recsPerPage
				
				if absolutePage = 0 then
					showAllRecs = true
					RS.absolutePage = 1
				end if
				
				If absolutePage = "" or not isNumeric(absolutePage) or cint(absolutePage) > cint(RS.PageCount) then
					showAllRecs = false
					absolutePage = 1
					RS.absolutePage = absolutePage
				else
					if not absolutePage = 0 then RS.absolutePage = absolutePage
				end if
			end if
		else
			fullsearchroutine(RS)
		end if
		
		set sqlCommands = nothing
		counter = 0
		sumCount = 0
		addToOutput("<div id=drawtable class=tableContainer " & lib.iif(height > 0 and lib.browser = "IE", "style=""height:" & height & "px""", empty) & ">")
		addToOutput("<table width=" & tablewidth & " style=""border-collapse:collapse"" cellspacing=" & tableCellspacing & " cellpadding=" & tableCellpadding & " border=0 class=""tableClass"">")
		if lib.browser = "IE" then addToOutput("<thead>")
		'show header first time. if nothing found also headers will be shown
		printHeader myFields,commonSort
		if stringBuilderDll then
			addToOutput(headerVariable.toString)
		else
			addToOutput(headervariable)
		end if
		
		'we write the current sort as value.
		addToOutput(hiddenFieldString("sortValue", mysort))
		'we save the current page number to
		addToOutput(hiddenFieldString("actualPageNumber", absolutePage))
		
		'show us the cool filter bar
		filterBar(myFields)
		
		if lib.browser = "IE" then addToOutput("</thead>")
		
		'if we want to delete a record with fastDelete then we write it in here
		if fastDelete then addToOutput(hiddenFieldString("fastDeleteID", empty))
		
		'firefox needs to be treated otherway in order to enable scrolling the headers
		if lib.browser = "FF" then
			addToOutput("<tbody " & lib.iif(height > 0, "style=""height:" & height & "px""", empty) & ">")
		else
			addToOutput("<tbody>")
		end if
		
		'we loop through the records 
		if not RS.eof then
			if timedebuging then
				bigWhileStart = timer()
			end if
			
			set currentRecordID = RS(me.PK) 'our primary-key in an object. should be faster
			
			while not RS.eof
				if (paging and counter < recsPerPage) or not paging or showAllRecs then
					
					xlsExporter.addOutput("<tr>")
					
					'SHOW US THE HEADERS. EVERY X LINE.
					if headersPerRow <> 0 then
						if not counter = 0 and counter mod headersPerRow = 0 then
							if stringBuilderDll then
								addToOutput(headerVariable.toString)
							else
								addToOutput(headervariable)
							end if
						end if
					end if
					
					if counter = 0 then
						addToOutput("</form>")
						addToOutput("<form name=rbfrm method=post action=""" & request.servervariables("URL") & "?" & lib.getallFromQuerystringBut("fastDeleteID") & """>")
						addToOutput(hiddenFieldString("strMatrix", empty))
						addToOutput(hiddenFieldString("save", "true"))
						
						'her we have a hidden input text to allow the programmer to store something
						addToOutput(hiddenFieldString("valueToStore", empty))
						
						'here we save some values if form will be updated. to remind the settings
						addToOutput(hiddenFieldString("sortValue", mysort))
						addToOutput(hiddenFieldString("actualPageNumber", absolutePage))
						
						'here we write our stored filter fields
						addToOutput(filterToStore)
					end if
					
					set aRow = new TableRow
					with aRow
						.index = counter
						.hoverEffect = hovereffect
						.recordID = currentRecordID
						.ID = "tr_" & .recordID
						.BGColor = lib.iif(.index mod 2 = 0, row1Color, row2Color)
					end with
					if onRowCreated <> empty then execute(onRowCreated & "(aRow)")
					if timedebuging then bigForStart = timer()
					
					TRString = empty
					
					if aRow.BGColor <> empty then TRString = " bgColor=" & aRow.BGColor
					if aRow.CSSClass <> empty then 
						TRString = TRString & " class=""dataRow " & aRow.CSSClass & """"
					else
						TRString = TRString & " class=""dataRow"""
					end if
					if aRow.attributes <> empty then TRString = TRString & " " & aRow.attributes
					if aRow.hoverEffect then TRString = TRString & " onmouseover=rowHoverIn(this) onmouseout=rowHoverOut(this)"
					addToOutput("<tr id=" & aRow.ID & TRString & ">")
					
					'go through all columns
					showTableData(aRow)
					
					'fast delete button
					if fastDelete then
						addToOutput("<td class=lastColumn align=right width=15><span class=noprint>")
						if not aRow.disabled then
							addToOutput("<button type=button title=""" & TOOLTIP_TXT_FASTDELETE & """ name=""XXXYYYZZZ"" onclick=""yesNoFastDelete(" & currentRecordID & ", '" & fastDeleteMessage & TXTFASTDELETEQUESTION & "')"" class=button_common>")
							addToOutput(TXTFASTDELETEBUTTON & "</button>")
						end if
						addToOutput("</span></td>")
					elseif showFilterBar then
						addToOutput("<td width=15 class=lastColumn></td>")
					end if
					
					addToOutput("</tr>")
					xlsExporter.addOutput("</tr>")
					
					set aRow = nothing
					
					if timedebuging then
						bigForEnd = timer()
						bigForTime = bigForTime + (bigForEnd - bigForStart)
						bigForEnd = 0
						bigForStart = 0
					end if
					
					counter = counter + 1
				end if
				sumCount = sumCount + 1
				RS.movenext()
			wend
			
			addToOutput("</tbody>")
			
			if timedebuging then
				bigWhileEnd = timer()
				addToOutput("big for needed: " & bigForTime & "<br>")
				addToOutput("big while needed: " & bigWhileEnd - bigWhileStart & "<br>")
			end if
		
		end if
		xlsExporter.addOutput("</table>")
		'addToOutput ("</form>")
		
		if paging then
			if not counter = 0 then
				
				pageCount = RS.PageCount
				'we calculate the sum of all records if paging is displayed
				if not absolutePage = 1 and not showAllRecs then
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
			sumRecs = 0
			pageCount = 0
		end if
		
		set RS = nothing
		tableFooter counter, pageCount, sumRecs
		tableLegend()
		
		createExcelForm()
		set currentRecordID = nothing
	end sub
	
	'******************************************************************************************************************
	'* hiddenFieldString 
	'******************************************************************************************************************
	private function hiddenFieldString(hiddenFieldName, hiddenFieldValue)
		if debuging then
			fieldType = "text"
			fieldDescription = hiddenFieldName & ": "
		else
			fieldType = "hidden"
			fieldDescription = empty
		end if
		
		hiddenFieldString = fieldDescription & "<input type=" & fieldType & " name=" & hiddenFieldName & " value=""" & hiddenFieldValue & """>"
	end function
	
	'******************************************************************************************************************
	'* pagingBar
	'******************************************************************************************************************
	private sub pagingBar(pageCount, counter)
		'show me the prev link
		if not absolutePage = 1 and not counter = 0 and not showAllRecs then
			addToOutput("<a title=""" & TOOLTIP_TXT_PREVPAGE & """ href=""javascript:document.myVeryHiddenForm.actualPageNumber.value='" & absolutePage - 1 & "';submitTableData();"" class=pagingPrevNext>" & TXTPAGINGPREV & "</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
		end if
		
		'show me the all link
		if not showAllRecs then
			addToOutput("<a title=""" & TOOLTIP_TXT_SHOWALL & """ href=""javascript:document.myVeryHiddenForm.actualPageNumber.value='0';submitTableData();"" class=pagingLink>" & TXTPAGINGALL & "</a>")
		else
			addToOutput("<span class=pagingCurrentPage>" & TXTPAGINGALL & "</span>")
		end if
		addToOutput("&nbsp;&nbsp;&nbsp;&nbsp;")
		
		'show me the page links
		for intPageCounter = 1 to pageCount
			'dont link actual page
			if CInt(intPageCounter) = CInt(absolutePage) then
			    addToOutput("<span class=pagingCurrentPage>" & intPageCounter & TXTPAGINGSEPERATOR & "</span>")
		    else
				if abs(intPageCounter - absolutePage) < pageNumbersAmount then
		        	addToOutput("<a href=""javascript:document.myVeryHiddenForm.actualPageNumber.value='" & intPageCounter & "';submitTableData();"" class=pagingLink>" & intPageCounter & TXTPAGINGSEPERATOR & "</a>")
				end if
			end if
		next
		
		'show me the next link
		if not cint(absolutePage) = cint(pageCount) and not counter = 0 and not showAllRecs then
			addToOutput("&nbsp;&nbsp;&nbsp;&nbsp;<a title=""" & TOOLTIP_TXT_NEXTPAGE & """ href=""javascript:document.myVeryHiddenForm.actualPageNumber.value='" & absolutePage + 1 & "';submitTableData();"" class=pagingPrevNext>" & TXTPAGINGNEXT & "</a>")
		end if
	end sub
	
	'******************************************************************************************************************
	'* tableFooter
	'* our table footer. incl. number of records and add new record link 
	'******************************************************************************************************************
	private sub tableFooter(counter, pageCount, sumRecs)
		if counter > 1 OR counter = 0 then
			records = " " & TXTRECORDSPLURAL
		else
			records = " " & TXTRECORDSSINGULAR
		end if
		
		if not sumRecs = 0 then sumRecsTXT = " / " & sumRecs else sumRecsTXT = empty end if
		if counter = 0 then
			'we make a fake record line to let the design be the same
			addToOutput("<tr>")
			for each col in allTableColumns.items
				addToOutput("<td " & col.tdAttributes & "></td>")
			next
			if fastDelete then addToOutput("<td width=15></td>")
			addToOutput("</tr>")
			'and now the real message that there is nothing
			addToOutput("<tr><td align=center colspan=30><div class=noRecsFound>")
			if fullsearch then fullsearchReset = "myVeryHiddenForm.fullsearchtext.value='';"
			addToOutput(TXTNORECSAVAILABLE)
			if enableResetFilter and (showFilterBar or fullsearch) then
				addToOutput("<br><a href=""javascript:" & fullsearchReset & "restoreMyFilter();myVeryHiddenForm.submit();"">" & TXTRESETALLFILTER & "</a>")
			end if
			addToOutput("</td></tr>")
		end if
		addToOutput("</table></div>")
		if showFooter then
			addToOutput("<table border=0 width=100% cellpadding=" & tableCellpadding & " cellspacing=0 class=tableFooterBG><tr>")
			addToOutput("<td nowrap valign=top><span class=recordsDisplayed>" & counter & sumRecsTXT & " " & records & ".</span></td>")
			
			if paging then
				addToOutput("<td width=100% align=right valign=top><span class=noprint>")
				PagingBar pageCount, counter
				addToOutput("&nbsp;&nbsp;</span></td>")
			end if
		end if
		
		addToOutput("	</tr></table>")
		addToOutput ("</form>")
	end sub
	
	'******************************************************************************************************************
	'* tableLegend
	'******************************************************************************************************************
	sub tableLegend()
		if not legend is nothing then addToOutput(legend.toString(TXTLEGEND))
	end sub
	
	'******************************************************************************************************************
	'* displayFullSearchField 
	'******************************************************************************************************************
	private sub displayFullSearchField()
		if restoreFilterOnSearch then
			mySubmit = " onclick=""myVeryHiddenForm.submit();"""
			onkeyPress = empty
		else
			mySubmit = " onclick=""restoreMyFilter();myVeryHiddenForm.submit();"""
			onkeyPress = " onkeypress=""checkChar(event);"""
		end if
		addToOutput("<table border=0 cellpadding=0 cellspacing=0 width=100% >")
		addToOutput("<tr>")
		addToOutput("	<td valign=middle nowrap><label for=fullsearchtext>" & TXTSEARCHALL & "</label></td>")
		addToOutput("	<td valign=middle width=100% >&nbsp;<input title=""" & TOOLTIP_TXT_SEARCHALL & """ id=fullsearchtext class=formField type=text name=fullsearchtext value=""" & str.HTMLEncode(tableSession.getSearchValue()) & """ " & onkeyPress & " style=""width:100px;""></td>")
		addToOutput("	<td valign=middle nowrap>&nbsp;<input type=button name=button class=button value=" & TXT_SEARCH & mySubmit & "></td>")
		addToOutput("</tr>")
		addToOutput("</table>")
		
		if not isInModal then addToOutput("<script language=JavaScript>myVeryHiddenForm.fullsearchtext.focus();</script>")
	end sub
	
	'******************************************************************************************************************
	'* fullsearchroutine 
	'******************************************************************************************************************
	private sub fullsearchroutine(RS)
		fullsearchtext = tableSession.getSearchValue()
		if not RS.eof then
			if fullsearch and fullsearchtext <> "" then
				filterstring = ""
				for each x in RS.fields
					if (x.type = 200) or (x.type = 201) or (x.type = 202) or (x.type = 203) then 'removed " or (x.type = 135)" because date-fields make troubles
						if filterstring = "" then 
							filterstring = x.name & " like '*" & str.SQLSafe(fullsearchtext) & "*'"
						else
							filterstring = filterstring & " or " & x.name & " like '*" & str.SQLSafe(fullsearchtext) & "*'"	
						end if
					end if
				next
				if debuging then str.write(filterstring)
				RS.filter = filterstring
			end if
		end if
	end sub
	
	'******************************************************************************************************************
	'* displayExcelExport 
	'******************************************************************************************************************
	private sub displayExcelExport()
		if excelexport then
			excelTxtNew = excelTxt
			if excelIcon <> "" then excelTxtNew = "<img border=0 src=" & excelIcon & " align=absmiddle>&nbsp;" & excelTxtNew
			addToOutput("<td>&nbsp;<button type=button onClick=""document.forms['xlform'].submit();"" class='button' name='exportto'>" & excelTxtNew & "</button></td>")
		end if
	end sub
	
	'******************************************************************************************************************
	'* createExcelForm 
	'******************************************************************************************************************
	private sub createExcelForm()
		if excelexport and showTitleBar then
			with xlsExporter
				xlsExporter.addOutput("</table>")
				if isInModal then excelFormTarget = " target=_blank"
				
				addToOutput("<form name=xlform style=""display:inline"" method=post action=""/gab_Library/class_excelexporter/index.asp""" & excelFormTarget & ">")
				addToOutput(.getHiddenFields())
				addToOutput("</form>")
			end with
		end if
	end sub
	
	'******************************************************************************************************************
	'* onChangeSubmit 
	'******************************************************************************************************************
	private function onChangeSubmit()
		onChangeSubmit = " onChange=""submitTableData();"""
	end function
	
	'******************************************************************************************************************
	'* showTableData 
	'******************************************************************************************************************
	private sub showTableData(rowObject)
		colsAmount = 0
		for each col in allTableColumns.items
			if colsAmount = 0 then
				colCSSClass = " class=firstColumn"
			'if there is no filterbar and no fast delete then we need to mark the last column
			elseif colsAmount = allTableColumns.count - 1 and not (fastDelete or showFilterBar) then
				colCSSClass = " class=lastColumn"
			else
				colCSSClass = empty
			end if
			
			if col.colType = "radiobutton" then
				if cint(col.value) = cint(RS(col.fieldname)) then
					checked = " checked"
					xlsExporter.addOutput("<td>X</td>")
				else
					checked = ""
					xlsExporter.addOutput("<td></td>")
				end if
				addToOutput("	<td align=center " & colCSSClass & " " & col.tdAttributes & " ><input type=radio name=rb_" & currentRecordID & " value=" & col.value & checked & " onclick=""changeBG('" & col.value & "', " & currentRecordID & ");""" & lib.iif(col.disabled or rowObject.disabled, " disabled", empty) & "></td>")
			else
				'we execute the displayfunction if needed
				if col.displayFunction <> "" then
					inputVar = empty
					if allowHTML then
						fieldContent = RS(col.fieldName)
					else
						fieldContent = str.stripTags(RS(col.fieldName))
					end if
					
					if not isNull(fieldContent) then
						inputVar = replace(fieldContent, """", """""")
						if instr(inputVar, vbCrLf) > -1 then inputVar = replace(inputVar, vbCrLf, " ")
					end if
					
					displayfunction = "field0 = " & col.displayFunction & "(""" & inputVar & """)"
					
					execute(displayfunction)
				else
					if allowHTML then
						field0 = RS(col.fieldName)
					else
						field0 = str.stripTags(RS(col.fieldName))
					end if
				end if
				
				if isEditable and not col.disableLink and not rowObject.disabled then
					if addUrlJS <> empty then
						tdRowJavascript = " onclick=""" & str.format(addUrlJS, array(currentRecordID)) & """ style='cursor:pointer;'"
					else
						if not target = empty then
							myTarget = " parent." & target & "."
						else
							myTarget = empty
						end if
						
						if formURL <> "" then
							myURL = str.format(formURL, array(currentRecordID))
						else
							if instr(addUrl, "?") then 'we check if the addurl has already some Querystring parameters
								qsPrefix = "&" 
							else
								qsPrefix = "?"
							end if
							myURL = addUrl & qsPrefix & me.PK & "=" & currentRecordID
						end if
						
						tdRowJavascript = " onclick=""" & myTarget & "location.href='" & myURL & "'"" style='cursor:pointer;'"
					end if
					
					editUrlBegin = "<a class=""navklein"">"
					editUrlEnd = "</a>"
				else
					editUrlBegin = empty
					editUrlEnd = empty
					tdRowJavascript = empty
				end if
				
				addToOutput("	<td " & colCSSClass & " " & col.tdAttributes & tdRowJavascript & ">" & editUrlBegin & field0 & editUrlEnd & "</td>")
				xlsExporter.addOutput("<td>" & field0 & "</td>")
			end if
			colsAmount = colsAmount + 1
		next
	end sub
	
	'******************************************************************************************************************
	'* addToOutput 
	'******************************************************************************************************************
	private sub addToOutput(str)
		if stringBuilderDLL then
			output.append str & vbCrLf
		else
			response.write str & vbCrLf
		end if
	end sub
	
	'******************************************************************************************************************
	'* addToHeaderVariable 
	'******************************************************************************************************************
	private function addToHeaderVariable(str)
		if stringBuilderDLL then
			headerVariable.append str & vbCrLf
		else
			headerVariable = headerVariable & str & vbCrLf
		end if
	end function
	
end class
lib.registerClass("Drawtable")
%>
