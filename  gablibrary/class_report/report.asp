<!--#include file="config.asp"-->
<!--#include virtual="/gab_LibraryConfig/_report.asp"-->
<!--#include file="class_reportParameter.asp"-->
<!--#include file="class_reportCell.asp"-->
<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003 - This file is part of GAB_LIBRARY		
'* For license refer to the license.txt in the root    						
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		Report
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		2006-10-02 20:56
'' @CDESCRIPTION:	NOT WORKING YET! UNFINISHED! Represents a control which lets create reports which consist of just numbers
''					(e.g. sales, etc.). As datasource a SQL-Query or Recordset is used. It provides
''					calculation possibilities which allow to use formulas within rows and columns as
''					e.g. excel does. This makes the underlying calculation nicer to read in the code.
''					Sorting of the table is implemented clientside with a script from Matt Kruse
''					at http://www.JavascriptToolbox.com/
'' @REQUIRES:		-
'' @POSTFIX:		rpt
'' @VERSION:		0.1

'**************************************************************************************************************
class Report

	'private members
	private rows, cols, output, values, regex, p_ID, adding, classLocation
	
	'TODO:
	'- excel-export
	'- highlight row and column on cell mouseover
	'- format types (percentage, money, class of rows and cols)
	'- precision. ,000 ,00
	'- round like in treasury
	'- calculator support?
	'- first column fixed!
	'- sort by column, row (javascript)
	
	'public members
	public datasource				''[string], [recordset] the datasource which is needed to generate the report.
									''normally grouped with the GROUPED BY clause. if string is given then it will be treated
									''as a SQL-Query which after executing the draw() method will change to a recordset.
	public title					''[string] title of the report. empty if no title. default = empty
	public groupBy					''[string] name of the field by which the cols/rows should be grouped. by default the
									''the grouping will be done over the columns. if axes are swaped then grouping will be done
									''over the rows
	public cssLocation				''[string] full virtual path to the stylesheet file. default = take from the config.asp
	public stringBuilder			''[bool] should the stringbuilder be used for rendering? default = true
	public width					''[string] width of the report. value is put into the width-attribute of the table. default = 80%
	public align					''[string] alignment of the report. default = center
	public cellpadding				''[int] padding of the cells within the report
	public cellspacing				''[int] spacing of the cells within the report
	public undefinedSign			''[string] the sign (value) which will be shown if a value is undefined. happens on division by zero
	public swapAxes					''[bool] NOT IMPLEMENTED YET! swaps the parameters and cols. so parameters are columns and cols are rows. default  = false
	public onCellCreated			''[string] name of the sub which should handle the onCellCreated-event
	public digitsAfterDecimal		''[int] how many digits should be displayed after the decimal. thats also the precision which will be
									''used for the calculations. default = 2
	
	public property get ID ''[string] gets the auto generated ID of the report
		if p_ID = "" then p_ID = "gl_report_" & lib.getUniqueID()
		ID = p_ID
	end property
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		set datasource = nothing
		p_ID = ""
		onCellCreated = ""
		set rows = server.createObject("scripting.dictionary")
		set cols = server.createObject("scripting.dictionary")
		classLocation = consts.gabLibLocation & "class_report/"
		cssLocation = lib.init(GL_RPT_CSSLOCATION, classLocation & "std.css")
		stringBuilder = true
		width = "80%"
		align = "center"
		title = empty
		cellpadding = 2
		cellspacing = 0
		values = array()
		undefinedSign = lib.init(GL_RPT_UNDEFINEDSIGN, "undef")
		swapAxes = false
		'we need the regex for the formulas
		set regex = new Regexp
		regex.global = true
		regex.ignoreCase = true
		'matches look like this {...} or this [...]
		regex.pattern = "\{[^\}]+\}|\[[^\]]+\]"
		adding = false
		digitsAfterDecimal = 2
	end sub
	
	'**********************************************************************************************************
	'* destructor 
	'**********************************************************************************************************
	public sub class_terminate()
		set rows = nothing
		set cols = nothing
		if stringBuilder then set output = nothing
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	draws the report
	'**********************************************************************************************************
	public sub draw()
		init()
		write("<div class=reportContainer id=" & ID & ">")
		write("<table class=""GLReport"" cellpadding=" & cellpadding & " cellspacing=" & cellspacing & " align=""" & align & """ width=""" & width & """>")
		drawHeader()
		write("<tbody>")
		colsArray = cols.items
		rowsArray = rows.items
		
		if swapAxes then datasource.movefirst()
		for i = 0 to uBound(rowsArray)
			set r = rowsArray(i)
			write("<tr class=""" & r.border & """>")
			write("<td class=""rowCaption inRowFirst"" nowrap>" & r.caption & "</td>")
			
			if not swapAxes then datasource.movefirst()
			for j = 0 to ubound(colsArray)
				set c = colsArray(j)
				
				result = calculateValue(r, c)
				
				set cell = new ReportCell
				with cell
					set .report = me
					set .row = r
					set .col = c
					set .param = result(1)
					.value = result(0)
				end with
				
				raiseOnCellCreated(cell)
				'store the value in the values matrix. the cell ensures the number of the value
				values(r.index, c.index) = cell.value
				cell.draw ubound(colsArray) = j, ubound(rowsArray) = i
				if not swapAxes and not datasource.eof then datasource.movenext()
			next
			
			if swapAxes and not datasource.eof then datasource.movenext()
			write("</tr>")
		next
		write("</tbody>")
		write("</table>")
		write("</div>")
		
		'render the output
		str.writeln(output.toString())
	end sub
	
	'**********************************************************************************************************
	'* raiseOnCellCreated 
	'**********************************************************************************************************
	private sub raiseOnCellCreated(cell)
		if onCellCreated <> "" then execute(onCellCreated & "(cell)")
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	adds a row
	'' @PARAM:			aRow [ReportParameter]
	'**********************************************************************************************************
	public sub addRow(aRow)
		if not adding and swapAxes then
			adding = true
			addCol(aRow)
			exit sub
		end if
		addParam aRow, rows
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	adds a column
	'' @PARAM:			aCol [ReportParameter]
	'**********************************************************************************************************
	public sub addCol(aCol)
		if not adding and swapAxes then
			adding = true
			addRow(aCol)
			exit sub
		end if
		addParam aCol, cols
	end sub
	
	'**********************************************************************************************************
	'* addParam 
	'**********************************************************************************************************
	private sub addParam(aParam, byRef collection)
		if collection.exists(aParam.name) or aParam.name = "" then
			lib.error("Parameter '" & aParam.caption & "' must have a unique name within the collection (rows or cols).")
		end if
		aParam.index = collection.count
		collection.add uCase(aParam.name), aParam
		adding = false
	end sub
	
	'**********************************************************************************************************
	'* calculateValue 
	'* returns an array with 2 fields. 1st = calculated value, 2nd = param which was used to calculate it
	'**********************************************************************************************************
	private function calculateValue(row, col)
		dim returnVal(1)
		rowHasPriority = true
		
		if col.hasFormula() or row.hasFormula() then
			if row.hasFormula() and col.hasFormula() then
				if col.priority then rowHasPriority = false
			elseif col.hasFormula() then
				rowHasPriority = false
			end if
		else
			if col.value <> "" then
				returnVal(0) = datasource(col.value)
				set returnVal(1) = col
				calculateValue = returnVal
				exit function
			elseif row.value <> "" then
				returnVal(0) = datasource(row.value)
				set returnVal(1) = row
				calculateValue = returnVal
				exit function
			else
				lib.error("No value defined for either row or column.")
			end if
		end if
		
		if rowHasPriority then
			set returnVal(1) = row
			'set priorityField = row
			set priorityCollection = rows
			set field = col
		else
			set returnVal(1) = col
			'set priorityField = col
			set priorityCollection = cols
			set field = row
		end if
		
		formulaParsed = uCase(returnVal(1).value)
		set vars = regex.execute(formulaParsed)
		
		'this array holds the values which the parsedforumla will use. the reason for that
		'is because we need to execute() the formula and when we bang in the number directly then
		'they will be written different to how they are stored. e.g. 5.333 will be written 5,333 => this results
		'in and ASP error.
		varValues = array()
		redim varValues(vars.count - 1)
		i = 0
		'iterate through all vars in the formula. if one is undefined then exit the function,
		'because undefined will stay undefined no matter what you do with it.
		for each var in vars
			varName = str.trimStart(str.trimEnd(var, 1), 1)
			'its a variable from parameters
			if str.startsWith(var, "{") then
				if not priorityCollection.exists(varName) then lib.error("Cannot find parameter called '" & varName & "'")
				varValues(i) = getValue(rowHasPriority, priorityCollection(varName).index, field.index)
			'its a function to execute
			elseif str.startsWith(var, "[") then
				select case varName
					case "SUM", "AVG"
						for j = 0 to priorityCollection(uCase(returnVal(1).name)).index
							currentValue = getValue(rowHasPriority, j, field.index)
							if currentValue = undefinedSign then
								varValues(i) = undefinedSign
								exit for
							end if
							varValues(i) = varValues(i) + currentValue
						next
						if varName = "AVG" and varValues(i) <> undefinedSign then
							if j > 1 then
								varValues(i) = varValues(i) / (j - 1)
							elseif j = 1 then
								varValues(i) = varValues(i)
							else
								varValues(i) = undefinedSign
							end if
						end if
					case "MIN", "MAX"
						for j = 0 to priorityCollection(uCase(returnVal(1).name)).index - 1
							currentValue = getValue(rowHasPriority, j, field.index)
							if currentValue = undefinedSign then
								varValues(i) = undefinedSign
								exit for
							end if
							if j = 0 then varValues(i) = currentValue
							if varName = "MIN" then
								if currentValue < varValues(i) then varValues(i) = currentValue
							else
								if currentValue > varValues(i) then varValues(i) = currentValue
							end if
						next
					case else
						lib.error("Unrecognized function '" & varName & "'")
				end select
			end if
			'we return undefined if something was undefined
			if varValues(i) = undefinedSign then
				returnVal(0) = undefinedSign
				calculateValue = returnVal
				exit function
			end if
			'we're getting sure and make the value a number
			varValues(i) = cDbl(varValues(i))
			formulaParsed = replace(formulaParsed, var, "varValues(" & i & ")")
			i = i + 1
		next
		
		'executes the parsed formula without the first char because this is an =.
		'we catch the errors on executing the function.
		'if its a division by zero then we set the value to undefined
		on error resume next
		execute("returnVal(0) " & formulaParsed)
		if err <> 0 then
			if err.number = 11 then
				returnVal(0) = undefinedSign
			else
				on error goto 0
				lib.error(array("Cannot execute function '" & returnVal(1).value & "'.", "Parsed to '" & formulaParsed & "'"))
			end if
		end if
		on error goto 0
		calculateValue = returnVal
	end function
	
	'**********************************************************************************************************
	'* getValue 
	'**********************************************************************************************************
	private function getValue(xIsRow, x, y)
		if xIsRow then
			getValue = values(x, y)
		else
			getValue = values(y, x)
		end if
	end function
	
	'**********************************************************************************************************
	'* init 
	'**********************************************************************************************************
	private sub init()
		if isObject(datasource) then
			if datasource is nothing then lib.error("Datasource must be assigned.")
		else
			if datasource = "" then lib.error("Datasource cannot be an empty string.")
		end if
		
		if not isObject(datasource) then set datasource = lib.getRecordset(datasource)
		
		lib.page.loadJavascriptFile(classLocation & "js/util.js")
		lib.page.loadJavascriptFile(classLocation & "js/table.js")
		lib.page.loadStylesheetFile classLocation & "report.css", empty
		lib.page.loadStylesheetFile cssLocation, empty
		
		if stringBuilder then
			set output = server.createObject("StringBuilderVB.StringBuilder")
			output.Init 40000, 7500
		end if
		
		'if there is no groupby then it means that there are no automatic cols
		if groupBy = "" then exit sub
		
		'we need to put the added section to the end, so we need to remember the added ones
		set addedParams = server.createObject("scripting.dictionary")
		if swapAxes then
			set fields = rows
		else
			set fields = cols
		end if
		'add the added cols to the end
		for each s in fields.items
			addedParams.add lib.getUniqueID(), s
		next
		fields.removeAll()
		'determine the cols
		if not datasource.eof then
			while not datasource.eof
				currentSection = datasource(groupBy) & ""
				set s = new ReportParameter
				s.caption = currentSection
				s.name = currentSection
				addCol(s)
				datasource.movenext()
			wend
			datasource.movefirst()
		end if
		'add the added cols to the end
		for each s in addedParams.items
			addCol(s)
		next
		set addedParams = nothing
		
		'this matrix stores the currently rendered values
		'its needed for the calculation with formulas, we need to remember the values
		redim values(rows.count - 1, cols.count - 1)
	end sub
	
	'******************************************************************************************************************
	'* drawHeader 
	'******************************************************************************************************************
	private sub drawHeader()
		'draw the report-title
		write("<thead>")
		if title <> empty then write("<tr class=title><td class=noborder></td><td colspan=" & cols.count & ">" & title & "</td></tr>")
		write("<tr>")
		'dummy cell
		write("<td class=dummyCell></td>")
		for each s in cols.items
			write("<td class=""sortable colCaption hand"" onclick=""Table.sort(this,{'sortType':Sort.Numeric,'rowShade':'alternate'})"">")
			write("<a href=""#"">" & s.caption & "</a>")
			write("</td>")
		next
		write("</tr>")
		write("</thead>")
		'is needed even if we dont use it
		write("<tfoot></tfoot>")
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes something to the output
	'******************************************************************************************************************
	public sub write(val)
		if stringBuilder then
			output.append(val)
		else
			str.write(val)
		end if
	end sub

end class
%>