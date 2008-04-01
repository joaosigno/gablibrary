<!--#include virtual="/gab_Library/class_dates/dates.asp"-->
<!--#include virtual="/gab_LibraryConfig/_datePicker.asp"-->
<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		datePicker
'' @CREATOR:		Michal Gabrukiewicz / Michael Rebec - gabru @ grafix.at
'' @CREATEDON:		12.07.2004
'' @CDESCRIPTION:	Draw a datePicker-control. It has a lot of features inside like
''					autmatically display holidays, tooltips, maximum and minimum-date selection, shortcuts,
''					etc. Easy to use.
'' @VERSION:		0.2

'**************************************************************************************************************
class datePicker

	private datesObject
	
	public title				''[string] The title which should be displayed on the calendar
	public displayHolidays		''[bool] Highlight the holidays in the calendar. default = true
	public defaultStylesheet	''[bool] use the default styles for this control. default = true
								''if you want to implement your own styles then look at the classes in the standard.css
	public selectedDate			''[date] which date is selected. if no date is given then todays date will be selected
	public shownYearRange		''[int] how many years should be displayed before/after the selected year? default = 50
	public autoResize			''[bool] automatically remove lines if there is no day of the selected month/year. default = false
	public JSTarget				''[string] the target of the field you want to input the selected date. e.g: frm.date
	public maximumAllowedDate	''[date] the maximum allowed date. if empty then every date is selectable to the upper-side
	public minimumAllowedDate	''[date] the minimum allowed date. if empty then every date is selectable to the lower-side
	public autoDisableNavi		''[bool] should the navigation be disabled if the allowed date cannot be reached
								''e.g. if the previous month is not allowed the back-month-button will be disabled, etc.
	public useStringBuilder		''[bool] use stringbuilder. default = true
	public cssLocation			''[string] the path and file name to the css file. Gets the default from the config.asp
	
	private classLocation		'Absolute path of the class itself
	private selectedMonth		'Currently selected Month
	private selectedYear		'Currently selected Year
	private chosenDate			'currently chosen Date of by the user. month and year. day is of today
	private dateToday			
	private originallyLCID		
	private output				
	
	'Konstruktor => set the default values
	private sub Class_Initialize()
		classLocation			= "/gab_Library/class_datePicker/"
		set datesObject			= new dates
		title					= empty
		displayHolidays			= true
		defaultStylesheet		= true
		shownYearRange			= 50
		cssLocation				= lib.init(GLDP_CSS_LOCATION, classLocation & "standard.css")
		selectedDate			= empty
		maximumAllowedDate		= empty
		minimumAllowedDate		= empty
		selectedMonth			= 0
		selectedYear			= 0
		autoResize				= false
		dateToday				= date()
		JSTarget				= empty
		originallyLCID			= session.lcid
		autoDisableNavi			= false
		useStringBuilder		= true
		
		'this is the lcid the calendar uses.
		'the lcid will be set to the old at the end
		session.lcid			= 1031 
	end sub
	
	'Destruktor
	private sub Class_Terminate()
		session.lcid = originallyLCID
		if useStringBuilder then set output = nothing
	end sub
	
	public property get holidayName() ''[string] Returns the name of a holiday which were calculated through isHoliday-method
		holidayName = datesObject.holidayName
	end property
	
	'***********************************************************************************************************
	'' @SDESCRIPTION:	Draws the datepicker-control
	'***********************************************************************************************************
	public sub draw()
		initOutputMethod()
		initJavascript()
		initStyles()
		initValues()
		printHeader()
		printCalendar()
		printFooter()
		'if stringbuilder used we have to write the output
		if useStringBuilder then response.write(output.toString())
	end sub
	
	'***********************************************************************************************************
	'' @SDESCRIPTION:	Formats a date to an individual format.
	'***********************************************************************************************************
	public function formatDate(dateToFormat)
		formatDate = lib.custom.toCustomDateFormat(dateToFormat)
	end function
	
	'***********************************************************************************************************
	'* initOutputMethod 
	'***********************************************************************************************************
	private sub initOutputMethod()
		if useStringBuilder then
			set output = Server.CreateObject("StringBuilderVB.StringBuilder")
			output.init 40000, 7500
		end if
	end sub
	
	'***********************************************************************************************************
	'* initSelectedValues 
	'***********************************************************************************************************
	private sub initValues()
		
		'we cast the string-dates to a date-type
		if not minimumAllowedDate = empty then
			minimumAllowedDate = dateValue(minimumAllowedDate)
		end if
		if not maximumAllowedDate = empty then
			maximumAllowedDate = dateValue(maximumAllowedDate)
		end if
		
		if page.isPostback() then
			chosenDate = dateValue(request.form("chosenDate"))
			selectedMonth = month(chosenDate)
			selectedYear = year(chosenDate)
		else
			chosenDate = dateToday
			
			if selectedDate = empty then
				selectedDate = chosenDate
			else
				'first we have to test if the given value is a date. 
				'because if it wont be a valid date then datePicker wont work.
				on error resume next
					test = cdate(selectedDate)
				if err = 0 then
					'we check if the date is in the max-min-range.
					if dateIsInAllowedRange(selectedDate) then
						chosenDate = selectedDate
					else
						selectedDate = dateToday
					end if
				else
					selectedDate = chosenDate
				end if
			end if
			
			selectedMonth = month(selectedDate)
			selectedYear = year(selectedDate)
		end if
	end sub
	
	'***********************************************************************************************************
	'* print 
	'***********************************************************************************************************
	private sub print(outputString)
		if useStringBuilder then
			output.append(outputString)
		else
			str.writeln(outputString)
		end if
	end sub
	
	'***********************************************************************************************************
	'* printCalendar 
	'***********************************************************************************************************
	private sub printCalendar()
		print("<table border=0 class=""calendar"" cellpadding=0 cellspacing=0 align=center onMousewheel=""checkWheel();"">")
		
		if weekday(dateadd("d", ((day(chosenDate) - 1) * -1), chosenDate), 2) = 1 then
			firstday = dateadd("d", ((day(chosenDate) - 1 + 7) * -1), chosenDate)
		else
			firstday = dateadd("d", ((day(chosenDate)-1 + weekday(dateadd("d", ((day(chosenDate))* -1), chosenDate), 2)) * -1), chosenDate)
		end if
		
		'print out weekdaynames
		print("<tr>")
		print("<td class=weekdayName></td>")
		for i = 2 to 7
			if i = 7 then cssClass = " weekdayNameWeekend"
			print("<td class=""weekdayName" & cssClass & """ title=""" & WEEKDAY_NAMES(i) & """>" & WEEKDAY_SHORTNAMES(i) & "</td>")
		next
		print("<td class=""weekdayName weekdayNameWeekend"" title=""" & WEEKDAY_NAMES(1) & """>" & WEEKDAY_SHORTNAMES(1) & "</td>")
		print("</tr>")
		
		'print the days
		for i = 0 to 5
			showLine = true
			
			if autoResize then
				'we check if there are some line at the beginning of the displayed area who has no day of this month
				if i = 0 and month(dateAdd("d", 6, firstDay)) <> month(chosenDate) then showLine = false
				
				'the same with the last lines
				if i = 5 and month(dateadd("d", i * 7, firstDay)) <> month(chosenDate) then showLine = false
			end if
			
			if showLine then
				currentWeek = datesObject.weekOfTheYear(dateadd("d", i * 7, firstDay))
				print("<tr>")
				print("<td class=weekNumber title=""" & currentWeek & " " & DP_LANG_WEEK_OF_YEAR & """>" & currentWeek & "</td>")
				
				for j = 0 to 6
					currentDay = dateadd("d", j + i * 7, firstDay)
					formatedCurrentDay = formatDate(currentDay)
					toolTip = WEEKDAY_NAMES(weekday(currentDay)) & ", " & formatedCurrentDay
					
					'we check if the day is in the wanted month
					if month(currentDay) <> month(chosenDate) then
						cssClass = " otherMonth"
					else
						cssClass = empty
						'we check if its a weekend day
						if weekday(currentDay) = 7 or weekday(currentDay) = 1 then
							cssClass = " calendarDayWeekend"
						end if
					end if
					
					'if the currentday is today
					if currentDay = dateToday then
						cssClass = cssClass & " calendarDayToday"
						toolTip = toolTip & " - " & ucase(DP_LANG_TODAY)
					end if
					
					'if the day is a holiday
					if isHoliday(currentDay) then
						cssClass = cssClass & " holidayDay"
						toolTip = toolTip & " (" & holidayName & ")"
					end if
					
					'if the day is the selectedDay
					if selectedDate = formatedCurrentDay then
						cssClass = cssClass & " calendarDaySelected"
						toolTip = toolTip & " - " & ucase(DP_LANG_CURRENTLY_SELECTED)
					end if
					
					if dateIsInAllowedRange(currentDay) then
						disabled = empty
					else
						disabled = " disabled "
					end if
					
					print("<td class=calendarDayField>")
					print("	<button class=""calendarDay" & cssClass & """ title=""" & toolTip & """ " & disabled & "" &_
									"onmouseover=""hoverIn(this, 'calendarHover'); setDisplay('" & DP_LANG_CLICK_TO_SELECT & "<br>" & WEEKDAY_SHORTNAMES(weekday(currentDay)) & ", " & formatedCurrentDay & "')"" " &_
									"onmouseout=""hoverOut(this); clearDisplay();"" tabindex=""-1"" " &_
									"onclick=""sendDate('" & JSTarget & "', '" & formatedCurrentDay & "')"">" & day(currentDay) & "</button>")
					print("</td>")
				next
				print("</tr>")
			end if
		next
		print("</table>")
	end sub
	
	'***********************************************************************************************************
	'* dateIsInAllowedRange 
	'***********************************************************************************************************
	private function dateIsInAllowedRange(ByVal myDate)
		dateIsInAllowedRange = true
		myDate = dateValue(myDate)
		
		'check the maximumallowed-date
		if not maximumAllowedDate = empty then
			if myDate > maximumAllowedDate then
				dateIsInAllowedRange = false
			end if
		end if
		
		'check the minimumallowed-date
		if not minimumAllowedDate = empty then
			if myDate < minimumAllowedDate then
				dateIsInAllowedRange = false
			end if
		end if
	end function
	
	'***********************************************************************************************************
	'* printFooter 
	'***********************************************************************************************************
	private sub printFooter()
		print("<div class=""endline"">")
		if dateIsInAllowedRange(dateToday) then
			print("	<button class=button onclick=""sendDate('" & JSTarget & "', '" & formatDate(dateToday) & "');"" title=""" & DP_LANG_SELECT_TODAY_HELP & """>" & DP_LANG_SELECT_TODAY & "</button>&nbsp;")
		end if
		print("	<button class=button onclick=""closeMe();"" name=cancelButton title=""" & DP_LANG_CANCEL_HELP & """>" & DP_LANG_CANCEL & "</button>")
		print("</div>")
	end sub
	
	'***********************************************************************************************************
	'* printHeader
	'***********************************************************************************************************
	private sub printHeader()
		
		'its needed to load the datedisplay as fast as we can because maybe the users mouse
		'is over a date and so an error will happend.
		print("<div class=bottom>")
		print("	<span id=dateDisplay></span>")
		print("</div>")
		
		print("<form name=datePickerFrm style=""display:inline"" method=post action='" & request.serverVariables("SCRIPT_NAME") & "?" & request.queryString & "'>")
		if not title = empty then
			print("<div class=hl>" & title & "</div>")
		end if
		
		print("<div class=headline>")
		
		'Month navigation
		drawPrevMonthButton()
		drawMonthDropdown()
		drawNextMonthButton()
		
		'jump to today-button
		call drawNaviButton("jumpToday", DP_LANG_GO_TODAY, DP_LANG_TODAY_HELP, "changeChosenDate('" & dateToday & "')", "jumpToToday", autoDisableNavi and not dateIsInAllowedRange(dateToday))
		
		'Year navigation
		drawPrevYearButton()
		drawYearsDropdown()
		drawNextYearButton()
		
		print("</div>")
		
		'represents the current selected month and year.
		print("<input type=Hidden value=""" & day(dateToday) & "." & selectedMonth & "." & selectedYear & """ name=chosenDate>")
		print("</form>")
	end sub
	
	'***********************************************************************************************************
	'* drawPrevMonthButton 
	'***********************************************************************************************************
	private sub drawPrevMonthButton()
		prevMonth = dateAdd("m", -1, chosenDate)
		disable = false
		if autoDisableNavi then
			if not minimumAllowedDate = empty then
				disable = (year(prevMonth) <= year(minimumAllowedDate)) and (month(prevMonth) < month(minimumAllowedDate))
			end if
		end if
		
		call drawNaviButton("prevMonth", "&lt;", DP_LANG_PREV_MONTH, "changeChosenDate('" & prevMonth & "')", empty, disable)
		
		'we store the value in a hidden-field so we can take the value with javascript to handle the shortcuts.
		print("<input type=Hidden value=""" & prevMonth & """ name=prevMonthValue>")
	end sub
	
	'***********************************************************************************************************
	'* drawNextMonthButton 
	'***********************************************************************************************************
	private sub drawNextMonthButton()
		nextMonth = dateAdd("m", +1, chosenDate)
		disable = false
		if autoDisableNavi then
			if not maximumAllowedDate = empty then
				disable = (year(nextMonth) >= year(maximumAllowedDate)) and (month(nextMonth) > month(maximumAllowedDate))
			end if
		end if
		
		call drawNaviButton("nextMonth", "&gt;", DP_LANG_NEXT_MONTH, "changeChosenDate('" & nextMonth & "')", empty, disable)
		
		'we store the value in a hidden-field so we can take the value with javascript to handle the shortcuts.
		print("<input type=Hidden value=""" & nextMonth & """ name=nextMonthValue>")
	end sub
	
	'***********************************************************************************************************
	'* drawPrevYearButton 
	'***********************************************************************************************************
	private sub drawPrevYearButton()
		prevYear = dateAdd("yyyy", -1, chosenDate)
		disable = false
		if autoDisableNavi then
			if not minimumAllowedDate = empty then
				'if new year is smaller than the minimumallowed-year => disable
				'or if the year is equal to the minimumallowed-year BUT the month is smaller => disable
				disable = (year(prevYear) < year(minimumAllowedDate)) or ((year(prevYear) = year(minimumAllowedDate)) and (month(prevYear) < month(minimumAllowedDate)))
			end if
		end if
		
		call drawNaviButton("prevYear", "&lt;", DP_LANG_PREV_YEAR, "changeChosenDate('" & prevYear & "')", empty, disable)
		
		'we store the value in a hidden-field so we can take the value with javascript to handle the shortcuts.
		print("<input type=Hidden value=""" & prevYear & """ name=prevYearValue>")
	end sub
	
	'***********************************************************************************************************
	'* drawNextYearButton 
	'***********************************************************************************************************
	private sub drawNextYearButton()
		nextYear = dateAdd("yyyy", 1, chosenDate)
		disable = false
		if autoDisableNavi then
			if not maximumAllowedDate = empty then
				'if new year is larger than the maximumallowed-year => disable
				'or if the year is equal to the maximumallowed-year BUT the month is larger => disable
				disable = (year(nextYear) > year(maximumAllowedDate)) or ((year(nextYear) = year(maximumAllowedDate)) and (month(nextYear) > month(maximumAllowedDate)))
			end if
		end if
		
		call drawNaviButton("nextYear", "&gt;", DP_LANG_NEXT_YEAR, "changeChosenDate('" & nextYear & "')", empty, disable)
		
		'we store the value in a hidden-field so we can take the value with javascript to handle the shortcuts.
		print("<input type=Hidden value=""" & nextYear & """ name=nextYearValue>")
	end sub
	
	'***********************************************************************************************************
	'* drawNaviButton 
	'***********************************************************************************************************
	private sub drawNaviButton(name, value, toolTip, onClick, cssClass, disabled)
		if disabled then
			disabled = "disabled"
		else
			disabled = empty
		end if
		
		print("<button tabindex=""-1"" class=""navButton " & cssClass & """ name=" & name & " " & disabled & " " &_
						"title=""" & toolTip & """ onclick=""" & onClick & ";"">" & value & "</button>")
	end sub
	
	'***********************************************************************************************************
	'* drawMonthDropdown
	'***********************************************************************************************************
	private sub drawMonthDropdown()
		set dd = new createDropdown
		with dd
			atLeastOne = false
			for i = 1 to 12
				monthCanBeShown = true
				if autoDisableNavi then
					if not minimumAllowedDate = empty then
						if (selectedYear <= year(minimumAllowedDate)) and (i < month(minimumAllowedDate)) then monthCanBeShown = false
					end if
					if not maximumAllowedDate = empty then
						if (selectedYear >= year(maximumAllowedDate)) and (i > month(maximumAllowedDate)) then monthCanBeShown = false
					end if
				end if
				
				if monthCanBeShown then
					if atLeastOne then
						.sqlQuery = .sqlQuery & ":"
						.pk = .pk & ":"
					end if
					.sqlQuery = .sqlQuery & MONTH_NAMES(i)
					.pk = .pk & i
					atLeastOne = true
				end if
			next
			.name = "chosenMonth"
			.idToMatch = selectedMonth
			.forceArray = true
			.onAttribute = "onChange=""changeChosenDate('01.' + this.value + '." & selectedYear & "')"""
			print(.getAsString())
		end with
	end sub
	
	'***********************************************************************************************************
	'* drawYearsDropdown
	'***********************************************************************************************************
	private sub drawYearsDropdown()
		set dd = new createDropdown
		with dd
			'check the minimum and maximum-ranges in the dropdown
			upperBound = selectedYear + shownYearRange
			lowerBound = selectedYear - shownYearRange
			
			if autoDisableNavi then
				'we have to calculate to bounds in dependency on the min and max-dates.
				if not minimumAllowedDate = empty then
					lowerBound = lowerBound + shownYearRange - datediff("yyyy", minimumAllowedDate, chosenDate)
					if lowerBound = year(minimumAllowedDate) and selectedMonth < month(minimumAllowedDate) then
						lowerBound = lowerBound + 1
					end if
				end if
				if not maximumAllowedDate = empty then
					upperBound = upperBound - shownYearRange - datediff("yyyy", maximumAllowedDate, chosenDate)
					if upperBound = year(maximumAllowedDate) and selectedMonth > month(maximumAllowedDate) then
						upperBound = upperBound - 1
					end if
				end if
			end if
			
			i = lowerBound
			do
				if i <> lowerBound then
					.sqlQuery = .sqlQuery & ":"
					.pk = .pk & ":"
				end if
				.sqlQuery = .sqlQuery & i
				.pk = .pk & i
				i = i + 1
			loop until i = upperBound + 1
			
			.forceArray = true
			.name = "chosenYear"
			.idToMatch = selectedYear
			.onAttribute = "onChange=""changeChosenDate('" & day(dateToday) & "/" & selectedMonth & "/' + this.value)"""
			print(.getAsString())
		end with
	end sub
	
	'***********************************************************************************************************
	'* initStyles
	'***********************************************************************************************************
	private sub initStyles()
		if defaultStylesheet then lib.page.loadStylesheetFile cssLocation, empty
	end sub
	
	'***********************************************************************************************************
	'* initJavascript
	'***********************************************************************************************************
	private sub initJavascript()
		print("<script language=JavaScript src=""" & classLocation & "javascript.js""></script>")
	end sub
	
	'***********************************************************************************************************
	'' @SDESCRIPTION:	OBSOLETE! Use dates.isHoliday instead (dates class)
	'' @PARAM:			theDate [date]: see dates.isHoliday
	'' @RETURN:			[bool] see dates.isHoliday
	'***********************************************************************************************************
	public function isHoliday(theDate)
		isHoliday = datesObject.isHoliday(theDate)
	end function

end class
lib.registerClass("DatePicker")
%>