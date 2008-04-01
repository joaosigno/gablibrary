<%
'**************************************************************************************************************

'' @CLASSTITLE:		Calendar
'' @CREATOR:		Michal Gabrukiewicz - gabru @ grafix.at
'' @CREATEDON:		20.07.2004
'' @CDESCRIPTION:	Draws a calendar-control with the possibillity to switch views. weekview, dayview, monthview
''					and yearview. it also allows you to provide information inside the calendar. For example
''					you can place some date (birthdays, events, etc.) on a special day. It also hightlights
''					all the holidays automatically. onDayCreated is obligatory!
'' @VERSION:		1.1
'' @REQUIRES:		Dropdown

'**************************************************************************************************************
'TODO: 
'	BUG WITH DOCUMENTOR, neet to check this, ebcause documentor does not recognise this as a
' 	class because of the include files
%>
<!--#include virtual="/gab_Library/class_dates/dates.asp"-->
<!--#include virtual="/gab_LibraryConfig/_calendar.asp"-->
<!--#include file="class_calendarDay.asp"-->
<!--#include file="class_dayMenuitem.asp"-->
<!--#include file="language.asp"-->
<!--#include file="config.asp"-->
<%
class Calendar

	'private members
	private datesObject
	private dateToday
	private classLocation
	private p_currentView
	private firstDay
	private numberOfDisplayedYears
	private imgSrcSwitchDayView
	private imgSrcSwitchWeekView
	private imgSrcSwitchMonthView
	private imgSrcSwitchYearView
	private imgSrcGoToday
	private imgSrcGoToDate
	private dayMenuitems
	private datePickerUrl
	private todayIsOnScreen
	
	'public members
	public defaultView				''[calendarday-enum] what view should be shown by defualt?
									''CALENDARVIEW_DAY, CALENDARVIEW_WEEK, CALENDARVIEW_MONTH
	public defaultStylesheet		''[bool] use the default styles for this control. default = true
	public selectedDate				''[date] which date should be selected. If nothing given then today-date will be displayed
	public enabledCalendarViews		''[binary] which calendarviews are allowed. Use logical connection of the calendarview-enum
									''e.g. CALENDARVIEW_DAY or CALENDARVIEW_MONTH enables month and day view. default = all views
	public showClock				''[bool] show-clock on todays day? default = true
	public runningClock				''[bool] should the clock be automatically running? default = true
	public timeline					''[bool] show timeline. currently only in day-view. default = false
	public enableWeekends			''[bool] show the weekends? defualt = true
	public cssLocation				''[string] the path and file name to the css file. Gets the default from the config.asp
	public onPreDayCreated			''[sub] implement this function to have access to a day before it gets drawn (has byRef calendarDay as parameter)
									'' Note: do not draw anything here !! Only do calculations and assigning here
	
	public property get currentView ''[calendarday-enum] gets the currentView
		currentView = p_currentView
	end property
	
	private property let currentView(value)
		p_currentView = CALENDARVIEW_NO
		if isNumeric(value) then p_currentView = cInt(value)
	end property
	
	public property get getFirstDay ''[date] the first day, e.g. of the week
		getFirstDay = firstday
	end property
	
	'***********************************************************************************************************
	'* constructor 
	'***********************************************************************************************************
	public sub class_Initialize()
		lib.require("Dropdown")
		set datesObject				= new dates
		defaultView					= CALENDARVIEW_WEEK
		p_currentView				= CALENDARVIEW_NO
		defaultStylesheet			= true
		dateToday					= date()
		cssLocation					= lib.init(GL_CAL_CSSLOCATION, CAL_CLASSLOCATION & "standard.css")
		classLocation				= CAL_CLASSLOCATION
		selectedDate				= empty
		firstDay					= empty
		numberOfDisplayedYears		= 10
		imgSrcSwitchDayView			= classLocation & "icons/icon_dayview_0.gif"
		imgSrcSwitchWeekView		= classLocation & "icons/icon_weekview_0.gif"
		imgSrcSwitchMonthView		= classLocation & "icons/icon_monthview_0.gif"
		imgSrcSwitchYearView		= classLocation & "icons/icon_yearview_0.gif"
		imgSrcGoToday				= classLocation & "icons/icon_goToday_0.gif"
		imgSrcGoToDate				= "/gab_Library/class_datePicker/icons/icon_0.gif"
		set dayMenuitems			= Server.createObject("Scripting.Dictionary")
		datePickerUrl				= "/gab_Library/class_datePicker/index.asp"
		enabledCalendarViews		= CALENDARVIEW_DAY or CALENDARVIEW_WEEK or CALENDARVIEW_MONTH
		showClock					= true
		runningClock				= true
		todayIsOnScreen				= false
		timeline					= false
		enableWeekends				= true
		onPreDayCreated				= empty
	end sub
	
	'***********************************************************************************************************
	'* destructor 
	'***********************************************************************************************************
	private sub class_terminate()
		set datesObject = nothing
		set dayMenuitems = nothing
	end sub
	
	'***********************************************************************************************************
	'' @SDESCRIPTION:	Draws the calendar
	'***********************************************************************************************************
	public sub draw()
		initJavascript()
		initStyles()
		initValues()
		printDayMenu()
		printHeader()
		printCalendar()
		printFooter()
		checkClock()
	end sub
	
	'***********************************************************************************************************
	'' @SDESCRIPTION:	Adds a menuitem to the daymenu
	'' @PARAM:			dayMenuitemObject [dayMenuitemObject]
	'***********************************************************************************************************
	public sub addDayMenuitem(dayMenuitemObject)
		dayMenuitems.add lib.getUniqueID(), dayMenuitemObject
	end sub
	
	'***********************************************************************************************************
	'' @SDESCRIPTION:	Adds a seperator to the dayMenu
	'***********************************************************************************************************
	public sub addDayMenuSeperator()
		set item = new dayMenuitem
		item.caption = "<hr>"
		item.hoverEffect = false
		addDayMenuitem(item)
		set item = nothing
	end sub
	
	'***********************************************************************************************************
	'' @SDESCRIPTION:	checks if a given weekday lies on a weekend
	'' @PARAM:			aWeekday [int]: weekday number.
	'' @RETURN: 		[bool] true if it lies on weekend
	'***********************************************************************************************************
	public function isWeekend(aWeekday)
		isWeekend = datesObject.isWeekend(aWeekday)
	end function
	
	'***********************************************************************************************************
	' checkClock 
	'***********************************************************************************************************
	private sub checkClock()
		if runningClock and showClock and todayIsOnScreen and (currentView and (CALENDARVIEW_DAY or CALENDARVIEW_WEEK)) then
			lib.execJS("updateClock();")
		end if
	end sub
	
	'***********************************************************************************************************
	'* drawSitchViewButton 
	'***********************************************************************************************************
	private sub printCalendar()
		str.writeln("<table cellpadding='0' cellspacing='0' border='0' class='calendar' align='center'>")
		
		select case currentView
			case CALENDARVIEW_DAY
				printDayWeekView()
			case CALENDARVIEW_WEEK
				printDayWeekView()
			case CALENDARVIEW_MONTH
				printMonthView()
			case CALENDARVIEW_YEAR
				printYearView()
		end select
		
		str.writeln("</table>")
	end sub
	
	'***********************************************************************************************************
	'* initJavascript
	'***********************************************************************************************************
	private sub initJavascript()
		str.writeln("<script language=""JavaScript"" src=""" & classLocation & "javascript.js""></script>")
	end sub
	
	'***********************************************************************************************************
	'* initValues 
	'***********************************************************************************************************
	sub initValues()
		
		'init current-calenderView
		currentView = request.queryString("calendarView")
		if currentView = CALENDARVIEW_NO then
			sessionCalendarView = session("gabLib_calendar_calendarView")
			currentView = lib.iif(sessionCalendarView <> "", sessionCalendarView, defaultView)
		end if
		
		'we have to check if the wanted view is allowed. if not we set defaultview
		if not enabledCalendarViews and currentView then currentView = defaultView
		
		'init selectedDate
		selectedDateQs = request.queryString("selectedDate")
		if selectedDateQs = empty then
			sessionSelectedDate = session("gabLib_calendar_selectedDate")
			if sessionSelectedDate <> "" then
				selectedDate = dateValue(sessionSelectedDate)
			else
				selectedDate = dateToday
			end if
		else
			selectedDate = dateValue(selectedDateQs)
		end if
		
		'init firstDay. its the day which will be displayed first in a view
		select case currentView
			case CALENDARVIEW_DAY
				firstDay = selectedDate
			case CALENDARVIEW_WEEK
				adjust = weekDay(selectedDate, 2) - 1
		        firstDay = dateAdd("d", - adjust, selectedDate)
			case CALENDARVIEW_MONTH
				if weekday(dateadd("d", ((day(selectedDate) - 1) * -1), selectedDate), 2) = 1 then
					firstDay = dateadd("d", ((day(selectedDate) - 1 + 7) * -1), selectedDate)
				else
					firstDay = dateadd("d", ((day(selectedDate)-1 + weekday(dateadd("d", ((day(selectedDate))* -1), selectedDate), 2)) * -1), selectedDate)
				end if
			case CALENDARVIEW_YEAR
				firstDay = cDate("1.1." & year(selectedDate))
		end select
		
		'we store the view and the selected-date in a session,
		'so its easy to request the calendar without any params and get the last view
		session("gabLib_calendar_calendarView") = currentView
		session("gabLib_calendar_selectedDate") = selectedDate
	end sub
	
	'***********************************************************************************************************
	'* currentViewIs 
	'***********************************************************************************************************
	private function currentViewIs(view)
		currentViewIs = (enabledCalendarViews and view)
	end function
	
	'***********************************************************************************************************
	'* printHeader
	'***********************************************************************************************************
	private sub printHeader()
		with str
			.writeln("<form name=""calendarForm"" style=""display:inline;"">")
			.writeln("<div class=""calendarHeadline"">")
			.writeln("<span class=""notForPrint"">")
			
			drawGoToDateButton()
			drawGoToTodayButton()
			
			if currentViewIs(CALENDARVIEW_DAY) then
				drawSitchViewButton "switchViewDay", imgSrcSwitchDayView, CALENDARVIEW_DAY, LANG_SIWTCH_DAY
			end if
			if currentViewIs(CALENDARVIEW_WEEK) then
				drawSitchViewButton "switchViewWeek", imgSrcSwitchWeekView, CALENDARVIEW_WEEK, LANG_SIWTCH_WEEK
			end if
			if currentViewIs(CALENDARVIEW_MONTH) then
				drawSitchViewButton "switchViewMonth", imgSrcSwitchMonthView, CALENDARVIEW_MONTH, LANG_SIWTCH_MONTH
			end if
			if currentViewIs(CALENDARVIEW_YEAR) then
				drawSitchViewButton "switchViewYear", imgSrcSwitchYearView, CALENDARVIEW_YEAR, LANG_SIWTCH_Year
			end if
			
			.writeln("&nbsp;")
			.writeln("</span>")
			
			drawWeekNavigation()
			drawDayNavigation()
			drawMonthNavigation()
			drawYearNavigation()
			
			.writeln("</div>")
		end with
	end sub
	
	'***********************************************************************************************************
	'* printFooter 
	'***********************************************************************************************************
	private sub printDayMenu()
		with str
			.writeln("<div id=""dayMenu"" onmousemove=""document.onclick = hideDayMenu;"">")
			.writeln("<input type=""Hidden"" value="""" id=""dayMenuClickedDate"" name=""dayMenuClickedDate"">")
			.writeln("<div id=""dayMenuItemHeadline"" class=""dayMenuItem cursorCommon""></div>")
			
			if dayMenuitems.count > 0 then
				for each menuitem in dayMenuitems.items
					menuitem.draw()
				next
			end if
			
			'Go today item
			set item = new DayMenuitem
			item.caption = " <img src=""" & imgSrcGoToDay & """ border=""0"" align=""absmiddle"" height=""15"" width=""15"">&nbsp;" & LANG_GOTOTODAY
			item.onClick = "goToUrl('" & getScriptUrl(currentView, lib.custom.toCustomDateFormat(dateToday)) & "');"
			item.toolTip = LANG_GOTOTODAYHELP
			item.draw()
			set item = nothing
			
			'jump to date item
			set item = new DayMenuitem
			item.caption = "<img src=""" & imgSrcGoToDate & """ border=""0"" align=""absmiddle"" height=""15"" width=""15"" >&nbsp;" & LANG_GOTODATE
			item.onClick = "goToDate('" &  getScriptUrl(currentView, "dateToChange") & "', '" & selectedDate & "', '" & datePickerUrl & "?JSTarget=calendarForm.dummyGoToDateField')"
			item.toolTip = LANG_GOTODATEHELP
			item.draw()
			set item = nothing
			
			'day-view item
			if currentViewIs(CALENDARVIEW_DAY) then
				set item = new DayMenuitem
				item.caption = "<img src=""" & imgSrcSwitchDayView & """ border=""0"" align=""absmiddle"" height=""15"" width=""15"">&nbsp;" & LANG_SIWTCH_DAY
				item.onClick = "dayMenuGoToUrl('" & getScriptUrl(CALENDARVIEW_DAY, "dateToChange") & "')"
				item.toolTip = LANG_SIWTCH_DAY
				item.disabled = (currentView = CALENDARVIEW_DAY)
				item.draw()
				set item = nothing
			end if
			
			'week-view item
			if currentViewIs(CALENDARVIEW_WEEK) then
				set item = new DayMenuitem
				item.caption = "<img src=""" & imgSrcSwitchWeekView & """ border=""0"" align=""absmiddle""  height=""15"" width=""15"">&nbsp;" & LANG_SIWTCH_WEEK
				item.onClick = "dayMenuGoToUrl('" & getScriptUrl(CALENDARVIEW_WEEK, "dateToChange") & "')"
				item.toolTip = LANG_SIWTCH_WEEK
				item.disabled = (currentView = CALENDARVIEW_WEEK)
				item.draw()
				set item = nothing
			end if
			
			'Month-view item
			if currentViewIs(CALENDARVIEW_MONTH) then
				set item = new DayMenuitem
				item.caption = "<img src=""" & imgSrcSwitchMonthView & """ border=""0"" align=""absmiddle""  height=""15"" width=""15"">&nbsp;" & LANG_SIWTCH_MONTH
				item.onClick = "dayMenuGoToUrl('" & getScriptUrl(CALENDARVIEW_MONTH, "dateToChange") & "')"
				item.toolTip = LANG_SIWTCH_MONTH
				item.disabled = (currentView = CALENDARVIEW_MONTH)
				item.draw()
				set item = nothing
			end if
			
			'Month-view item
			if currentViewIs(CALENDARVIEW_YEAR) then
				set item = new DayMenuitem
				item.caption = "<img src=""" & imgSrcSwitchYearView & """ border=""0"" align=""absmiddle""  width=""15"" height=""15"" border=""0"">&nbsp;" & LANG_SIWTCH_YEAR
				item.onClick = "dayMenuGoToUrl('" & getScriptUrl(CALENDARVIEW_YEAR, "dateToChange") & "')"
				item.toolTip = LANG_SIWTCH_YEAR
				item.disabled = (currentView = CALENDARVIEW_YEAR)
				item.draw()
				set item = nothing
			end if
			
			.writeln("</div>")
		end with
	end sub
	
	'***********************************************************************************************************
	'* printFooter 
	'***********************************************************************************************************
	private sub printFooter()
		str.writeln("<div class=""footline""></div>")
		str.writeln("</form>")
	end sub
	
	'***********************************************************************************************************
	'* getNewNavigationDD 
	'***********************************************************************************************************
	private function getNewNavigationDD(name, values, captions)
		set getNewNavigationDD = new Dropdown
		with getNewNavigationDD
			.name = name
			.datasource = captions
			.valuesDatasource = values
			.selectedValue = selectedDate
			.attributes = "onChange=""goToUrl('" & getScriptUrl(currentView, empty) & "selectedDate=' + this.value);"""
		end with
	end function
	
	'***********************************************************************************************************
	'* drawYearNavigation 
	'***********************************************************************************************************
	sub drawYearNavigation()
		currentYear = year(selectedDate)
		str.write("<span class=""navi"">")
		drawNAVIButtonPrev "yearPrev", dateAdd("yyyy", -1, selectedDate), LANG_PREV_YEAR
		
		values = array()
		captions = array()
		for i = currentYear - numberOfDisplayedYears to currentYear + numberOfDisplayedYears
			redim preserve values(uBound(values) + 1)
			values(uBound(values)) = dateAdd("yyyy", i - currentYear, selectedDate)
			redim preserve captions(uBound(captions) + 1)
			captions(uBound(captions)) = i
		next
		set dd = getNewNavigationDD("navYear", values, captions)
		dd.draw()
		set dd = nothing
		
		drawNAVIButtonNext "yearNext", dateAdd("yyyy", +1, selectedDate), LANG_NEXT_YEAR
		str.write("</span>")
	end sub
	
	'***********************************************************************************************************
	'* drawMonthNavigation
	'***********************************************************************************************************
	private sub drawMonthNavigation()
		str.write("<span class=""navi"">")
		drawNAVIButtonPrev "monthPrev", dateAdd("m", -1, selectedDate), LANG_PREV_MONTH
		
		values = array() : redim values(11)
		captions = array() : redim captions(11)
		for i = 0 to 11
			values(i) = dateAdd("m", (i + 1) - month(selectedDate), selectedDate)
			captions(i) = MONTH_NAMES_SHORT(i + 1)
		next
		set dd = getNewNavigationDD("navMonth", values, captions)
		dd.draw()
		set dd = nothing
		
		drawNAVIButtonNext "monthNext", dateAdd("m", +1, selectedDate), LANG_NEXT_MONTH
		str.write("</span>")
	end sub
	
	'***********************************************************************************************************
	'* drawDayNavigation 
	'***********************************************************************************************************
	sub drawDayNavigation()
		str.write("<span class=navi>")
		
		'we need to check if the next day is not a weekend if weekends are disabled
		prevDay = dateAdd("d", -1, selectedDate)
		if not enableWeekends then
			while isWeekend(weekday(prevDay))
				prevDay = dateAdd("d", -1, prevDay)
			wend
		end if
		drawNAVIButtonPrev "dayPrev", prevDay, LANG_PREV_DAY
		
		values = array()
		captions = array()
		currentMonth = month(selectedDate)
		for i = 1 to 31
			currentDay = dateAdd("d", i - 1, "01/" & month(selectedDate) & "/" & year(selectedDate))
			if month(currentDay) = currentMonth then
				if enableWeekends or (not enableWeekends and not isWeekend(weekday(currentDay))) then
					redim preserve values(uBound(values) + 1)
					values(uBound(values)) = currentDay
					redim preserve captions(uBound(captions) + 1)
					captions(uBound(captions)) = i
				end if
			else
				exit for
			end if
		next
		set dd = getNewNavigationDD("navDay", values, captions)
		dd.draw()
		set dd = nothing
		
		nextDay = dateAdd("d", +1, selectedDate)
		if not enableWeekends then
			while isWeekend(weekday(nextDay))
				nextDay = dateAdd("d", +1, nextDay)
			wend
		end if
		drawNAVIButtonNext "dayNext", nextDay, LANG_NEXT_DAY
		str.write("</span>")
	end sub
	
	'***********************************************************************************************************
	'* drawWeekNavigation 
	'***********************************************************************************************************
	sub drawWeekNavigation()
		with str
			.write("<span class=""navi"">")
			drawNAVIButtonPrev "weekPrev", dateAdd("ww", -1, selectedDate), LANG_PREV_WEEK
			
			weekNumber = datesObject.weekOfTheYear(selectedDate)
			upperBound = datesObject.weekOfTheYear(dateValue("31/12/" & year(selectedDate)))
			'maybe the weeknumber is already the first in the next year
			if upperBound = 1 then upperBound = datesObject.weekOfTheYear(dateadd("ww", - 1, dateValue("31/12/" & year(selectedDate))))
			values = array()
			captions = array()
			for i = 1 to upperBound
				redim preserve values(uBound(values) + 1)
				values(uBound(values)) = dateAdd("ww", i - weekNumber, selectedDate)
				redim preserve captions(uBound(captions) + 1)
				captions(uBound(captions)) = LANG_WEEK & i
			next
			set dd = getNewNavigationDD("navDropdown", values, captions)
			dd.draw()
			set dd = nothing
			
			drawNAVIButtonNext "weekNext", dateAdd("ww", +1, selectedDate), LANG_NEXT_WEEK
			.write("</span>")
		end with
	end sub
	
	'***********************************************************************************************************
	'* drawNAVIButtonPrev 
	'***********************************************************************************************************
	private sub drawNAVIButtonPrev(name, gotoDate, tooltip)
		str.write("<button type=""button"" class=""navButton notForPrint"" name=""" & name & """ onclick=""goToUrl('" & getScriptUrl(currentView, gotoDate) & "')"" title=""" & tooltip & """>&lt;</button>")
	end sub
	
	'***********************************************************************************************************
	'* drawNAVIButtonNext 
	'***********************************************************************************************************
	private sub drawNAVIButtonNext(name, gotoDate, tooltip)
		str.write("<button type=""button"" class=""navButton notForPrint"" name=""" & name & """ onclick=""goToUrl('" & getScriptUrl(currentView, gotoDate) & "')"" title=""" & tooltip & """>&gt;</button>")
	end sub
	
	'***********************************************************************************************************
	'* drawGoToTodayButton 
	'***********************************************************************************************************
	private sub drawGoToTodayButton()
		onClick = "goToUrl('" & getScriptUrl(currentView, lib.custom.toCustomDateFormat(dateToday)) & "');"
		str.writeln("<button type=""button"" class=""icon"" title=""" & LANG_GOTOTODAYHELP & """ onclick=""" & onClick & """>" &_
					"<img src=""" & imgSrcGoToDay & """ width=""15"" height=""15"" border=""0"" ></button>")
	end sub
	
	'***********************************************************************************************************
	'* drawGoToDateButton 
	'***********************************************************************************************************
	private sub drawGoToDateButton()
		str.write("<input type=Hidden value=0 id=""dummyGoToDateField"" name=""dummyGoToDateField""> ")
		onClick = "goToDate('" &  getScriptUrl(currentView, "dateToChange") & "','" & selectedDate & "', '" & datePickerUrl & "?JSTarget=calendarForm.dummyGoToDateField');"
		str.write(" <button type=""button"" tabindex=""-1"" width=15 height=15 border=0  class=""icon"" onClick=""" & onClick & """ title=""" & LANG_GOTODATEHELP & """><img src=""" & imgSrcGoToDate & """ border=0></button>")
	end sub
	
	'***********************************************************************************************************
	'* drawSitchViewButton 
	'***********************************************************************************************************
	private sub drawSitchViewButton(name, imgSrc, enumView, toolTip)
		str.writeln(" <button type=""button"" name=""" & name & """ class=""icon"" " &_
						"onclick=""goToUrl('" & getScriptUrl(enumView, selectedDate) & "');"">" &_
						"<img src=""" & imgSrc & """ border=0 title=""" & toolTip & """ width=""15"" height=""15"" border=""0""></button>")
	end sub
	
	'***********************************************************************************************************
	'* getNewCalendarDay 
	'***********************************************************************************************************
	private function getNewCalendarDay(index, dat)
		set getNewCalendarDay = new CalendarDay
		with getNewCalendarDay
			.index = index
			.dat = dat
		end with
	end function
	
	'***********************************************************************************************************
	'* printYearView 
	'***********************************************************************************************************
	private sub printYearView()
		str.writeln("<tr>")
		
		title = "<strong>{1}</strong>"
		if currentViewIs(CALENDARVIEW_MONTH) then
			title = "<a href=" & getScriptUrl(CALENDARVIEW_MONTH, day(selectedDate) & ".{0}." & year(selectedDate)) & ">" & title & "</a>"
		end if
		
		for i = 1 to 12
			printDayHeader str.format(title, array(i, MONTH_NAMES(i))), empty, MONTH_NAMES(i), false, empty, empty
		next
		str.writeln("</tr>")
		
		index = 0
		for i = 1 to 31
			str.writeln("<tr valign=top class=""yearDays"">")
			for j = 1 to 12
				cssClass = empty
				toolTip = empty
				isValidDay = isDate(i & "." & j & "." & year(selectedDate))
				if not isValidDay then cssClass = "monthCorner"
				
				if isValidDay then
					set aDay = getNewCalendarDay(index, dateAdd("d", i - 1, dateadd("m", j - 1, firstDay)))
					if onPreDayCreated <> empty then execute(onPreDayCreated & "(aDay)")
					toolTip = aDay.getWeekdayName() & ", " & lib.custom.toCustomDateFormat(aDay.dat)
					if aDay.isWeekend() then cssClass = "calendarHoliday"
					
					if aDay.isHoliday() then
						cssClass = cssClass & " calendarHoliday"
						toolTip = toolTip & vbcrlf & "(" & aDay.getHolidayName() & ")"
					end if
					
					if aDay.isToday() then
						todayIsOnScreen = true
						cssClass = cssClass & " calendarToday"
						toolTip = toolTip & " - " & ucase(LANG_TODAY)
					end if
					index = index + 1
				end if
				
				str.write("<td class=""" & cssClass & """>")
				if isValidDay then
					onClick = "showDayMenu('" & getScriptUrl(CALENDARVIEW_DAY, aDay.dat) & "', '" & lib.custom.toCustomDateFormat(aDay.dat) & "', '" & aDay.getWeekdayName() & "')"
					str.write("<span class=""monthDayText"" onmouseover=""menuHoverIn(this, true)"" onmouseout=""menuHoverOut(this, true)"" title=""" & tooltip & """ onclick=""" & onClick & """>" & i & "</span>")
					onDayCreated(aDay)
				end if
				str.write("</td>")
			next
			str.writeln("</tr>")
		next
	end sub
	
	'***********************************************************************************************************
	'* printMonthView 
	'***********************************************************************************************************
	private sub printMonthView()
		str.writeln("<tr>")
		str.writeln("<td class=""monthCorner"" style='width:18px; margin:0px; padding-right:2px; '></td>")
		for i = 1 to 7
			index = lib.iif(i = 7, 1, i + 1)
			printDayHeader "<strong>" & WEEKDAY_NAMES(index) & "</strong>", empty, WEEKDAY_NAMES(weekDayNumber), isWeekend(index), empty, lib.iif(enableWeekends, "seventhOfPage", "fifthOfPage")
		next
		str.writeln("</tr>")
		
		for i = 0 to 5
			firstDayOfNewWeek = dateadd("d", i * 7, firstDay)
			currentWeek = datesObject.weekOfTheYear(firstDayOfNewWeek)
			
			weekLink = currentWeek
			if currentViewIs(CALENDARVIEW_WEEK) then
				weekLink = "<a href=" & getScriptUrl(CALENDARVIEW_WEEK, firstDayOfNewWeek) & ">" & currentWeek & "</a>"
			end if
			
			str.writeln("<tr valign=top>")
			str.writeln("<td  class=""weekNumber"" title=""" & currentWeek & " " & LANG_WEEK_OF_YEAR & """>" & weekLink & "</td>")
			for j = 0 to 6
				index = j + (i * 7)
				set currentDay = getNewCalendarDay(index, dateadd("d", index, firstDay))
				if onPreDayCreated <> empty then execute(onPreDayCreated & "(currentDay)")
				isInSelectedMonth = month(currentDay.dat) = month(selectedDate)
				cssClass2 = empty
				cssClass = empty
				
				if (currentDay.isWeekend() and enableWeekends) or not currentDay.isWeekend() then 
					toolTip = WEEKDAY_NAMES(weekday(currentDay.dat)) & ", " & lib.custom.toCustomDateFormat(currentDay.dat)
					displayDate = day(currentDay.dat)
					
					'we check if the day is in the wanted month
					if isInSelectedMonth then
						cssClass = empty
						if currentDay.isWeekend() then cssClass = " calendarHoliday"
					else
						cssClass = " otherMonth"
						cssClass2 = " otherMonthText"
					end if
					
					'if the day is a holiday
					if currentDay.isHoliday() then
						cssClass = cssClass & " calendarHoliday"
						cssClass2 = " holidayName"
						if not isInSelectedMonth then cssClass2 = cssClass2 & " holidayNameOtherMonth"
						displayDate = currentDay.getHolidayName()
						toolTip = toolTip & vbcrlf & "(" & displayDate & ")"
					end if
					
					'if the currentday is today
					if currentDay.isToday() then
						todayIsOnScreen = true
						if currentDay.isHoliday() or not isInSelectedMonth then
							cssClass = " calendarToday"
						else
							cssClass = cssClass & " calendarToday"
						end if
						toolTip = toolTip & " - " & ucase(LANG_TODAY)
					end if
					
					onClick = "showDayMenu('" & getScriptUrl(CALENDARVIEW_DAY, currentDay.dat) & "', '" & lib.custom.toCustomDateFormat(currentDay.dat) & "', '" & currentDay.getWeekdayName() & "')"
					str.writeln("<td class=""monthDay" & cssClass & """ " & currentDay.attributes & ">")
					str.writeln("<div onmouseover=""menuHoverIn(this, true)"" onmouseout=""menuHoverOut(this, true)"" class=""monthDayText" & cssClass2 & """ title=""" & toolTip & """ onclick=""" & onClick & """ " & currentDay.HLattributes & ">" & displayDate & "</div>")
					onDayCreated(currentDay)
					str.writeln("</td>")
				end if
				
				set currentDay = nothing
			next
			str.writeln("</tr>")
		next
	end sub
	
	'***********************************************************************************************************
	'* printWeekView 
	'***********************************************************************************************************
	private sub printDayWeekView()
		oppositeView = CALENDARVIEW_DAY + CALENDARVIEW_WEEK - currentView
		upperBound = lib.iif(currentView = CALENDARVIEW_DAY, 1, 7)
		
		'print out weekdaynames
		str.writeln("<tr>")
		for i = 1 to upperBound
			newDay = lib.custom.toCustomDateFormat(dateAdd("d", i - 1, firstDay))
			weekDayNumber = weekday(newDay)
			
			printDayHeader "<strong>" & WEEKDAY_NAMES(weekDayNumber) & "</strong><br>" & newDay, _
							getScriptUrl(oppositeView, newDay), _
							WEEKDAY_NAMES(weekDayNumber), _
							isWeekend(weekDayNumber), newDay, _
							lib.iif(enableWeekends, "seventhOfPage", "fifthOfPage")
		next
		str.writeln("</tr>")
		
		'print the timeline
		str.writeln("<tr class=weekViewDay>")
		for i = 1 to upperBound
			cssClass = empty
			
			set currentDay = getNewCalendarDay(i - 1, dateAdd("d", i - 1, firstDay))
			if onPreDayCreated <> empty then execute(onPreDayCreated & "(currentDay)")
			
			if (currentDay.isWeekend() and enableWeekends) or not currentDay.isWeekend() then 
				if currentDay.isToday() then
					todayIsOnScreen = true
					cssClass = " calendarToday"
				end if
				
				if currentDay.isWeekend() or currentDay.isHoliday() then cssClass = " calendarHoliday"
				
				str.writeln("<td class=""calendarDay" & cssClass & """ valign=top " & currentDay.attributes & ">")
				if currentDay.isHoliday() then str.write("<div class=holidayName>" & currentDay.getHolidayName() & "</div>")
				if currentDay.isToday() and showClock then
					currentTime = "init clock ..."
					if not runningClock then
						currentHour = right("00" & hour(now()), 2)
						currentMinute = right("00" & minute(now()), 2)
						currentTime = currentHour & ":" & currentMinute
					end if
					str.write("<div class=todaysTime id=todayClock>" & currentTime & "</div>")
				end if
				showTimeLine()
				onDayCreated(currentDay)
				showTimeLine()
				str.writeln("</td>")
			end if
			set currentDay = nothing
		next
		str.writeln("</tr>")
	end sub
	
	'***********************************************************************************************************
	'* printDayHeader 
	'***********************************************************************************************************
	private sub printDayHeader(title, url, toolTip, isWeekend, dat, cssClass)
		if isWeekend and enableWeekends or not isWeekend then
			cssClass = lib.iif(isWeekend, "weekdayNameWeekend", empty) & " " & cssClass
			
			if url <> empty then
				onClick = " onClick=""showDayMenu('" & getScriptUrl(CALENDARVIEW_DAY, dateToday) & "', '" & dat & "', '" & WEEKDAY_NAMES(weekday(dat)) & "')"""
				onElse = " onmouseover=""menuHoverIn(this)"" onmouseout=""menuHoverOut(this)"""
				cssClass = cssClass & " handCursor"
			end if
			
			str.writeln("<td" & onClick & " class=""weekdayName whiteRightBorder " & cssClass & """ title=""" & toolTip & """><div" & onElse & ">" & title & "</div></td>")
		end if
	end sub
	
	'******************************************************************************************
	'* showTimeLine 
	'******************************************************************************************
	sub showTimeLine()
		if timeline and currentView = CALENDARVIEW_DAY then
			str.write("<div class=calendarTimeline>")
			for i = 0 to 5
				str.write("<span class=""calendarTimelineHour"">" & right("00" & i * 4, 2) & ":00</span>")
			next
			str.write("</div>")
		end if
	end sub
	
	'***********************************************************************************************************
	'* getScriptUrl 
	'***********************************************************************************************************
	private function getScriptUrl(calendarView, dateSelected)
		queryString = lib.getAllFromQueryStringBut("calendarView,selectedDate")
		scriptName = request.serverVariables("SCRIPT_NAME") & "?"
		
		if not queryString = empty then scriptName = scriptName & queryString & "&"
		if not calendarView = empty then scriptName = scriptName & "calendarView=" & calendarView & "&"
		if not dateSelected = empty then scriptName = scriptName & "selectedDate=" & dateSelected & "&"
		
		getScriptUrl = scriptName
	end function
	
	'***********************************************************************************************************
	'* initStyles
	'***********************************************************************************************************
	private sub initStyles()
		if defaultStylesheet then lib.page.loadStylesheetFile cssLocation, empty
	end sub

end class
lib.registerClass("Calendar")
%>