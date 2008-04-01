<!--#include virtual="/gab_Library/class_datePicker/datePicker.asp"-->
<!--#include file="class_calendarDay.asp"-->
<!--#include file="class_dayMenuitem.asp"-->
<!--#include file="language.asp"-->
<%
'**************************************************************************************************************

'' @CLASSTITLE:		calendar
'' @CREATOR:		Michal Gabrukiewicz - gabru@gmx.at
'' @CREATEDON:		20.07.2004
'' @CDESCRIPTION:	Draws a calendar-control with the possibillity to switch views. weekview, dayview and
''					monthview. it also allows you to provide information inside the calendar. For example
''					you can place some date (birthdays, events, etc.) on a special day.
'' @VERSION:		1.0

'**************************************************************************************************************

const CALENDARVIEW_DAY		= 1
const CALENDARVIEW_WEEK		= 2
const CALENDARVIEW_MONTH	= 4

class calendar

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
	
	private dateToday				
	private classLocation			
	private p_currentView			
	private datePickerObject		
	private firstDay				
	private YEAR_SHOW_RANGE			
	private imgSrcSwitchMonthView	
	private imgSrcSwitchWeekView	
	private imgSrcSwitchDayView		
	private imgSrcGoToday			
	private imgSrcGoToDate			
	private dayMenuitems			
	private datePickerUrl			
	private todaysDateOnDisplay		
	
	'Konstruktor => set the default values
	private sub Class_Initialize()
		defaultView					= CALENDARVIEW_WEEK
		defaultStylesheet			= true
		dateToday					= date()
		classLocation				= "/gab_Library/class_calendar/"
		p_currentView				= empty
		selectedDate				= empty
		set datePickerObject		= new datePicker
		firstDay					= empty
		YEAR_SHOW_RANGE				= 10
		imgSrcSwitchWeekView		= classLocation & "icons/icon_weekview_0.gif"
		imgSrcSwitchMonthView		= classLocation & "icons/icon_monthview_0.gif"
		imgSrcGoToday				= classLocation & "icons/icon_goToday_0.gif"
		imgSrcSwitchDayView			= classLocation & "icons/icon_dayview_0.gif"
		imgSrcGoToDate				= "/gab_Library/class_datePicker/icons/icon_0.gif"
		set dayMenuitems			= Server.createObject("Scripting.Dictionary")
		datePickerUrl				= "/gab_Library/class_datePicker/index.asp"
		enabledCalendarViews		= CALENDARVIEW_DAY or CALENDARVIEW_WEEK or CALENDARVIEW_MONTH
		showClock					= true
		runningClock				= true
		todaysDateOnDisplay			= false
		timeline					= false
		enableWeekends				= true
	end sub
	
	public property get currentView() ''[calendarday-enum] Returns the currentView
		currentView = p_currentView
	end property
	
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
	' checkClock 
	'***********************************************************************************************************
	private sub checkClock()
		if runningClock and showClock and todaysDateOnDisplay and currentView <> CALENDARVIEW_MONTH then
			str.writeln("<script language=JavaScript>")
			str.writeln("	updateClock();")
			str.writeln("</script>")
		end if
	end sub
	
	'***********************************************************************************************************
	'' @SDESCRIPTION:	Adds a menuitem to the daymenu
	'' @PARAM:			- dayMenuitemObject [dayMenuitemObject]
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
	'* initJavascript
	'***********************************************************************************************************
	private sub initJavascript()
		str.writeln("<script language=JavaScript src=""" & classLocation & "javascript.js""></script>")
	end sub
	
	'***********************************************************************************************************
	'* printError 
	'***********************************************************************************************************
	private sub printError(msg)
		response.write("<div class=nosuccess>" & msg & "</div>")
		response.end
	end sub
	
	'***********************************************************************************************************
	'* initValues 
	'***********************************************************************************************************
	sub initValues()
		
		'init current-calenderView
		p_currentView = request.queryString("calendarView")
		if p_currentView = empty then
			sessionCalendarView = session("gabLib_calendar_calendarView")
			if sessionCalendarView <> "" then
				p_currentView = sessionCalendarView
			else
				p_currentView = defaultView
			end if
		else
			p_currentView = cint(p_currentView)
		end if
		
		'we have to check if the wanted view is allowed. if not we set defaultview
		if not enabledCalendarViews and currentView then
			p_currentView = defaultView
		end if
		
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
		        firstDay = dateAdd("d", -adjust, selectedDate)
			case CALENDARVIEW_MONTH
				if weekday(dateadd("d", ((day(selectedDate) - 1) * -1), selectedDate), 2) = 1 then
					firstDay = dateadd("d", ((day(selectedDate) - 1 + 7) * -1), selectedDate)
				else
					firstDay = dateadd("d", ((day(selectedDate)-1 + weekday(dateadd("d", ((day(selectedDate))* -1), selectedDate), 2)) * -1), selectedDate)
				end if
		end select
		
		'we store the view and the selected-date in a session,
		'so its easy to request the calendar without any params and get the last view
		session("gabLib_calendar_calendarView") = currentView
		session("gabLib_calendar_selectedDate") = selectedDate
	end sub
	
	'***********************************************************************************************************
	'* printHeader
	'***********************************************************************************************************
	private sub printHeader()
		with str
			.writeln("<form name=calendarForm style=""display:inline;"">")
			.writeln("<div class=calendarHeadline>")
			.writeln("<span class=notForPrint>")
			
			drawGoToDateButton()
			drawGoToTodayButton()
			
			'switch-view buttons
			if enabledCalendarViews and CALENDARVIEW_DAY then
				call drawSitchViewButton("switchViewDay", imgSrcSwitchDayView, CALENDARVIEW_DAY, LANG_SIWTCH_DAY)
			end if
			if enabledCalendarViews and CALENDARVIEW_WEEK then
				call drawSitchViewButton("switchViewWeek", imgSrcSwitchWeekView, CALENDARVIEW_WEEK, LANG_SIWTCH_WEEK)
			end if
			if enabledCalendarViews and CALENDARVIEW_MONTH then
				call drawSitchViewButton("switchViewMonth", imgSrcSwitchMonthView, CALENDARVIEW_MONTH, LANG_SIWTCH_MONTH)
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
			.writeln("<div id=dayMenu onmousemove=""document.onclick = hideDayMenu;"">")
			.writeln("	<input type=Hidden value="""" name=dayMenuClickedDate>")
			.writeln("	<div id=dayMenuItemHeadline class=""dayMenuItem cursorCommon""></div>")
			
			if dayMenuitems.count > 0 then
				for each menuitem in dayMenuitems.items
					menuitem.draw()
				next
			end if
			
			'Go today item
			set item = new dayMenuitem
			item.caption = "<img src=""" & imgSrcGoToDay & """ width=15 height=15 border=0 align=absmiddle>&nbsp;" & LANG_GOTOTODAY
			item.onClick = "goToUrl('" & getScriptUrl(currentView, datePickerObject.formatDate(dateToday)) & "');"
			item.toolTip = LANG_GOTOTODAYHELP
			item.draw()
			set item = nothing
			
			'jump to date item
			set item = new dayMenuitem
			item.caption = "<img src=""" & imgSrcGoToDate & """ width=15 height=15 border=0 align=absmiddle>&nbsp;" & LANG_GOTODATE
			item.onClick = "goToDate('" &  getScriptUrl(currentView, "dateToChange") & "', '" & selectedDate & "', '" & datePickerUrl & "?JSTarget=calendarForm.dummyGoToDateField')"
			item.toolTip = LANG_GOTODATEHELP
			item.draw()
			set item = nothing
			
			'day-view item
			if enabledCalendarViews and CALENDARVIEW_DAY then
				set item = new dayMenuitem
				item.caption = "<img src=""" & imgSrcSwitchDayView & """ width=15 height=15 border=0 align=absmiddle>&nbsp;" & LANG_SIWTCH_DAY
				item.onClick = "dayMenuGoToUrl('" & getScriptUrl(CALENDARVIEW_DAY, "dateToChange") & "')"
				item.toolTip = LANG_SIWTCH_DAY
				item.disabled = (currentView = CALENDARVIEW_DAY)
				item.draw()
				set item = nothing
			end if
			
			'week-view item
			if enabledCalendarViews and CALENDARVIEW_WEEK then
				set item = new dayMenuitem
				item.caption = "<img src=""" & imgSrcSwitchWeekView & """ width=15 height=15 border=0 align=absmiddle>&nbsp;" & LANG_SIWTCH_WEEK
				item.onClick = "dayMenuGoToUrl('" & getScriptUrl(CALENDARVIEW_WEEK, "dateToChange") & "')"
				item.toolTip = LANG_SIWTCH_WEEK
				item.disabled = (currentView = CALENDARVIEW_WEEK)
				item.draw()
				set item = nothing
			end if
			
			'Month-view item
			if enabledCalendarViews and CALENDARVIEW_MONTH then
				set item = new dayMenuitem
				item.caption = "<img src=""" & imgSrcSwitchMonthView & """ width=15 height=15 border=0 align=absmiddle>&nbsp;" & LANG_SIWTCH_MONTH
				item.onClick = "dayMenuGoToUrl('" & getScriptUrl(CALENDARVIEW_MONTH, "dateToChange") & "')"
				item.toolTip = LANG_SIWTCH_MONTH
				item.disabled = (currentView = CALENDARVIEW_MONTH)
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
		str.writeln("<div class=footline></div>")
		str.writeln("</form>")
	end sub
	
	'***********************************************************************************************************
	'* drawGoToTodayButton 
	'***********************************************************************************************************
	private sub drawGoToTodayButton()
		onClick = "goToUrl('" & getScriptUrl(currentView, datePickerObject.formatDate(dateToday)) & "');"
		str.writeln("<button class=icon title=""" & LANG_GOTOTODAYHELP & """ onclick=""" & onClick & """>" &_
					"<img src=""" & imgSrcGoToDay & """ width=15 height=15 border=0 align=absmiddle></button>")
	end sub
	
	'***********************************************************************************************************
	'* drawGoToDateButton 
	'***********************************************************************************************************
	private sub drawGoToDateButton()
		response.write("<input type=Hidden value=0 name=dummyGoToDateField>")
		onClick = "goToDate('" &  getScriptUrl(currentView, "dateToChange") & "','" & selectedDate & "', '" & datePickerUrl & "?JSTarget=calendarForm.dummyGoToDateField');"
		response.write("<button tabindex=""-1"" class=icon onClick=""" & onClick & """ title=""" & LANG_GOTODATEHELP & """>" &_
							"<img src=""" & imgSrcGoToDate & """ width=15 height=15 border=0>" &_
						"</button>")
	end sub
	
	'***********************************************************************************************************
	'* drawYearNavigation 
	'***********************************************************************************************************
	sub drawYearNavigation()
		currentYear = year(selectedDate)
		response.write("<span class=navi>")
		response.write("<button class=""navButton notForPrint"" name=yearPrev onclick=""goToUrl('" & getScriptUrl(currentView, dateAdd("yyyy", -1, selectedDate)) & "')"" title=""" & LANG_PREV_YEAR & """>&lt;</button>")
		set dd = new createDropdown
		with dd
			for i = currentYear - YEAR_SHOW_RANGE to currentYear + YEAR_SHOW_RANGE
				if i > currentYear - YEAR_SHOW_RANGE then
					.sqlQuery = .sqlQuery & ":"
					.pk = .pk & ":"
				end if
				.sqlQuery = .sqlQuery & i
				.pk = .pk & dateAdd("yyyy", i - year(selectedDate), selectedDate)
			next
			.idToMatch = selectedDate
			.onAttribute = "onChange=""goToUrl('" & getScriptUrl(currentView, empty) & "selectedDate=' + this.value);"""
			.draw()
		end with
		
		response.write("<button class=""navButton notForPrint"" name=yearNext onclick=""goToUrl('" & getScriptUrl(currentView, dateAdd("yyyy", +1, selectedDate)) & "')"" title=""" & LANG_NEXT_YEAR & """>&gt;</button>")
		response.write("</span>")
	end sub
	
	'***********************************************************************************************************
	'* drawMonthNavigation
	'***********************************************************************************************************
	private sub drawMonthNavigation()
		response.write("<span class=navi>")
		response.write("<button class=""navButton notForPrint"" name=monthPrev onclick=""goToUrl('" & getScriptUrl(currentView, dateAdd("m", -1, selectedDate)) & "')"" title=""" & LANG_PREV_MONTH & """>&lt;</button>")
		
		set dd = new createDropdown
		with dd
			for i = 1 to 12
				if i > 1 then
					.sqlQuery = .sqlQuery & ":"
					.pk = .pk & ":"
				end if
				.sqlQuery = .sqlQuery & left(MONTH_NAMES(i), 3)
				.pk = .pk & dateAdd("m", i - month(selectedDate), selectedDate)
			next
			.name = "navMonth"
			.idToMatch = selectedDate
			.forceArray = true
			.onAttribute = "onChange=""goToUrl('" & getScriptUrl(currentView, empty) & "selectedDate=' + this.value);"""
			.draw()
		end with
		
		response.write("<button class=""navButton notForPrint"" name=monthNext onclick=""goToUrl('" & getScriptUrl(currentView, dateAdd("m", +1, selectedDate)) & "')"" title=""" & LANG_NEXT_MONTH & """>&gt;</button>")
		response.write("</span>")
	end sub
	
	'***********************************************************************************************************
	'* drawDayNavigation 
	'***********************************************************************************************************
	sub drawDayNavigation()
		currentMonth = month(selectedDate)
		
		response.write("<span class=navi>")
		response.write("<button class=""navButton notForPrint"" name=dayPrev onclick=""goToUrl('" & getScriptUrl(currentView, dateAdd("d", -1, selectedDate)) & "')"" title=""" & LANG_PREV_DAY & """>&lt;</button>")
		set dd = new createDropdown
		with dd
			.name = "navDay"
			for i = 1 to 31
				currentDay = dateAdd("d", i - 1, "01." & month(selectedDate) & "." & year(selectedDate))
				if (enableWeekends or (not enableWeekends and (not weekday(currentDay) = 7 and not weekday(currentDay) = 1))) then
					if month(currentDay) = currentMonth then
						if not i = 1 then
							.sqlQuery = .sqlQuery & ":"
							.pk = .pk & ":"
						end if
						
						.sqlQuery = .sqlQuery & i
						.pk = .pk & currentDay
					else
						exit for
					end if
				end if
			next
			.onAttribute = "onChange=""goToUrl('" & getScriptUrl(currentView, empty) & "selectedDate=' + this.value);"""
			.idToMatch = selectedDate
			.draw()
		end with
		response.write("<button class=""navButton notForPrint"" name=dayNext onclick=""goToUrl('" & getScriptUrl(currentView, dateAdd("d", +1, selectedDate)) & "')"" title=""" & LANG_NEXT_DAY & """>&gt;</button>")
		response.write("</span>")
	end sub
	
	'***********************************************************************************************************
	'* drawWeekNavigation 
	'***********************************************************************************************************
	sub drawWeekNavigation()
		with response
			.write("<span class=navi>")
			.write("<button class=""navButton notForPrint"" name=weekPrev onclick=""goToUrl('" & getScriptUrl(currentView, dateAdd("ww", -1, selectedDate)) & "')"" title=""" & LANG_PREV_WEEK & """>&lt;</button>")
			
			sqlQuery = empty
			pk = empty
			weekNumber = lib.weekOfTheYear(selectedDate)
			upperBound = lib.weekOfTheYear(dateValue("31.12." & year(selectedDate)))
			'maybe the weeknumber is already the first in the next year
			if upperBound = 1 then
				upperBound = lib.weekOfTheYear(dateadd("ww", - 1, dateValue("31.12." & year(selectedDate))))
			end if
			for i = 1 to upperBound
				if not i = 1 then
					sqlQuery = sqlQuery & ":"
					pk = pk & ":"
				end if
				
				mondayOfWeek = dateAdd("ww", i - weekNumber, selectedDate)
				pk = pk & mondayOfWeek
				sqlQuery = sqlQuery & LANG_WEEK & i
			next
			
			set dd = new createDropdown
			with dd
				.name = "navDropdown"
				.sqlQuery = sqlQuery
				.pk = pk
				.idToMatch = selectedDate
				.forceArray = true
				.onAttribute = "onChange=""goToUrl('" & getScriptUrl(currentView, empty) & "selectedDate=' + this.value);"""
				.draw()
			end with
			
			.write("<button class=""navButton notForPrint"" name=weekNext onclick=""goToUrl('" & getScriptUrl(currentView, dateAdd("ww", +1, selectedDate)) & "')"" title=""" & LANG_NEXT_WEEK & """>&gt;</button>")
			.write("</span>")
		end with
	end sub
	
	'***********************************************************************************************************
	'* drawSitchViewButton 
	'***********************************************************************************************************
	private sub drawSitchViewButton(name, imgSrc, enumView, toolTip)
		str.writeln("<button name=" & name & " class=icon " &_
						"onclick=""goToUrl('" & getScriptUrl(enumView, selectedDate) & "');"">" &_
						"<img src=""" & imgSrc & """ width=15 height=15 border=0 title=""" & toolTip & """></button>")
	end sub
	
	'***********************************************************************************************************
	'* drawSitchViewButton 
	'***********************************************************************************************************
	private sub printCalendar()
		str.writeln("<table cellpadding=0 cellspacing=0 border=0 class=calendar align=center>")
		
		select case currentView
			case CALENDARVIEW_DAY
				printDayWeekView()
			case CALENDARVIEW_WEEK
				printDayWeekView()
			case CALENDARVIEW_MONTH
				printMonthView()
		end select
		
		str.writeln("</table>")
	end sub
	
	'***********************************************************************************************************
	'* printMonthView 
	'***********************************************************************************************************
	private sub printMonthView()
		str.writeln("<td class=monthCorner></td>")
		for i = 1 to 7
			index = lib.iif(i = 7, 1, i + 1)
			call printDayHeader("<strong>" & WEEKDAY_NAMES(index) & "</strong>", empty, WEEKDAY_NAMES(weekDayNumber), index = 1 or index = 7, empty)
		next
		
		for i = 0 to 5
			firstDayOfNewWeek = dateadd("d", i * 7, firstDay)
			currentWeek = lib.weekOfTheYear(firstDayOfNewWeek)
			
			if enabledCalendarViews and CALENDARVIEW_WEEK then
				weekLink = "<a href=" & getScriptUrl(CALENDARVIEW_WEEK, firstDayOfNewWeek) & ">" & currentWeek & "</a>"
			else
				weekLink = currentWeek
			end if
			
			str.writeln("<tr>")
			str.writeln("<td valign=top class=weekNumber title=""" & currentWeek & " " & LANG_WEEK_OF_YEAR & """>" & weekLink & "</td>")
			for j = 0 to 6
				set currentDay = new calendarDay
				cssClass2 = empty
				cssClass = empty
				currentDay.index = j + i * 7
				currentDay.dat = dateadd("d", j + i * 7, firstDay)
				currentDay.isWeekend = (weekday(currentDay.dat) = 7 or weekday(currentDay.dat) = 1)
				
				if (currentDay.isWeekend and enableWeekends) or not currentDay.isWeekend then 
					currentDay.isBeforeToday = currentDay.dat < dateToday
					currentDay.weekOfTheYear = currentWeek
					formatedCurrentDay = datePickerObject.formatDate(currentDay.dat)
					toolTip = WEEKDAY_NAMES(weekday(currentDay.dat)) & ", " & formatedCurrentDay
					
					displayDate = day(currentDay.dat)
					
					'we check if the day is in the wanted month
					if month(currentDay.dat) <> month(selectedDate) then
						currentDay.isInSelectedMonth = false
						cssClass = " otherMonth"
						cssClass2 = " otherMonthText"
						dis = " disabled"
					else
						cssClass = empty
						'we check if its a weekend day
						if currentDay.isWeekend then
							cssClass = " calendarHoliday"
						end if
						dis = empty
					end if
					
					'if the day is a holiday
					if datePickerObject.isHoliday(currentDay.dat) then
						currentDay.isHoliday = true
						currentDay.holidayName = datePickerObject.holidayName
						cssClass = cssClass & " calendarHoliday"
						cssClass2 = " holidayName"
						if not currentDay.isInSelectedMonth then
							cssClass2 = cssClass2 & " holidayNameOtherMonth"
						end if
						displayDate = currentDay.holidayName
						toolTip = toolTip & vbcrlf & "(" & displayDate & ")"
					end if
					
					'if the currentday is today
					if currentDay.dat = dateToday then
						currentDay.isToday = true
						todaysDateOnDisplay = true
						if currentDay.isHoliday or not currentDay.isInSelectedMonth then
							cssClass = " calendarToday"
						else
							cssClass = cssClass & " calendarToday"
						end if
						toolTip = toolTip & " - " & ucase(LANG_TODAY)
					end if
					
					onClick = "showDayMenu('" & getScriptUrl(CALENDARVIEW_DAY, currentDay.dat) & "', '" & formatedCurrentDay & "', '" & WEEKDAY_NAMES(weekday(currentDay.dat)) & "')"
					str.writeln("<td class=""monthDay" & cssClass & """ valign=top>")
					str.writeln("<div style=""width:100%;"" onmouseover=""menuHoverIn(this, true)"" onmouseout=""menuHoverOut(this, true)"" class=""monthDayText" & cssClass2 & """ title=""" & toolTip & """ onclick=""" & onClick & """>" & displayDate & "</div>")
					onDayCreated(currentDay)
					str.writeln("</td>")
				end if
				
				set currentDay = nothing
			next
			str.writeln("</tr>")
		next
	end sub
	
	'***********************************************************************************************************
	'* printDayHeader 
	'***********************************************************************************************************
	private sub printDayHeader(title, url, toolTip, isWeekend, dat)
		if isWeekend  and enableWeekends or not isWeekend then
			if isWeekend then
				cssClass = " weekdayNameWeekend"
			else
				cssClass = empty
			end if
			
			if not url = empty then
				onClick = " onClick=""showDayMenu('" & getScriptUrl(CALENDARVIEW_DAY, dateToday) & "', '" & dat & "', '" & WEEKDAY_NAMES(weekday(dat)) & "')"""
				onElse = " onmouseover=""menuHoverIn(this)"" onmouseout=""menuHoverOut(this)"""
				cssClass = cssClass & " handCursor"
			end if
			
			str.writeln("<td" & onClick & " class=""weekdayName " & lib.iif(enableWeekends, "seventhOfPage", "fifthOfPage") & " whiteRightBorder" & cssClass & """ title=""" & toolTip & """><div" & onElse & ">" & title & "</div></td>")
		end if
	end sub
	
	'***********************************************************************************************************
	'* printWeekView 
	'***********************************************************************************************************
	private sub printDayWeekView()
		oppositeView = CALENDARVIEW_DAY + CALENDARVIEW_WEEK - currentView
		if currentView = CALENDARVIEW_DAY then
			upperBound = 1
		else
			upperBound = 7
		end if
		
		str.writeln("<tr>")
		
		'print out weekdaynames
		for i = 1 to upperBound
			index = lib.iif(i = 7, 1, i + 1)
			newDay = datePickerObject.formatDate(dateAdd("d", i - 1, firstDay))
			weekDayNumber = weekday(newDay)
			
			call printDayHeader("<strong>" & WEEKDAY_NAMES(weekDayNumber) & "</strong><br>" & newDay, _
								getScriptUrl(oppositeView, newDay), _
								WEEKDAY_NAMES(weekDayNumber), _
								weekDayNumber = 1 or weekDayNumber = 7, newDay)
		next
		str.writeln("</tr>")
		
		'print the timeline
		str.writeln("<tr class=weekViewDay>")
		for i = 1 to upperBound
			cssClass = empty
			newDay = dateAdd("d", i - 1, firstDay)
			
			set currentDay = new calendarDay
			currentDay.index = i - 1
			currentDay.dat = newDay
			currentDay.isBeforeToday = currentDay.dat < dateToday
			weekDayNumber = weekday(currentDay.dat)
			currentDay.isWeekend = weekDayNumber = 1 or weekDayNumber = 7
			
			if (currentDay.isWeekend and enableWeekends) or not currentDay.isWeekend then 
				currentDay.isHoliday = datePickerObject.isHoliday(currentDay.dat)
				
				currentDay.isToday = currentDay.dat = dateToday
				if currentDay.isToday then
					todaysDateOnDisplay = true
					cssClass = " calendarToday"
				end if
				
				if currentDay.isWeekend or currentDay.isHoliday then
					cssClass = " calendarHoliday"
				end if
				
				str.writeln("<td class=""calendarDay" & cssClass & """ valign=top>")
				if currentDay.isHoliday then
					currentDay.holidayName = datePickerObject.holidayName
					response.write("<div class=holidayName>" & currentDay.holidayName & "</div>")
				end if
				if currentDay.isToday and showClock then
					currentTime = "init clock ..."
					if not runningClock then
						currentHour = right("00" & hour(now()), 2)
						currentMinute = right("00" & minute(now()), 2)
						currentTime = currentHour & ":" & currentMinute
					end if
					response.write("<div class=todaysTime id=todayClock>" & currentTime & "</div>")
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
	
	'******************************************************************************************
	'* showTimeLine 
	'******************************************************************************************
	sub showTimeLine()
		if timeline and currentView = CALENDARVIEW_DAY then
			response.write("<div class=calendarTimeline>")
			for i = 0 to 5
				response.write("<span class=""calendarTimelineHour"">" & right("00" & i * 4, 2) & ":00</span>")
			next
			response.write("</div>")
		end if
	end sub
	
	'***********************************************************************************************************
	'* getScriptUrl 
	'***********************************************************************************************************
	private function getScriptUrl(calendarView, dateSelected)
		queryString = lib.getAllFromQueryStringBut("calendarView,selectedDate")
		scriptName = request.serverVariables("SCRIPT_NAME") & "?"
		
		if not queryString = empty then
			scriptName = scriptName & queryString & "&"
		end if
		
		if not calendarView = empty then
			scriptName = scriptName & "calendarView=" & calendarView & "&"
		end if
		
		if not dateSelected = empty then
			 scriptName = scriptName & "selectedDate=" & dateSelected & "&"
		end if
		getScriptUrl = scriptName
	end function
	
	'***********************************************************************************************************
	'* initStyles
	'***********************************************************************************************************
	private sub initStyles()
		if defaultStylesheet then
			str.writeln("<link rel=stylesheet type=text/css href='" & classLocation & "standard.css'>")
		end if
	end sub

end class
%>