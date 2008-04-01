<!--#include virtual="/gab_Library/class_page/generatepage.asp"-->
<!--#include virtual="/gab_Library/class_calendar/calendar.asp"-->
<%
dim cal

set page = new generatePage
with page
	.DBConnection 	= true
	.title 			= "Calendar Demo"
	.debugMode		= false
	.loginRequired	= false
	.devWarning		= false
	.draw()
end with
set page = nothing

'******************************************************************************************
'* main - we start . DONT FORGET TO KILL ALL OBJECTS!! 
'******************************************************************************************
sub main()
	set cal = new calendar
	with cal
		'.enableWeekends = false
		'.defaultView = CALENDARVIEW_WEEK
		'.enabledCalendarViews = CALENDARVIEW_DAY or CALENDARVIEW_WEEK
		
		set item = new dayMenuitem
		item.caption = "New Event"
		item.toolTip = "Click here to add a new Event"
		item.onClick = "alert('a new event')"
		.addDayMenuitem(item)
		set item = nothing
		
		set item = new dayMenuitem
		item.caption = "Something disabled"
		item.toolTip = "Click here to add an absence"
		item.disabled = true
		.addDayMenuitem(item)
		set item = nothing
		
		.addDayMenuSeperator()
		
		.draw()
	end with
	
	set cal = nothing
end sub

'******************************************************************************************
'* onDayCreated 
'******************************************************************************************
sub onDayCreated(ByVal currentDate)
	with str
		if currentDate.isHoliday then
			.writeln("I am a holiday")
		elseif currentDate.isToday then
			.writeln(currentDate.dat & "<br>is today. Cool!")
		elseif not currentDate.isInSelectedMonth then
			.writeln("<div disabled>not in this month</div>")
		end if
	end with
end sub
%>