<%
'**************************************************************************************************************

'' @CLASSTITLE:		calendarDay
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		05.08.2004
'' @CDESCRIPTION:	Represents a calendarDay for the calendar.
'' @VERSION:		1.0

'**************************************************************************************************************

class calendarDay

	public dat						''[string] the date of the calendarday
	public index					''[int] index of the day in the calendar. example: weekview has 7 days so
									''the first day has index 0 and last day index 6.
	public isHoliday				''[bool] is it a holiday
	public holidayName				''[string] helds the name of the holiday. (if isHoliday)
	public isInSelectedMonth		''[bool] is the day in the selectedMonth? important for MONTH_VIEW
	public isWeekend				''[bool] is the day a saturday or sunday
	public weekOfTheYear			''[int] determines in what week of the year the day belongs to
	public isToday					''[bool] determines if the day is today
	public isBeforeToday			''[bool] determines if the day is before todays day
	
	'Konstruktor => set the default values
	private sub Class_Initialize()
		dat							= empty
		index						= 0
		isHoliday					= false
		isInSelectedMonth			= true
		isWeekend					= false
		holidayName					= empty
		isToday						= false
		isBeforeToday				= false
	end sub

end class
%>