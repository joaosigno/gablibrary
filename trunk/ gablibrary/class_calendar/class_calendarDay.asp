<%
'**************************************************************************************************************

'' @CLASSTITLE:		calendarDay
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		05.08.2004
'' @CDESCRIPTION:	Represents a calendarDay for the calendar.
'' @VERSION:		1.2

'**************************************************************************************************************

class CalendarDay

	'private members
	private datesObject
	private today
	private p_dat
	
	'public members
	public index			''[int] index of the day in the calendar. example: weekview has 7 days so
							''the first day has index 0 and last day index 6.
	public attributes		''[string] attributes for the calendar day (e.g. in the month view the attribute for the whole day cell)
	public HLattributes		''[string] head line attributes (used in the month view for the monthDayText Div)
	
	public property let dat(value) ''[date] sets the date
		p_dat = value
	end property
	
	public property get dat ''[date] gets the date
		dat = p_dat
	end property
	
	'***********************************************************************************************************
	'* constructor 
	'***********************************************************************************************************
	public sub class_Initialize()
		today					= date()
		set datesObject			= new dates
		dat						= empty
		attributes				= empty
		HLattributes			= empty
		index					= 0
	end sub
	
	'***********************************************************************************************************
	'* destructor 
	'***********************************************************************************************************
	private sub class_terminate()
		set datesObject = nothing
	end sub
	
	'***********************************************************************************************************
	'' @SDESCRIPTION:	indicates if the date lies on a weekend or not
	'' @RETURN:			[bool] true if its a weekend
	'***********************************************************************************************************
	public function isWeekend()
		isWeekend = datesObject.isWeekend(weekday(dat))
	end function
	
	'***********************************************************************************************************
	'' @SDESCRIPTION:	indicates if the date is Today
	'' @RETURN:			[bool] true if its today
	'***********************************************************************************************************
	public function isToday()
		isToday = (dat = today)
	end function
	
	'***********************************************************************************************************
	'' @SDESCRIPTION:	indicates if the date is before todays date
	'' @RETURN:			[bool] true if it is before today
	'***********************************************************************************************************
	public function isBeforeToday()
		isBeforeToday = (dat < today)
	end function
	
	'***********************************************************************************************************
	'' @SDESCRIPTION:	gets the week of the year the date is in
	'' @RETURN:			[int] weeknumber
	'***********************************************************************************************************
	public function getWeekOfTheYear()
		getWeekOfTheYear = datesObject.weekOfTheYear(dat)
	end function
	
	'***********************************************************************************************************
	'' @SDESCRIPTION:	indicates if the date is a holiday
	'' @RETURN:			[bool] true if its a holiday
	'***********************************************************************************************************
	public function isHoliday()
		isHoliday = datesObject.isHoliday(dat)
	end function
	
	'***********************************************************************************************************
	'' @SDESCRIPTION:	gets the holidayname if it is a holiday. useful in combination with isHoliday()
	'' @RETURN:			[bool] true if its a holiday
	'***********************************************************************************************************
	public function getHolidayName()
		getHolidayName = datesObject.holidayName
	end function
	
	'***********************************************************************************************************
	'' @SDESCRIPTION:	gets the name of the weekday. e.g. Monday
	'' @RETURN:			[string] weekdayname
	'***********************************************************************************************************
	public function getWeekdayName()
		getWeekdayName = WEEKDAY_NAMES(weekday(dat))
	end function

end class
%>