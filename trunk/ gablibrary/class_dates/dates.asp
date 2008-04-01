<!--#include virtual="/gab_LibraryConfig/_dates.asp"-->
<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003 - This file is part of GAB_LIBRARY		
'* For license refer to the license.txt in the root    						
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		Dates
'' @CREATOR:		Michael Rebec, Michal Gabrukiewicz
'' @CREATEDON:		2007-01-05 13:58
'' @CDESCRIPTION:	Contains functions for manipulating dates. Normaly it should be named "date" but thanks to
''					ms this word is already in use. So we have choosen the classname "dates" instead.
'' @REQUIRES:		-
'' @POSTFIX:		dat
'' @VERSION:		0.1

'**************************************************************************************************************
class Dates

	private p_holidayName
	
	public property get holidayName ''[string] holds the name of the holiday after executing isHoliday()
		holidayName = p_holidayName
	end property
	
	'**********************************************************************************************************
	''	@SDESCRIPTION: 	Gets an independent sql representation for a datetime string standardised by ISO 8601 
	''	@DESCRIPTION:	If you want to create a sql statement containing a date query in the where clause,
	''					use this function to create a datetime object.
	''					EXAMPLE:
	''					In Sql Server: Declare @timeTable TABLE (name int, starttime datetime)
	''					In VBScript: sql = "SELECT * FROM timeTable WHERE starttime = " & date.toMsSqlDateFormat(cDate('2007-01-01 15:30:00'))
	''					Results in: SELECT * FROM timeTable WHERE starttime = cast('2006-01-01T15:30:00' as datetime)
	'' 					NOTE: only for MS SQL Server
	'' @PARAM:			dat [date]: the date/time you want to cast
	'' @RETURN:			[string] the formatted date string
	'**********************************************************************************************************
	public function toMsSqlDateFormat(dat)
		if dat = empty then exit function
		toMsSqlDateFormat = "cast('" & year(dat) & "-" & lib.custom.forceZeros(month(dat), 2) & "-" & _
							lib.custom.forceZeros(day(dat), 2) & "T" & lib.custom.forceZeros(hour(dat), 2) & ":" & _
							lib.custom.forceZeros(minute(dat), 2) & ":" & lib.custom.forceZeros(second(dat), 2) & "' as datetime)"
	end function
	
	''******************************************************************************************
	'' @SDESCRIPTION: 	Gets the last day of a given date (e.g. 15.01.2006 returns 31
	'' @PARAM:			currentDate [date]: the date
	'' @RETURN:			[int] the last day of the month
	''******************************************************************************************
	public function getLastDayOfTheMonth(currentDate)
		getLastDayOfTheMonth = DatePart("d", dateSerial(Year(currentDate), month(currentDate)+1, 0))
	end function
	
	''******************************************************************************************
	'' @SDESCRIPTION: 	Gets the number of days of a given month (e.g. 31 for january)
	'' @PARAM:			mth [int]: the month you want the days of (e.g. 2 (Feb))
	'' @PARAM:			yr [int]: the year in which the month is placed (e.g. 2006)
	'' @RETURN:			[int] number of days
	''******************************************************************************************
	public function getDaysInMonth(mth, yr)
		nextMonth = lib.iif(mth = 12, 1, mth + 1)
		getDaysInMonth = day(dateadd("d", -1, cdate(yr & "-" & nextMonth & "-01")))
	end function
	
	'***********************************************************************************************************
	'' @SDESCRIPTION:	Returns the week of the year of the given date
	'' @DESCRIPTION:	due to a MS bug the common asp week function returns sometimes 53 when it 
	''					should return 1. so there was a need to code this function to return the right number
	'' @PARAM:			myDate [date]: the date
	'' @RETURN:			[int] week of the year
	'***********************************************************************************************************
	public function weekOfTheYear(myDate)
		weekOfTheYear = datepart("ww", myDate, vbMonday, vbFirstFourDays)
		if weekOfTheYear > 52 then
			if datepart("ww", myDate + 7, vbMonday, vbFirstFourDays) = 2 then weekOfTheYear = 1
		end If
	end function
	
	'***********************************************************************************************************
	'' @SDESCRIPTION:	checks if a given weekday is on a weekend
	'' @PARAM:			aWeekday [int]: weekday number.
	'' @RETURN: 		[bool] true if it lies on weekend
	'***********************************************************************************************************
	public function isWeekend(aWeekday)
		isWeekend = (aWeekday = vbSunday or aWeekday = vbSaturday)
	end function
	
	'***********************************************************************************************************
	'' @SDESCRIPTION:	Says if the given date is a holiday or not
	'' @DESCRIPTION:	If the given date is a holiday the method writes into the holidayName-member-var the
	''					holidayName like its defined in the language.asp
	'' @PARAM:			- theDate [date]: the date you want to check against holiday
	'' @RETURN:			[bool] returns true if the date is a holiday.
	'***********************************************************************************************************
	public function isHoliday(theDate)
		p_holidayName = empty
		
		dateString = Day(theDate) & "." & Month(theDate)
		select case dateString
			case "1.1":		p_holidayName = HOLIDAY_NAME(1)
			case "6.1":		p_holidayName = HOLIDAY_NAME(2)
			case "1.5":		p_holidayName = HOLIDAY_NAME(3)
			case "15.8":	p_holidayName = HOLIDAY_NAME(4)
			case "26.10":	p_holidayName = HOLIDAY_NAME(5)
			case "1.11":	p_holidayName = HOLIDAY_NAME(6)
			case "8.12":	p_holidayName = HOLIDAY_NAME(7)
			case "25.12":	p_holidayName = HOLIDAY_NAME(8)
			case "26.12":	p_holidayName = HOLIDAY_NAME(9)
		case else:
			select case (CDate(theDate) - calculateEaster(theDate))
				case  0:	p_holidayName = HOLIDAY_NAME(10)
				case  1:	p_holidayName = HOLIDAY_NAME(11)
				case 39:	p_holidayName = HOLIDAY_NAME(12)
				case 49: 	p_holidayName = HOLIDAY_NAME(15)
				case 50:	p_holidayName = HOLIDAY_NAME(13)
				case 60:	p_holidayName = HOLIDAY_NAME(14)
			end select
		end select
		
		isHoliday = (p_holidayName <> empty)
	end function
	
	'***********************************************************************************************************
	'* calculateEaster
	'***********************************************************************************************************
	private function calculateEaster(theDate)
		mYear = Year(theDate)
		mA = mYear Mod 19
		mD = (19 * mA + 24) Mod 30
		mE_ = (2 * (mYear Mod 4) + 4 * (mYear Mod 7) + 6 * mD + 5) Mod 7
		easterSunday = CDate(DateSerial(mYear, 3, 22 + mD + mE_))
		if Month(easterSunday) = 4 then
			if Day(easterSunday) = 26 Or (Day(easterSunday) = 25 And mE_ = 6 And mA > 10) then
				easterSunday = easterSunday - 7
			end if
		end if
		
		calculateEaster = easterSunday
	end function

end class
lib.registerClass("Dates")
%>