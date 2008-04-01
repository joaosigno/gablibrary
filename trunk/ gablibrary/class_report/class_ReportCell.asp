<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003 - This file is part of GAB_LIBRARY		
'* For license refer to the license.txt in the root    						
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		ReportCell
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		2006-10-20 18:43
'' @CDESCRIPTION:	Represents a cell for the report
'' @FRIENDOF:		Report
'' @VERSION:		0.1

'**************************************************************************************************************
class ReportCell

	'private members
	private p_value
	
	'public members
	public report				''[Report] report which the cell belongs to
	public row					''[ReportParameter] the row to which the cell belongs to
	public col					''[ReportParameter] the column to which the cell belongs to
	public attributes			''[string] attributes of the cell
	public param				''[ReportParameter] the parameter which was used to get the value
	
	public default property get value ''[float] gets the value which is held by the cell
		value = p_value
	end property
	
	public property let value(val) ''[float] sets the value. when its not a number then it will be set to 0. undefined is accepted as well.
		v = val & ""
		p_value = 0
		if val = report.undefinedSign then
			p_value = val
		elseif isNumeric(v) then
			p_value = cdbl(v)
		end if
	end property
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		p_value = 0
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	draws the cell
	'**********************************************************************************************************
	public sub draw(isLastCol, isLastRow)
		report.write("<td class=""" & getCssClass(isLastCol) & """ nowrap " & attributes & ">")
		report.write(formatValue())
		report.write("</td>")
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	indicates if the stored value is undefined or not
	'' @RETURN:			[bool] true if undefined (e.g. division by zero)
	'**********************************************************************************************************
	public function valueIsUndefined()
		valueIsUndefined = (value = report.undefinedSign)
	end function
	
	'**********************************************************************************************************
	'* formatValue 
	'**********************************************************************************************************
	private function formatValue()
		if valueIsUndefined() then
			formatValue = value
		else
			numAfterDigits = report.digitsAfterDecimal
			if not isEmpty(param.digitsAfterDecimal) then numAfterDigits = param.digitsAfterDecimal
			formatValue = formatNumber(value, numAfterDigits)
		end if
	end function
	
	'**********************************************************************************************************
	'* getCssClass 
	'**********************************************************************************************************
	private function getCssClass(isLastColumn)
		getCssClass = "data"
		if isLastColumn then getCssClass = getCssClass & " inRowLast"
		if valueIsUndefined() then
			getCssClass = getCssClass & " undefined"
		elseif value = 0 then
			getCssClass = getCssClass & " zero"
		elseif value <= 0 then
			getCssClass = getCssClass & " negative"
		end if
	end function

end class
%>