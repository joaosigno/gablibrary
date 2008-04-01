<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003 - This file is part of GAB_LIBRARY		
'* For license refer to the license.txt in the root    						
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		ReportParameter
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		2006-10-02 21:23
'' @CDESCRIPTION:	Represents a parameter of the report
'' @REQUIRES:		-
'' @VERSION:		0.1

'**************************************************************************************************************
class ReportParameter

	'public members
	public value				''[string] value of the field.
								''The name of a column which can be found in the recordset generated by the sql of the Report.
								''if value starts with "=" then a formula will be applied to the field. example: ={a}+{b}
								''means that the parameter with named "a" will be added to the parameter name "b".
								''Formula usage:
								''{} = value of the parameter with the name between the brackets
								''[] = these brackets mean that the value between is a function. supported functions yet (SUM, AVG)
	public caption				''[string] caption which will be displayed for the field
	public name					''[string] name of the parameter which is used within formulas. if you need
								''to refer to this parameter within a formula then you must provide a name
	public border				''[string] which border should be applied to the parameter. left, right, top, bottom (and all combinations)
								''example: "top bottom left right" would draw a border around the whole parameter
	public index				''[int] index of the parameter. starting with 0
	public priority				''[bool] indicates if this parameter has priority when it comes to a colision between section and parameter. default = false
	public digitsAfterDecimal	''[int] number of the digits when formating the number. by default its
								''taken from the report property
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		name = ""
		index = 0
		border = ""
		priority = false
		digitsAfterDecimal = empty
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	indicates if the value holds a formula or not. if empty then it does not hold one ;)
	'' @RETURN:			[bool] true if it holds a formula
	'**********************************************************************************************************
	public function hasFormula()
		hasFormula = str.startsWith(value, "=")
	end function

end class
%>