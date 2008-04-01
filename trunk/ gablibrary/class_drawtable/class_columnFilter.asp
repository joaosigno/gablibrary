<%
'**************************************************************************************************************

'' @CLASSTITLE:		columnFilter
'' @CREATOR:		Michal Gabrukiewicz - gabru @ grafix.at
'' @CREATEDON:		25.09.2003
'' @CDESCRIPTION:	Needed for drawtable-class. It will create a filter for a table-column
'' @VERSION:		0.1
'' @FRIENDOF:		DrawTable

'**************************************************************************************************************
class columnFilter

	public fieldName			''[string] The name of the column the filter should be placed to.
	public showDropdownSQL		''[string] If you want a dropdown then provide a sql-query if you want a common text-field leave it blank.
	public dropdownAutosplit	''[bool] indicates if the dropdownSQL should be automatically detected if its an sql or values for the dropdown. default = true
								''if you have ':' values in your SQL then you might turn this off
	public primaryKey			''[string] What field should be written in the value-attribute of the option-tag.
	public displayField			''[string] The name of the field should be displayed.
	public defaultSelect		''[string] You can set a default value to set the filter.
	public fieldToMatch			''[string] If the field you want to match differs from the fieldName then write it here.
	public description			''[string] Leave it blank to use the default description.
	public tdAttributes			''[string] You have access to the TD the filter will be placed in. e.g. align=center
	public commaStyle			''[string] You can choose how the comma looks like
	public inputAttributes		''[string] Attributes for the input field. e.g. readonly, onclick='...', etc.
	
	'Konstruktor => set the default values
	private sub Class_Initialize()
		fieldName		= empty
		showDropdownSQL	= empty
		primaryKey		= empty
		displayField	= empty
		defaultSelect	= empty
		fieldToMatch	= empty
		description		= empty
		tdAttributes	= empty
		commaStyle		= empty
		inputAttributes = empty
		dropdownAutosplit = true
	end sub

end class
%>