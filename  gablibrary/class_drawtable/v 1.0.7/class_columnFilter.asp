<%
'**************************************************************************************************************

'' @CLASSTITLE:		columnFilter
'' @CREATOR:		Michal Gabrukiewicz - gabru@gmx.at
'' @CREATEDON:		25.09.2003
'' @CDESCRIPTION:	Needed for drawtable-class. It will create a filter for a table-column
'' @VERSION:		0.1

'**************************************************************************************************************
class columnFilter

	public fieldName			''The name of the column the filter should be placed to.
	public showDropdownSQL		''If you want a dropdown then provide a sql-query if you want a common text-field leave it blank.
	public primaryKey			''What field should be written in the value-attribute of the option-tag.
	public displayField			''The name of the field should be displayed.
	public defaultSelect		''You can set a default value to set the filter.
	public fieldToMatch			''If the field you want to match differs from the fieldName then write it here.
	public description			''Leave it blank to use the default description.
	public tdAttributes			''You have access to the TD the filter will be placed in. e.g. align=center
	public commaStyle			''You can choose how the comma looks like
	
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
	end sub

end class
%>