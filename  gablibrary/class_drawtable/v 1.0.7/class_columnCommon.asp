<%
'**************************************************************************************************************

'' @CLASSTITLE:		ColumnCommon
'' @CREATOR:		Michal Gabrukiewicz - gabru@gmx.at
'' @CREATEDON:		25.09.2003
'' @CDESCRIPTION:	Needed for drawtable-class. It will create a column
'' @VERSION:		0.1

'**************************************************************************************************************
class columnCommon

	public fieldName			''Name of the filed in the sql-query. So you see every field must be available in your sql-query!
	public displayString		''The name displayed in the header of the table for this column.
	public displayFunction		''Name of the function you want to execute on the column.
								''e.g. if you choose "checkme" as display-function you will need to write your own function
								''named "checkme" including one parameter. This parameter is always the current value of the field.
	public tdAttributes			''TD-Attributes for every TD in this column. e.g: width=20
	public headerTdAttributes	''TD-Attributes for the TD the sort is in.
	public colType				''Dont touch this!
	public isNumber				''[bool] - is this a column with numbers?
	public disableLink			''[bool] - if true then column will be without link (not clickable). addurl/addurljs members of
								''drawtable will be ignored. you can make your own links then
	
	'Konstruktor => set the default values
	private sub Class_Initialize()
		fieldName			= empty
		displayString		= empty
		displayFunction		= empty
		tdAttributes		= empty
		headerTdAttributes	= empty
		colType				= "common"
		isNumber			= false
		disableLink			= false
	end sub

end class
%>