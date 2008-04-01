<%
'**************************************************************************************************************

'' @CLASSTITLE:		columnRadioButton
'' @CREATOR:		Michal Gabrukiewicz - gabru@gmx.at
'' @CREATEDON:		25.09.2003
'' @CDESCRIPTION:	Needed for drawtable-class. It will create a column with Radiobuttons.
''					Just add more than one Column with the same "fieldName" and the class will know that they
''					shold be grouped together! IMPORTANT: Always write one after another. Never mix it with
''					column-types: RIGHT: id_group,id_group,id_group WRONG: id_group,otherCol,id_group,id_group
'' @VERSION:		0.1

'**************************************************************************************************************
class columnRadioButton

	public fieldName			''Name of the filed in the sql-query. So you see every field must be available in your sql-query!
	public displayString		''The name displayed in the header of the table for this column.
	public value				''Whats the value of the Radiobutton in this column?
	public tdAttributes			''TD-Attributes for every TD in this column. e.g: width=20
	public headerTdAttributes	''TD-Attributes for the TD the sort is in.
	public colType				''Dont touch this!
	public isNumber				''true/false - is this a column with numbers?
	public disabled				''[bool] indicates if the radiobuttons are disabled or not? default = false
	
	'Konstruktor => set the default values
	private sub Class_Initialize()
		fieldName			= empty
		displayString		= empty
		value				= empty
		tdAttributes		= empty
		headerTdAttributes	= empty
		colType				= "radiobutton"
		isNumber			= false
		disabled 			= false
	end sub

end class
%>