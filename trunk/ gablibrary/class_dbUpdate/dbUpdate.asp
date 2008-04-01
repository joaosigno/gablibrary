<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		DBUpdate
'' @CREATOR:		Michal Gabrukieiwcz - gabru @ grafix.at
'' @CREATEDON:		07.08.2003
'' @CDESCRIPTION:	This is needed to make Database updates easier. Automatically generates sql-query
''					including all neccessary functions for UPDATE, INSERT and DELETE. You just need to
''					name every form-field you want to use in your DBUpdate the same as it is named in the
''					table you want to update. The class finds every match and creates the sql. You just
''					need object.update, object.delete or object.insert. Thats all!
'' @VERSION:		1.0

'**************************************************************************************************************
class DBUpdate

	private RS
	private fields
	private values
	private connOBJ
	private fieldFunctions
	
	public table				''[string ]We need to know which table should be updated
	public pkName				''[string] Name of the column in the specified table which is the primary-key-column
	public pk					''[int] ID of the record to be updated. just needed for deleting and updating
	public verifyTxt			''[bool] Check text for <BR> and ' etc. VerfiyTextForDB-function from lib will be exectued on every field before updating.
	public debuging				''[bool] turns debugging on/off
	public ignorePK				''[bool] If the PK is not an AUTO-Value then you can use this property to allow updating the PK field
	public acceptNulls			''[bool] writes null into columns which allow nulls and the input is empty. default = falst
	
	'************************************************************************************************************
	'* Constructor 
	'************************************************************************************************************
	public sub class_Initialize()
		fields				= empty
		values				= empty
		table				= empty
		pk					= 0
		pkName				= empty
		verifyTxt			= true
		set connOBJ			= lib.databaseConnection
		debuging			= false
		ignorePK			= false
		set fieldFunctions	= server.createObject("scripting.dictionary")
		acceptNulls			= false
	end sub
	
	'************************************************************************************************************
	'* destructor 
	'************************************************************************************************************
	private sub class_terminate()
		set fieldFunctions = nothing
		set RS = nothing
	end sub
	
	'************************************************************************************************************
	''@SDESCRIPTION:	Use this method to specify the connection object.
	''@DESCRIPTION:		So you first connect to your database and then provide the connection to the dbupdate.
	''					Just use this method if your connection-object-name differs "lib.databaseConnection" cause this is set as default.
	''@PARAM:			- obj: Your connection Object
	'************************************************************************************************************
	public function connection(obj)
		if not obj = empty then set connOBJ = obj
	end function
	
	'************************************************************************************************************
	''@SDESCRIPTION:	Executes a specified function on a field.
	''@PARAM:			fieldname: Fieldname on which the function should be executed
	''@PARAM:			functionToExecute: Name of the function to execute.
	'************************************************************************************************************
	public function changeFieldValue(fieldname, functionToExecute)
		if functionToExecute <> "" and fieldname <> "" then
			if not fieldFunctions.exists(uCase(fieldname)) then fieldFunctions.add uCase(fieldname), functionToExecute
		end if
	end function
	
	'***************************************************************************************************************
	' main 
	'***************************************************************************************************************
	private function main(action)
		myFields = empty
		set RS = Server.CreateObject("ADODB.Recordset")
		RS.ActiveConnection	= connOBJ
		RS.source = table
		RS.open()
		
		'we durchforst all fields from the table
		for each ofield in rs.fields
			found = false
			
			'and now we "durchforst" the form fields
			for each field in request.form
				if cstr(ucase(ofield.name)) = cstr(ucase(field)) AND ((not cstr(ucase(ofield.name)) = cstr(ucase(pkName)) AND not ignorePK) OR ignorePK) then
					found = true
				end if
			next
			
			if found then
				myFormField = request.form(ofield.name)
				if verifyTxt then myFormField = verifyFieldForDB(myFormField)
				if fieldFunctions.exists(ucase(ofield.name)) then execute("myFormField = " & fieldFunctions(uCase(oField.name)))
				
				select case action
					case "insert"
						fields = fields & ofield.name & ","
					case "update"
						fields = fields & ofield.name & "="
				end select
				
				'we check if the field accepts nulls (adFldIsNullable) then we dont do anything with it
				if (trim(myFormField) = "" or isNull(myFormField)) and acceptNulls and (oField.attributes and &H20) then
					select case action
						case "insert"
							values = values & "NULL,"
						case "update"
							fields = fields & "NULL,"
					end select
				else
					'we check whether its a text-column or a number-column in the db
					if ofield.type = 200 or ofield.type = 201 or ofield.type = 202 or ofield.type = 203 or ofield.type = 135 then
						'truncate to length of field in db, that we dont get an error if field is not well dimensioned
						myFormField = left(myFormField, ofield.definedsize)
						select case action
							case "insert"
								values = values & "'" & myFormField & "',"
							case "update"
								fields = fields & "'" & myFormField & "',"
						end select
					else
						myFormField = toDot(myFormField)
						select case action
							case "insert"
								values = values & myFormField & ","
							case "update"
								fields = fields & myFormField & ","
						end select
					end if
				end if
				
			end if
		next
	end function
	
	'***************************************************************************************************************
	'* verifyFieldForDB 
	'***************************************************************************************************************
	public function verifyFieldForDB(inputStr)
		if not inputStr = empty then
			tmp = str.SQLSafe(inputStr)
			
			'tmp = replace(tmp,chr(13),vbcrlf)		' RETURNS
			newtmp = replace(tmp, chr(39), "´")
			'tmp = replace(tmp,chr(34),"&rdquo;")	' "
			
			verifyFieldForDB = newtmp
		end if
	end function
	
	'***************************************************************************************************************
	' toDot 
	'***************************************************************************************************************
	private function toDot(num)
		if not isnull(num) then 
			toDot = Replace(num, ",", ".")
		else
			toDot = num
		end if
	end function
	
	'***************************************************************************************************************
	' executeMe 
	'***************************************************************************************************************
	private function executeMe(sql)
		if debuging then str.write(sql)
		on error resume next
		lib.getRecordset(sql)
		executeMe = (err = 0)
		set connOBJ = nothing
	end function
	
	'************************************************************************************************************
	''@SDESCRIPTION:	Insert a new record to the table.
	''@RETURN:			[bool] - was the insert successfull?
	'************************************************************************************************************
	public function insert()
		'first we call the main function
		main("insert")
		
		'all done we insert the record
		if not fields = empty and not values = empty and not table = empty then
			insertSql = "INSERT INTO " & table & " (" & left(fields, (len(fields) - 1)) & ") VALUES (" & left(values, (len(values) - 1)) & ")"
			insert = executeME(insertSql)
		end if
		fields = empty
		values = empty
	end function
	
	'************************************************************************************************************
	''@SDESCRIPTION:	Update the table with the data. Needs a pk-value
	''@RETURN:			[bool] - was the update successfull?
	'************************************************************************************************************
	public function update()
		'first we call the main function
		main("update")
		
		'all done we update the record
		if not fields = empty and not table = empty and not pk = empty and not pkName = empty then
			updateSql = "UPDATE " & table & " SET " & left(fields, (len(fields) - 1)) & " WHERE " & pkName & " = " & pk
			update = executeME(updateSql)
		end if
		fields = empty
		values = empty
	end function
	
	'************************************************************************************************************
	''@SDESCRIPTION:	Delete this record. Needs a pk-value 
	''@RETURN:			[bool] - was the delete successfull?
	'************************************************************************************************************
	public function delete()
		if not pk = 0 and not pkName = empty and not table = empty then
			deleteSql = "DELETE FROM " & table & " WHERE " & pkName & " = " & pk
			delete = executeME(deleteSql)
		end if
	end function

end class
lib.registerClass("DBUpdate")
%>