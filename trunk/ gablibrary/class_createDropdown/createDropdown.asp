<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		CreateDropdown
'' @CREATOR:		Michal Gabrukieiwcz - gabru @ grafix.at
'' @CREATEDON:		09.07.2003
'' @CDESCRIPTION:	Easily create a dropdown. Weather common or multiple. Just needed a sql-query
''					or even an array. This class will be automatically included if you load page-class
''					because many other classes need this class. Its a basic-control
'' @VERSION:		1.03

'**************************************************************************************************************
class createDropdown
	
	private fieldsToShow
	private displayfunction 'if you want to do something with an value. e.g. cut the field to a number of letter or different colors
	
	public sqlquery			''[string] The Sql-Query for the Dropdown. If you dont want an sql then enter a string sperated with ":"
	public pk				''[string] PrimaryKey. Will be shown as Value of every OPTION-tag. Must be available in the sql-Statement. If you use a ':' separated string as sql then PK also needs a ':' sperated string with the same amount of fields.
	public name				''[string] Name of the dropdown
	public idToMatch		''[string/array] what option(s) should be selected. if you want to select more options in a multiple dropdown you need to provide an array
	public isMultiple		''[bool] multiple select dropdown
	public cssClass			''[string] if you have a cssclass put the name of the class in here.
	public commonTxt		''[string] just like "please select a value from the dropdown".
	public commonTxtVal		''[string] Value for the commonTxtvalue. Default = 0
	public onAttribute		''[string] if you want to execute e.g. a javascript onclick, onchange, etc. e.g. "onclick=''"
	public fieldSeperator	''[string] if you want to show more than one field per OPTION-tag you have the option to put a seperator in here
	public isDisabled		''[bool] dropdown will be disabled
	public tdAddy			''[string] if you want to add something yourself to the dropdown like style= width=,etc. It is quite the same as "onAttribute" property
	public rows				''[int] For multiple Dropdown. How many rows?
	public forceArray		''[bool] Sometimes you need to say its an array instead of sql query. e.g. just one value
	public id				''[string] The ID for the control. Displayed in the ID-Attribute
	public enableAutosplit	''[bool] prevent auto split if there is a ":" in your sql. e.g. if you have time-statements in your query
	public records			''[recordset] recordset. you can access using displayfunction
	
	private sub Class_Initialize()
		fieldsToShow	= empty
		sqlquery		= empty
		pk				= empty
		name			= empty
		idToMatch		= 0
		isMultiple		= false
		cssClass		= empty
		commonTxt		= empty
		onAttribute		= empty
		fieldSeperator	= "&nbsp;"
		isDisabled		= false
		commonTxtVal	= 0
		tdAddy			= empty
		displayfunction = empty
		rows			= 0
		forceArray		= false
		id				= empty
		enableAutosplit = true
		set records		= nothing
	end sub
	
	'************************************************************************************************************
	''@SDESCRIPTION:	Specifies a function to execute on every itteration
	''@DESCRIPTION:		The function will be executed on every field from the dropdown. So it is important
	''					you have one input to the function which holds the current value. Thats why
	''					it is important to tell the function which parameter it is.
	''					Example: object.set_displayFunction("checkme(myField,10)","myField")
	''@PARAM:			- myFunction [string]: Name of the function to execute
	''@PARAM:			- param [string]: Whats the name of the parameter which stores the value.
	'************************************************************************************************************
	public function set_displayFunction(myFunction, param)
		if not myFunction = empty and not param = empty then
			displayfunction = replace(myFunction,param,"var_name")
		else
			response.write "<BR><strong>Erroro! </strong>Displayfunction needs all parameters.<BR>"
			response.end
		end if
	end function
	
	'************************************************************************************************************
	''@SDESCRIPTION:	What field from the sql should be showed in the dropdown?
	''@DESCRIPTION:		Add as many fields you want. If you use an array then you dont need this method.
	''@PARAM:			- field [string]: Name of the field. Must be selected in sql-query
	'************************************************************************************************************
	public function addDisplayField(field)
		fieldsToShow = fieldsToShow & field & ","
	end function
	
	'************************************************************************************************************
	''@SDESCRIPTION:	Draws our nice new dropdown
	'************************************************************************************************************
	public function draw()
		response.write generate_select()
	end function
	
	'************************************************************************************************************
	''@SDESCRIPTION:	You can also get the whole dropdown back as String if you want to
	''@RETURN:			[string] Returns the dropdown as a string
	'************************************************************************************************************
	public function getAsString()
		getAsString = generate_select()
	end function
	
	'******************************************************************************************************************
	' isSelected 
	'******************************************************************************************************************
	private function isSelected(currentVal)
		isSelected = false
		
		if not isMultiple then
			if cstr(currentVal) = cstr(idToMatch) then isSelected = true
		else
			'if its a multiple dropdown we check if the idTomatch is an array, so we have to select more options.
			if isArray(idToMatch) then
				found = -1
				'it is an array, so we have to check the value exist there.
				for i = 0 to uBound(idToMatch)
					'it exist
					if cStr(currentVal) = cStr(idToMatch(i)) then
						found = i
						exit for
					end if
				next
				
				'last but not least we take the last value and place it on the current place, then we resize the array
				'so on the next loop we dont have to loop again through the whole array.
				if found > -1 then
					isSelected = true
					idToMatch(found) = idToMatch(uBound(idToMatch))
					redim preserve idToMatch(uBound(idToMatch) - 1)
				end if
			else
				if cstr(currentVal) = cstr(idToMatch) then isSelected = true
			end if
		end if
	end function
	
	'******************************************************************************************************************
	' generate_select
	'******************************************************************************************************************
	private function generate_select()
		select_str = empty
		
		'we check if there is a valid Sql-query or not
		if (instr(sqlquery,":") OR forceArray) AND enableAutosplit then
			'we create dropdown from a common array
			arrayField = true
		else
			'we create dropdown from sql-query
			arrayField = false
		end if
		
		'if there is an individual css-class needed we get it
		if cssClass = empty then
			css_class = empty
		else
			css_class = " class='" & cssClass & "'"
		end if
		
		'if we need a onclick, onchange, etc.
		if not onAttribute = empty then
			onClick = " " & onAttribute
		else
			onClick = empty
		end if
		
		'maybe we want to disable the dropdown.
		if isDisabled then
			disabled = " disabled"
		else
			disabled = empty
		end if
		
		if isMultiple then
			multipleStr = " multiple"
		else
			multipleStr = empty
		end if
		
		if not rows = 0 then
			myRows = " size=" & rows
		else
			myRows = empty
		end if
		
		if not id = empty then
			myControlID = " id=" & id
		else
			myControlID = empty
		end if
		
		'we begin to write the dropdown. the beginning
		select_str = select_str & "<SELECT NAME=""" & name & """ " & myControlID & multipleStr & css_class & onClick & disabled & " " & tdAddy & myRows & ">" & vbcrlf
		if not commonTxt = "" then
			var_name = commonTxt
			
			if isSelected(commonTxtVal) then
				selectd = " selected"
			else
				selectd = ""
			end if
			
			select_str = select_str & "	<OPTION VALUE=""" & commonTxtVal & """" & selectd & ">" & var_name & "</option>" & vbcrlf
		end if
		
		'now we set some things depends on array or sql
		if not arrayField then 'we make a dropdown with SQL
			set records = lib.getRecordset(sqlquery)
			nameArr = split(fieldsToShow, ",")
			while not records.eof 
				var_id = records(pk)
				var_name = ""
				for i = lbound(nameArr) to ubound(nameArr) - 1
					if i = 0 then
						var_name = records(nameArr(i))
					else
						var_name = var_name & fieldSeperator & records(nameArr(i))
					end if
				next
				'if we have a function to execute on our field then do it now
				if not displayfunction = empty then
					execute("var_name = " & displayfunction)
				end if
				
				if isSelected(var_id) then
					select_str = select_str & "	<OPTION VALUE=""" & var_id & """ SELECTED>" & var_name & "</OPTION>" & vbcrlf
				else
					select_str = select_str & "	<OPTION VALUE=""" & var_id & """>" & var_name & "</OPTION>" & vbcrlf
				end if
				records.movenext()
			wend
			set records = nothing
		else 'we make dropdown with common ARRAY
			myArr = split(sqlquery,":")
			
			'if we have also an array in pk then we split it. it must have the same number of fields as the myArr field
			if not pk = empty then
				myPk = split(pk,":")
			else
				myPk = split(sqlquery,":")
			end if
			
			for i = lbound(myArr) to ubound(myArr)
				var_name = myArr(i)
				'if we have a function to execute on our field then do it now
				if not displayfunction = empty then
					execute("var_name = " & displayfunction)
				end if
				
				if isSelected(myPk(i)) then
					select_str = select_str & "	<OPTION VALUE=""" & myPk(i) & """ SELECTED>" & var_name & "</OPTION>" & vbcrlf
				else
					select_str = select_str & "	<OPTION VALUE=""" & myPk(i) & """>" & var_name & "</OPTION>" & vbcrlf
				end if
			next
		end if
		
		'we close the dropdown. end
		select_str = select_str & "</SELECT>"
		generate_select = select_str
			
	end function
	
end class
lib.registerClass("CreateDropdown")
%>