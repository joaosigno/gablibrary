<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		Validateable
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		2005-02-06 15:25
'' @CDESCRIPTION:	Represents a validation container which can be used for the validation of business objects.
''					or any other kind of validation.
''					It stores invalid fields (e.g. property of a class) with an associated error message
''					(why is the field invalid). The valid ones are not needed because they are valid anyway ;)
''					Example for usage:
'' @POSTFIX:		val
'' @VERSION:		0.2

'**************************************************************************************************************
class Validateable

	'private members
	private uniqueID
	
	'protected members
	public dictInvalidData		''[dictionary] holds the invalid field. get it only with getInvalidData()
	public reflectItemPrefix	''[string] the prefix for each item within the summary which is returned on reflection. default = "<li>"
	public reflectItemPostfix	''[string] the prefix for each item within the summary which is returned on reflection. default = "</li>"
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	private sub class_initialize()
		set dictInvalidData	= server.createObject("scripting.dictionary")
		uniqueID = 0
		reflectItemPrefix = "<li>"
		reflectItemPostfix = "</li>"
	end sub
	
	'**************************************************************************************************************
	'' @SDECRIPTION:	Returns a dictionary with all invalid-fields 
	'' @RETURN:			[dictionary] with all descriptions and fieldnames of invalid fields
	'**************************************************************************************************************
	public function getInvalidData()
		set getInvalidData = dictInvalidData
	end function
	
	'**************************************************************************************************************
	'' @SDECRIPTION:	Returns the description of the error for the requested-field.
	'' @PARAM:			fieldName [string]: the name of your field to get the error-description for.
	'' @RETURN:			[string] the description of the error for the requested-field. empty if there isnt any error
	'**************************************************************************************************************
	public function getInvalidDescriptionFor(fieldName)
		fieldName = uCase(fieldName)
		if dictInvalidData.exists(fieldName) then
			getInvalidDescriptionFor = dictInvalidData(fieldName)
		else
			getInvalidDescriptionFor = empty
		end if
	end function
	
	'**************************************************************************************************************
	'' @SDECRIPTION:	Returns true if the requested fieldname is invalid
	'' @PARAM:			fieldName [string]: the name of your field to check.
	'' @RETURN:			[bool] true if the field is invalid
	'**************************************************************************************************************
	public function fieldIsInvalid(byVal fieldName)
		fieldName = uCase(fieldName)
		fieldIsInvalid = false
		if dictInvalidData.exists(fieldName) then
			fieldIsInvalid = true
		end if
	end function
	
	'**************************************************************************************************************
	'' @SDECRIPTION:	Returns true if everything is valid
	'' @RETURN:			[bool] true if is valid
	'**************************************************************************************************************
	public function isValid()
		isValid = (dictInvalidData.count <= 0)
	end function
	
	'**************************************************************************************************************
	'' @SDECRIPTION:	Adds a new invalid field. only if it does not exists yet.
	'' @PARAM:			fieldName [string]: the name of your field. leave empty if you want the field be auto-generated.
	'' @PARAM:			errorDescription [string]: a reason why the field is invalid
	'' @RETURN:			[bool] true if added, false if not added
	'**************************************************************************************************************
	public function add(byVal fieldName, errorDescription)
		if fieldname = empty then fieldName = getUniqueID()
		if not dictInvalidData.exists(uCase(fieldName)) then
			dictInvalidData.add uCase(fieldName), errorDescription
			add = true
		end if
	end function
	
	'**************************************************************************************************************
	'' @SDECRIPTION:	returns a custom formatted error-summary.
	'' @DESCRIPTION:	usefull if you want to show the errors for example in a list. summary will be just
	''					returned if there are any error. (so at least one field must be invalid)
	'' @PARAM:			overallPrefix [string]: prefix for the whole summary e.g. <ul>
	'' @PARAM:			overallPostfix [string]: postfix for the whole summary e.g. </ul>
	'' @PARAM:			itemPrefix [string]: prefix for each item <li>
	'' @PARAM:			itemPostfix [string]: prefix for each item </li>
	'' @RETURN:			[string] formatted error-summary
	'**************************************************************************************************************
	public function getErrorSummary(overallPrefix, overallPostfix, itemPrefix, itemPostfix)
		getErrorSummary = empty
		if not isValid() then
			getErrorSummary = getErrorSummary & overallPrefix
			for each key in dictInvalidData.keys
				getErrorSummary = getErrorSummary & itemPrefix & dictInvalidData(key) & itemPostfix
			next
			getErrorSummary = getErrorSummary & overallPostfix
		end if
	end function
	
	'**************************************************************************************************************
	'' @SDECRIPTION:	OBSOLETE! use add() instead
	'**************************************************************************************************************
	public sub addInvalidField(byVal fieldName, errorDescription)
		me.add fieldName, errorDescription
	end sub
	
	'**************************************************************************************************************
	'* getUniqueID 
	'**************************************************************************************************************
	private function getUniqueID()
		uniqueID = uniqueID + 1
		getUniqueID = uniqueID + 1
	end function
	
	'**************************************************************************************************************
	'' @SDECRIPTION:	reflection.
	'' @DESCRIPTION:	as the class has no real properties the status is exposed with:
	''					- data: holds a dictionary with the invalid fields
	''					- isValid: indicates if its valid or not
	''					- summary: holds a summary of the invalid data (reflectItemPrefix and reflectItemPostfix can be used to format the items)
	'**************************************************************************************************************
	public function reflect()
		set reflect = lib.newDict(empty)
		with reflect
			.add "isValid", isValid()
			.add "data", getInvalidData()
			.add "summary", getErrorSummary("", "", reflectItemPrefix, reflectItemPostfix)
		end with
	end function

end class
lib.registerClass("Validateable")
%>