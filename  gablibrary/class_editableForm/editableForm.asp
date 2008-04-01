<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		editableForm
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		2005-02-05 11:46
'' @CDESCRIPTION:	OBSOLETE! use Form, Validateable instead. easy possibility to make your forms editable
'' @VERSION:		0.1

'**************************************************************************************************************
class editableForm

	'private members
	private p_ID
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	private sub class_initialize()
		p_ID = 0
	end sub
	
	public property let id(val) ''[int] sets the ID of the current form.
		if val <> "" and isNumeric(val) then p_ID = cInt(val)
	end property
	
	public property get id() ''[int] gets the ID of the current form. 0 if no id given.
		id = p_ID
	end property
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	indicates if the form is to create a new entity or to modify one
	'' @RETURN:			[bool] true if its a new one
	'**********************************************************************************************************
	public function isNew()
		isNew = (not (id > 0))
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	returns the value for a given form-field. 
	'' @DESCRIPTION:	if your form has been posted then the posted value will be returned, etc.
	'' @RETURN:			[variant] your value
	'**********************************************************************************************************
	public function valueFor(fieldName, defaultNewValue, defaultEditValue)
		if request.form.count > 0 then
			valueFor = request.form(fieldName)
		else
			if isNew() then
				valueFor = defaultNewValue & ""
			else
				valueFor = defaultEditValue & ""
			end if
		end if
		
		valueFor = encQuotesForFormFields(valueFor)
	end function
	
	'********************************************************************************************************
	'' @SDESCRIPTION:	OBSOLETE! use str.HTMLEncode() or lib.RFE()
	'********************************************************************************************************
	public function encQuotesForFormFields(val)
		encQuotesForFormFields = server.HTMLencode(val & "")
	end function
	
	'********************************************************************************************************
	'' @SDESCRIPTION:	OBSOLETE! use str.HTMLEncode() or lib.RFE()
	'********************************************************************************************************
	public function formEncode(val)
		formEncode = server.HTMLencode(val & "")
	end function

end class
%>