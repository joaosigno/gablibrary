<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		ExcelExporter
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		2005-09-05 12:14
'' @CDESCRIPTION:	Gives you the opportunity to export HTML or Plaintext into excel
''					it takes the POST-data from a posted form and sends it to the browser
''					using the EXCEL-contenttype. So everything you need to do is to create a page
''					where you use this class and then post your data into this page
'' @VERSION:		0.1

'**************************************************************************************************************
class ExcelExporter

	'private members
	private maxHiddenFieldLength		
	private p_output					
	private stringBuilderInstanciated	
	
	'public members
	public useStringBuilder				''[bool] use stringBuilder for concating the output string?
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		maxHiddenFieldLength			= 100000
		useStringBuilder				= lib.useStringBuilder
		stringBuilderInstanciated		= false
	end sub
	
	public property get output ''[string] gets the current output-string
		if useStringBuilder then
			output = p_output.toString()
		else
			output = p_output
		end if
	end property
	
	'Format-styles: useful if you want to format your excelsheet and dont remember these cryptic values ;)
	public property get formatStyleText ''[string] gets the style attribute to force a column to appear as a text-column
		formatStyleText = "mso-number-format:\@"
	end property
	
	'**********************************************************************************************************
	'* destructor 
	'**********************************************************************************************************
	private sub class_terminate()
		set p_output = nothing
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	exports the submitted fields
	'' @DESCRIPTION:	warning: posted data in a hidden-field is limited by size. so if you expect
	''					more data being posted, you need to split up your data and use more than one
	''					hidden-field. useful function: str.divide()
	'**********************************************************************************************************
	public sub export()
		response.contentType = "application/vnd.ms-excel"
		response.addHeader "Content-Disposition", "attachment"
		
		for each field in lib.form
			str.write(lib.RF(field))
		next
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	exports the saved output
	'**********************************************************************************************************
	public sub exportOutput()
		response.contentType = "application/vnd.ms-excel"
		response.addHeader "Content-Disposition", "attachment"
		
		str.writeln(output)
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	adds an string to the excel-output
	'' @DESCRIPTION:	uses the stringbuilder-append() method if stringbuilder is active.
	'' @PARAM			inputStr [string]: ATTENTION: only strings with a maximum length of 40000 chars allowed
	'**********************************************************************************************************
	public sub addOutput(inputStr)
		if useStringBuilder then
			if not stringBuilderInstanciated then
				set p_output = server.createObject("StringBuilderVB.StringBuilder")
				p_output.init 40000, 7500
				stringBuilderInstanciated = true
			end if
			p_output.append(inputStr)
		else
			p_output = p_output & inputStr
		end if
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	generates hidden-fields with your given input
	'' @DESCRIPTION:	generates hidden-fields with the given input. these hidden-inputs can be used in your
	''					form which will be submitted to the excelExporter. if the length of the output
	''					is larger than the allowed length of a hidden-field a divide is done.
	'**********************************************************************************************************
	public function getHiddenFields()
		parts = str.divide(output, maxHiddenFieldLength)
		for i = 0 to uBound(parts)
			getHiddenFields = getHiddenFields & str.getHiddenInput("xlText" & i, parts(i))
		next
	end function
	
	'********************************************************************************************************
	'' @SDESCRIPTION:	encodes double-quotes to &quot; so it can be displayed in input-fields
	'********************************************************************************************************
	public function encodeData(val)
		encodeData = val
		'using instr before replacing used to be faster
		if inStr(val, """") > 1 then encodeData = replace(val, """", "&quot;")
	end function

end class
lib.registerClass("ExcelExporter")
%>