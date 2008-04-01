<!--#include file="const.asp"-->
<%
'**************************************************************************************************************
'* GABLIB Copyright (C) 2003 - This file is part of GABLIB - http://gablib.grafix.at
'**************************************************************************************************************
'* This program is free software; you can redistribute it and/or modify it under the terms of
'* the GNU General Publ. License as published by the Free Software Foundation; either version
'* 2 of the License, or (at your option) any later version. 
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		TextTemplate
'' @CREATOR:		Michal Gabrukiewicz - gabru@gmx.at
'' @CREATEDON:		24.10.2003
'' @CDESCRIPTION:	This class "imports" all data from a specified file, replaces the specified variables
''					with values and returns a string. Practice: e.g. you want to make a template for an email
''					Then just store the email in a file e.g. email.txt and there you put the text including
''					Placeholders. The class will import the template and replace the place-holders with vars.
'' @VERSION:		0.99

'**************************************************************************************************************

class textTemplate

	private version
	private DICTvariables
	
	public fileName					''The filename of your template.
	public placeHolderBegin			''If you want to use your own placeholder characters. this is the beginning. e.g. <<<
	public placeHolderEnd			''If you want to use your own placeholder characters. this is the ending. e.g. >>>
	public customLineBreak			''You need a custom line-break? e.g. chr(13) or vbcrlf, etc.
	
	'******************************************************************************************************************
	'* constructor 
	'******************************************************************************************************************
	private sub Class_Initialize()
		version				= "TextTemplate Class - ver 0.1"
		set DICTvariables	= Server.createObject("Scripting.Dictionary")
		fileName			= empty
		placeHolderBegin	= DEFAULT_PLACEHOLDER_BEGIN
		placeHolderEnd		= DEFAULT_PLACEHOLDER_END
		customLineBreak		= DEFAULT_LINEBREAK
	end sub
	
	'******************************************************************************************************************
	'* destructor 
	'******************************************************************************************************************
	public sub class_terminate()
		set DICTvariables = nothing
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Adds a variable which should be replaced
	'' @DESCRIPTION:	Adds a variable. All placeholders in the template using this Variable will be replaced by the
	''					value of the variable. if the value was already added then it will be replaced by the new one
	'' @PARAM:			varName [string]: The name of you variable (Important: every name just once!)
	'' @PARAM:			varValue [string]: The value you want to put instead the varName
	'******************************************************************************************************************
	public sub addVariable(varName, varValue)
		if DICTvariables.exists(varName) then DICTvariables.remove(varName)
		DICTvariables.add varName, varValue
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Returns a string where the template and the placeholders are merged
	'******************************************************************************************************************
	public function returnString()
		fileToOpen = server.MapPath(fileName)
		set sys = server.CreateObject("Scripting.FileSystemObject")
		
		'we check if file exists
		if sys.FileExists(fileToOpen) then
			'we open the file
			set myfile = sys.openTextfile(fileToOpen,1,false) '1 = for Reading only ; 3 Param => dont create file when not exists
			
			'now we read every line from the file
			while not myFile.AtEndOfStream
				fileContent = fileContent & myfile.ReadLine() & customLineBreak
			wend
			
			'now replace the placeholders and return the whole String
			toReturn = replace_PlaceHolders(fileContent)
			
			myfile.close
			set myfile = nothing
		else 'file does not exist
			toReturn = "Error: Template-file does not exist."
		end if
		
		set sys = nothing
		
		'we return what should be returned
		returnString = toReturn
	end function
	
	'******************************************************************************************************************
	'* replace_PlaceHolders 
	'******************************************************************************************************************
	private function replace_PlaceHolders(input)
		tmp = input
		for each key in DICTvariables.keys
			toReplace = ucase(placeHolderBegin & key & placeHolderEnd)
			replaceBy = DICTvariables(key)
			tmp = replace(tmp, ucase(toReplace), checkIfNull(replaceBy))
		next
		
		'return the finished string
		replace_PlaceHolders = tmp
	end function
	
	'******************************************************************************************************************
	'* checkIfNull 
	'******************************************************************************************************************
	private function checkIfNull(val)
		if isnull(val) then
			checkIfNull = ""
		else
			checkIfNull = val
		end if
	end function

end class
%>