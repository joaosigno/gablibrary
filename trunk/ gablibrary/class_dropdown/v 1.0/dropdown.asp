<!--#include virtual="/gab_Library/class_dropdown/class_dropdownItem.asp"-->
<!--#include virtual="/gab_Library/class_dropdown/class_dropdownConnector.asp"-->
<%
'**************************************************************************************************************
'* GABLIB Copyright (C) 2003 
'**************************************************************************************************************
'* This program is free software; you can redistribute it and/or modify it under the terms of
'* the GNU General Publ. License as published by the Free Software Foundation; either version
'* 2 of the License, or (at your option) any later version.
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		Dropdown
'' @CREATOR:		Michal Gabrukiewicz - gabru@gmx.at
'' @CREATEDON:		2005-02-01 16:31
'' @CDESCRIPTION:	represents a HTML-selectbox in an OO-approach.
''					naming and functionallity is based on .NET-control dropdownList
'' @VERSION:		1.0

'**************************************************************************************************************

const DD_DATASOURCE_ARRAY			= 1
const DD_DATASOURCE_RECORDSET		= 2
const DD_DATASOURCE_DICTIONARY		= 4
const DD_OUTPUT_DIRECT				= 1
const DD_OUTPUT_STRING				= 2
const DD_SELECTIONTYPE_COMMON		= 1
const DD_SELECTIONTYPE_MULTIPLE		= 2
const DD_SELECTIONTYPE_SINGLE		= 4

class Dropdown

	'private members
	private output					'[string], [stringBuilder] holds the complete output-string if dropdown will be returned as string
	private outputMethod			'[int] 0 = direct output, 1 = output to string
	private datasourceType			'[int] 1 = array, 2 = adodb.recordset, 4 = dictionary
	private currentIteration		'[int] index of the current iteration for the items
	private tmpDatasourceLength		'[int] temporary stored length of datasource. saves time on looping!
	private datasourceItems			'[collection]
	private datasourceKeys			'[collection]
	private p_selectedValue			'[string], [array]
	private selectedFound			'[bool]
	
	'public members
	public name						''[string] Name of the control
	public datasource				''[array] datasource as array.
									''[recordset] datasource as recordset. be sure to use lib.getUnlockedRecordset!
									''[dictionary] datasource as dicitionary
									''[string] if datasource is a string it will be interpreted as a SQL-query. after draw() datasource changes to a recordset.
	public ID						''[string] ID of the control
	public onItemCreated			''[string] name of function which should handle onItemCreated. Event will be raised just before printing the item.
	public attributes				''[string] additional attributes which go into the &lt;select>-element. e.g. onClick, onChange, etc.
	public style					''[string] css-Styles for the control
	public cssClass					''[string] name of the css-class you want to assign to the control
	public valuesDatasource			''[array] if array is used as datasource, please provide an array of same length for values too. If no array is given, values are indexed
	public dataTextField			''[string], [int] field-name of the datasource that provides the text-content for each list-item.
	public dataValueField			''[string], [int] field-name of the datasource that provides the value for each list-item
	public tabIndex					''[int] index of the control. -1 = dont add to the tabindex-collection
	public multiple					''[bool] is it a multiple dropdown? default = false
	public size						''[int] number of displayed rows if its a multiple dropdown. default = 1. If its not a common-multiple-dropdown
									''then the size is used as pixels for the height of the dropdown.
	public disabled					''[bool] indicates whether the control is disabled or not. default = false
	public useStringBuilder			''[bool] indicates if stringbuilder should be used for rendering. DLL-must be installed. default = false
	public commonFieldText			''[string] text of the common-field. e.g. --- please select a value ---
	public commonFieldValue			''[string] value for the common-field. default = 0
	public multipleSelectionType	''[int] a value of the SELECTIONTYPE-Enumeration. Set this property if you want to change the
									''selectiontype of a multiple dropdown. Useful if you want to have formatted ITEMS,
									''because the dropdown isnt rendered as a SELECT-Element then. There are 3 variations:
									''DD_SELECTIONTYPE_COMMON = its the common one. hold down CTRL to select more
									''DD_SELECTIONTYPE_MULTIPLE = selection comes with checkboxes.
									''DD_SELECTIONTYPE_SINGLE = selection comes with radiobuttons. so just one selection is allowed.
	public autoDrawItems			''[bool] defines if each item will be drawn autmatically after the onItemCreated-Event
									''has been raised. default = true. disabling usefull if you want to add items during the runtime
	public connectedTo				''[DropdownConnector] connect the dropdown to another one using a DropdownConnector
	
	public property let selectedValue(val) ''[string], [array] what value(s) is selected. array needed if multiple dropdown
		if isArray(val) then
			p_selectedValue = val
		else
			redim p_selectedValue(0)
			p_selectedValue(0) = val & "" 'concat to allow NULLS
		end if
	end property
	
	public property get selectedValue ''[array], [string] returns the selected value(s)
		if uBound(p_selectedValue) = 0 then
			selectedValue = p_selectedValue(0)
		else
			selectedValue = p_selectedValue(0)
		end if
	end property
	
	public property get controlLocation ''[string] gets the location of the control
		controlLocation = consts.gabLibLocation & "class_dropdown/"
	end property
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		tmpDatasourceLength			= -1
		size						= 1
		currentIteration			= 0
		commonFieldValue			= 0
		p_selectedValue				= array()
		valuesDatasource			= array()
		multiple					= false
		disabled					= false
		useStringBuilder			= false
		autoDrawItems				= true
		selectedFound				= false
		commonFieldText				= empty
		output	 					= empty
		ID							= empty
		connectedTo					= empty
		outputMethod				= DD_OUTPUT_DIRECT
		datasourceType				= DD_DATASOURCE_ARRAY
		multipleSelectionType		= DD_SELECTIONTYPE_COMMON
	end sub
	
	'**********************************************************************************************************
	'* destructor 
	'**********************************************************************************************************
	private sub class_terminate()
		if useStringBuilder then set output = nothing
		set connectedTo = nothing
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	draws the dropdown on the page
	'**********************************************************************************************************
	public sub draw()
		outputMethod = DD_OUTPUT_DIRECT
		renderControl()
		if useStringBuilder then response.write(getOutput())
	end sub
	
	'**********************************************************************************************************
	'' @DESCRIPTION:	returns a new connector to another dropdown. it autmatically connects these two
	'' @SDESCRIPTION:	factory-method. you can assign the connector manually by using the connectedTo-property
	'' 					on both dropdowns. this method makes all this for you.
	'' @PARAM:			target [Dropdown]: the dropdown you want to connect this dropdown to.
	'' @RETURN:			[DropdownConnector]: returns a Connector-object where you can set all the specs
	'**********************************************************************************************************
	public function getNewConnector(byRef target)
		set getNewConnector = new DropdownConnector
		with getNewConnector
			set .source = me
			set .target = target
		end with
		set connectedTo = getNewConnector
		multipleSelectionType = DD_SELECTIONTYPE_COMMON
		multiple = true
		set target.connectedTo = getNewConnector
		target.multipleSelectionType = DD_SELECTIONTYPE_COMMON
		target.multiple = true
	end function
	
	'**********************************************************************************************************
	'' @DESCRIPTION:	gets you a new dropdown-item which can be drawn after that.
	'' @PARAM:			value [string] value of the item
	'' @PARAM:			text [string] text of the item
	'' @RETURN:			[DropdownItem] the created dropdownItem
	'**********************************************************************************************************
	public function getNewItem(itemValue, itemText)
		set getNewItem = getRawItem(getDatasourceLength() + 1, itemValue, itemText)
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	returns the dropdown as a string
	'**********************************************************************************************************
	public function toString()
		outputMethod = DD_OUTPUT_STRING
		renderControl()
		toString = getOutput()
	end function
	
	'**********************************************************************************************************
	' renderControl 
	'**********************************************************************************************************
	private function renderControl()
		selectedFound = false
		initStringBuilder()
		determineDatasource()
		
		printBeginTag()
		
		if commonFieldText <> empty and not multiple then addItem(getRawItem(-1, commonFieldValue, commonFieldText))
		
		for currentIteration = 0 to getDatasourceLength()
			addItem(getRawItem(currentIteration, getCurrentItemValue(), getCurrentItemText()))
			moveDatasourceCursor()
		next
		
		printEndTag()
	end function
	
	'**************************************************************************************************************
	' printBeginTag 
	'**************************************************************************************************************
	private sub printBeginTag()
		if isCommonDropdown() then
			print("<select" & _
				getAttribute("tabindex", tabindex) & _
				getAttribute("size", size) & _
				getAttribute("name", name) & _
				getAttribute("style", style) & _
				lib.iif(multiple, " multiple ", empty))
		else
			print("<div") & getAttribute("style", "white-space:nowrap;overflow:auto;height:" & size & "px;border:2 inset #eee;" & style)
		end if
		
		print(getAttribute("id", ID) & _
			getAttribute("class", cssClass) & _
			lib.iif(disabled, " disabled", empty) & _
			" " & attributes & _
			">" & vbCrLf)
	end sub
	
	'**************************************************************************************************************
	' printEndTag 
	'**************************************************************************************************************
	private sub printEndTag()
		if isCommonDropdown() then
			print("</select>")
		else
			print("</div>")
		end if
		print(vbCrLf)
	end sub
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	gets an Attribute HTML-string
	'' @DESCRIPTION:	just for internal use. public because the dropdownItem-class needs it too.
	'**************************************************************************************************************
	public function getAttribute(attributeName, attributeValue)
		if attributeValue <> "" then getAttribute = " " & attributeName & "=""" & attributeValue & """"
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	common Dropdown means that it will be rendered as a SELECT-Tag
	'' @DESCRIPTION:	just for internal use. public because the dropdownItem-class needs it too.
	'**************************************************************************************************************
	public function isCommonDropdown()
		isCommonDropdown = (not multiple or (multiple and multipleSelectionType = DD_SELECTIONTYPE_COMMON))
	end function
	
	'**************************************************************************************************************
	' initStringBuilder 
	'**************************************************************************************************************
	private sub initStringBuilder()
		if useStringBuilder then
			set output = Server.CreateObject("StringBuilderVB.StringBuilder")
			output.init 40000, 7500
		end if
	end sub
	
	'**********************************************************************************************************
	' getDropdownItem 
	'**********************************************************************************************************
	private function getRawItem(itemIndex, itemValue, itemText)
		set getRawItem = new dropdownItem
		with getRawItem
			set .dropdown = me
			.index = itemIndex
			.value = itemValue & ""
			.text = itemText & ""
		end with
	end function
	
	'**********************************************************************************************************
	'* addItem 
	'**********************************************************************************************************
	private sub addItem(item)
		item.selected = isSelected(item.value)
		raiseOnItemCreatedEvent(item)
		if autoDrawItems then item.draw()
	end sub
	
	'**********************************************************************************************************
	' isSelected 
	'**********************************************************************************************************
	private function isSelected(valueToCheck)
		if not selectedFound then
			for i = 0 to uBound(p_selectedValue)
				if cStr(p_selectedValue(i)) = valueToCheck then
					isSelected = true
					selectedFound = (not multiple)
					exit for
				end if
			next
		end if
	end function
	
	'**********************************************************************************************************
	' raiseOnItemCreatedEvent 
	'**********************************************************************************************************
	private sub raiseOnItemCreatedEvent(byRef currentDropdownItem)
		if onItemCreated <> empty then execute(onItemCreated & "(currentDropdownItem)")
	end sub
	
	'**********************************************************************************************************
	' determineDatasource 
	'**********************************************************************************************************
	private sub determineDatasource()
		if isArray(datasource) then
			datasourceType = DD_DATASOURCE_ARRAY
		else
			select case lCase(typename(datasource))
				case "recordset"
					datasourceType = DD_DATASOURCE_RECORDSET
				case "dictionary"
					datasourceItems = datasource.items
					datasourceKeys = datasource.keys
					datasourceType = DD_DATASOURCE_DICTIONARY
				case else 'its a string
					set datasource = lib.getUnlockedRecordset(datasource)
					datasourceType = DD_DATASOURCE_RECORDSET
			end select
		end if
	end sub
	
	'**********************************************************************************************************
	' getDatasourceLength 
	'' @RETURN: 	[int] gets the length of the datasource
	'**********************************************************************************************************
	private function getDatasourceLength()
		if tmpDatasourceLength = -1 then
			select case datasourceType
				case DD_DATASOURCE_ARRAY
					tmpDatasourceLength = uBound(datasource)
				case DD_DATASOURCE_RECORDSET
					tmpDatasourceLength = datasource.recordCount - 1
				case DD_DATASOURCE_DICTIONARY
					tmpDatasourceLength = datasource.count - 1
			end select
		end if
		getDatasourceLength = tmpDatasourceLength
	end function
	
	'**********************************************************************************************************
	' geCurrentItemValue 
	'' @RETURN: 	[string] gets the value of the item on current iteration
	'**********************************************************************************************************
	private function getCurrentItemValue()
		select case datasourceType
			case DD_DATASOURCE_ARRAY
				if uBound(valuesDatasource) > -1 then
					getCurrentItemValue = valuesDatasource(currentIteration)
				else
					getCurrentItemValue = datasource(currentIteration)
				end if
			case DD_DATASOURCE_RECORDSET
				getCurrentItemValue = datasource.fields(dataValuefield)
			case DD_DATASOURCE_DICTIONARY
				getCurrentItemValue = datasourceKeys(currentIteration)
		end select
	end function
	
	'**********************************************************************************************************
	' getCurrentItemText 
	'' @RETURN: 	[string] gets the text of the item on current iteration
	'**********************************************************************************************************
	private function getCurrentItemText()
		select case datasourceType
			case DD_DATASOURCE_ARRAY
				getCurrentItemText = datasource(currentIteration)
			case DD_DATASOURCE_RECORDSET
				getCurrentItemText = datasource.fields(dataTextfield)
			case DD_DATASOURCE_DICTIONARY
				getCurrentItemText = datasourceItems(currentIteration)
		end select
	end function
	
	'**********************************************************************************************************
	' moveDatasourceCursor 
	'**********************************************************************************************************
	private sub moveDatasourceCursor()
		select case datasourceType
			case DD_DATASOURCE_RECORDSET
				datasource.movenext()
		end select
	end sub
	
	'**********************************************************************************************************
	' getOutput 
	'**********************************************************************************************************
	private function getOutput()
		if useStringBuilder then
			getOutput = output.toString()
		else
			getOutput = output
		end if
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	prints out to the output
	'' @DESCRIPTION:	just for internal use. public because the dropdownItem-class needs it too.
	'**************************************************************************************************************
	public sub print(value)
		if useStringBuilder then
			output.append(value)
		else
			if outputMethod = DD_OUTPUT_DIRECT then
				response.write(value)
			else
				output = output & value
			end if
		end if
	end sub

end class
%>