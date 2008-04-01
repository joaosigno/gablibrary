<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		TableSessionObject
'' @CREATOR:		Michael Rebec
'' @CREATEDON:		31.03.2004
'' @CDESCRIPTION:	Attention: the name "tableObject" is used in the ErrorHandler Class !!!
'' @VERSION:		0.2.1

'**************************************************************************************************************
class tableSessionObject

	'************************************************************************************************************
	'* getURLCompareResult
	'************************************************************************************************************
	private function getURLCompareResult(objURL)
		if objURL = Request.ServerVariables("URL") then
			getURLCompareResult = true
		else
			getURLCompareResult = false
		end if
	end function
	
	'************************************************************************************************************
	''@SDESCRIPTION:	Get the active page number
	''@DESCRIPTION:		Get the active page number saved in the SO (Session Object), or the Request("..nr..")
	''@RETURN:			[int] the page number
	'************************************************************************************************************
	public function getAbsolutePage()
		if Request.Form.Count > 0 then
			getAbsolutePage = Request.form("actualPageNumber")
		elseif len(getSessionObject("pagenumber")) >= 1 then
			getAbsolutePage	= CInt(getSessionObject("pagenumber"))
		end if
	end function
	
	'************************************************************************************************************
	''@SDESCRIPTION:	Get a specified filter value
	''@DESCRIPTION:		Get a specified filter value saved in the SO, or the Request("..filter..")
	''@PARAM:			- filterValue: the key of the filter
	''@RETURN:			[string] the filter value if the SO exists, or the Request if postback happend
	'************************************************************************************************************
	public function getFilterValue(filterValue)
		if Request.Form.Count > 0 then
			getFilterValue = Request(filterValue)
		else
			getFilterValue = getSessionObject(filterValue)
		end if
	end function
	
	'************************************************************************************************************
	''@SDESCRIPTION:	Get the sort value
	''@DESCRIPTION:		Get the sort value saved in the SO, or the Request("..sort..")
	''@RETURN:			[string] the sort value if the SO exists, "" if not or the request if postback happend
	'************************************************************************************************************
	public function getSortValue(filterObj, isSortingAllowed)
		if Request.Form.Count > 0 then
			getSortValue = Request("sortValue")
		else
			mySort = getSessionObject("commonsort")
			if isSortingAllowed then
				if len(mySort) <= 1 then
					getSortValue = ""
				elseif not existsSortValue(mySort, filterObj) then
					getSortValue	= Request("sortValue")
				else
					getSortValue	= mySort
				end if
			else
				getSortValue = ""
			end if
		end if
	end function
	
	'************************************************************************************************************
	'* existsSortValue - checks if the sort value field exists
	'************************************************************************************************************
	private function existsSortValue(mysort, filterObj)
		existsSortValue = false
		
		if not isObject(filterObj) then
			exit function
		end if
		
		for each fil in filterObj.items
			if str.startsWith(mySort,fil.fieldName) then
				existsSortValue = true
			end if
		next
	end function
	
	'************************************************************************************************************
	''@SDESCRIPTION:	Get the search value
	''@DESCRIPTION:		Get the search value saved in the SO, or the Request("..search..")
	''@RETURN:			[string] the search value if the SO exists, "" if not or the request if postback happend
	'************************************************************************************************************
	public function getSearchValue()
		if Request.Form.Count > 0 then
			getSearchValue = request.form("fullsearchtext")
		else
			mySearch = getSessionObject("searchvalue")
			if len(mySearch) <= 1 then
				getSearchValue = ""
			else
				getSearchValue = mySearch
			end if
		end if
	end function
	
	'************************************************************************************************************
	''@SDESCRIPTION:	Returns a Collection(Dictionary or Request.Form) with all filters.
	''@DESCRIPTION:		With this function you can get all available filters. If there was a postback resp.
	''					the SO doesn't exist, you will receive the Request.Form Collection. 
	''					Otherwise you will get a Dictionary which will contain all the filters.
	''@RETURN:			[Collection] Dictionary if SO exists, Request.Form if not
	'************************************************************************************************************
	public function getAllFilters()
		if (checkSessionObject) and not (Request.Form.Count > 0) then
			set tableObject = Session("tableObject")
			if getURLCompareResult(tableObject.Item("pageurl")) then
				set dictTemp = Server.CreateObject("Scripting.Dictionary")
				for each key in tableObject.Keys
					if instr(lCase(key), "fltrfield_") then
						dictTemp.Add key, tableObject.Item(key)
					end if
				next
				set getAllFilters = dictTemp
			else
				set getAllFilters = Request.Form
			end if
			set tableObject = nothing
			set dictTemp 	= nothing
		else
			set getAllFilters = Request.Form
		end if
	end function
	
	'************************************************************************************************************
	''@SDESCRIPTION:	Gets a specified value of the Session Object.
	''@DESCRIPTION:		This function checks if the SO exists, and returns the value of the specified key
	''@PARAM:			- objValue: the key for the value you want to get.
	''@RETURN:			[string/int] returns the value if the SO exists or 0 if not
	'************************************************************************************************************
	public function getSessionObject(objValue)
		if checkSessionObject then
			set tableObject = Session("tableObject")
			if getURLCompareResult(tableObject.Item("pageurl")) then
				getSessionObject = tableObject.Item(objValue)
			end if
			set tableObject = nothing
		else
			getSessionObject = empty
		end if
	end function
	
	'************************************************************************************************************
	''@SDESCRIPTION:	Check if the Session Object is available.
	''@DESCRIPTION:		Check if the Session Object is available.
	''@RETURN:			[bool] true, if the Object exists
	'************************************************************************************************************
	public function checkSessionObject()
		if IsObject(Session("tableObject")) then
			checkSessionObject = true
		else
			checkSessionObject = false
		end if
	end function
	
	'************************************************************************************************************
	''@SDESCRIPTION:	Initialize the Session Object.
	''@DESCRIPTION:		If the form was submitted, the function gets all needed values and saves them into the
	''					Session Object.
	'************************************************************************************************************
	public sub InitializeSessionObject()
		if request.form.count > 0 then
			set tableSessionDict = Server.CreateObject("Scripting.Dictionary")
			
			tableSessionDict.Add "pageurl", Request.ServerVariables("URL")
			tableSessionDict.Add "pagenumber", Request("actualPageNumber")
			tableSessionDict.Add "commonsort", Request("sortValue")
			tableSessionDict.Add "searchvalue", Request("fullsearchtext")
			
			for each field in Request.Form
				if (instr(lCase(field), "fltrfield_")) then
					tableSessionDict.Add field, Request(field)
				end if
			next
			
			set Session("tableObject") 	= tableSessionDict
			set tableSessionDict 		= nothing
		end if
	end sub
	
end class
%>