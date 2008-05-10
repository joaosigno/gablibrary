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
'' @FRIENDOF:		DrawTable

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
		if lib.page.isPostback() then
			getAbsolutePage = lib.RF("actualPageNumber")
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
		if lib.page.isPostback() then
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
		if lib.page.isPostback() then
			getSortValue = Request("sortValue")
		else
			mySort = getSessionObject("commonsort")
			if isSortingAllowed then
				if len(mySort) <= 1 then
					getSortValue = ""
				elseif not existsSortValue(mySort, filterObj) then
					getSortValue = Request("sortValue")
				else
					getSortValue = mySort
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
		
		if not isObject(filterObj) then exit function
		
		for each fil in filterObj.items
			if str.startsWith(mySort,fil.fieldName) then existsSortValue = true
		next
	end function
	
	'************************************************************************************************************
	''@SDESCRIPTION:	Get the search value
	''@DESCRIPTION:		Get the search value saved in the SO, or the Request("..search..")
	''@RETURN:			[string] the search value if the SO exists, "" if not or the request if postback happend
	'************************************************************************************************************
	public function getSearchValue()
		if lib.page.isPostback() then
			getSearchValue = lib.RF("fullsearchtext")
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
		if (checkSessionObject) and not lib.page.isPostback() then
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
				set getAllFilters = lib.form
			end if
			set tableObject = nothing
			set dictTemp 	= nothing
		else
			set getAllFilters = lib.form
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
			if getURLCompareResult(tableObject.Item("pageurl")) then getSessionObject = tableObject.Item(objValue)
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
		set tableSessionDict = lib.newDict(empty)
		searchValue = request("fullsearchtext")
		
		'we allow to set the settings if page was posted or if it was not posted but a search value is
		'given through the querystring. with this we can achieve to call a table with a given search termn
		if lib.page.isPostback() or (not lib.page.isPostback() and searchValue <> "") then
			tableSessionDict.Add "pageurl", request.ServerVariables("URL")
			tableSessionDict.Add "pagenumber", request("actualPageNumber")
			tableSessionDict.Add "commonsort", request("sortValue")
			tableSessionDict.Add "searchvalue", searchValue
			
			for each field in lib.form
				if (instr(lCase(field), "fltrfield_")) then tableSessionDict.Add field, lib.RF(field)
			next
			set session("tableObject") = tableSessionDict
			set tableSessionDict = nothing
		end if
	end sub
	
end class
%>