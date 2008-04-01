<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		pageable
'' @CREATOR:		Michael Rebec, Michal Gabrukiewicz
'' @CREATEDON:		2005-01-02 12:15
'' @CDESCRIPTION:	An object to easily implement paging for recordsets, arrays, etc.
''					paging works like e.g. in google. not all pages are being displayed. just as much as you want
'' @VERSION:		0.21

'**************************************************************************************************************
class pageable

	private pagesArray		
	private p_pageCount		
	private p_currentPage	
	private p_numberOfPages	
	
	public recordCount		''[int] the number of all records; if you don´t know the max. number of records you can
							''set the pageCount directly (recordCount has to be 0 (default))
	public recordsPerPage	''[int] how many records do you display per page?
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	private sub class_initialize()
		p_currentPage 	= 1
		recordCount		= 0
		recordsPerPage	= 0
		p_numberOfPages	= 1
		p_pageCount		= 0
		pagesArray		= array()
	end sub
	
	public property let currentPage(value) ''[int] sets the current-page
		p_currentPage = cInt(value)
	end property
	public property get currentPage ''[int] gets the current-page
		currentPage = p_currentPage
	end property
	
	public property let numberOfPages(value) ''[int] sets the number of displayed page-numbers. even numbers will be round up!
        value = cInt(value)
        if value >= 2 then 
            p_numberOfPages = int(value / 2)
        end if
    end property
	public property get numberOfPages ''[int] gets the number of displayed pages
		numberOfPages = (p_numberOfPages * 2) + 1
    end property
	
	public property get pageCount ''[int] returns the pageCount
		pageCount = p_pageCount
		if recordsPerPage > 0 and recordCount > 0 then pageCount = ceil(recordCount, recordsPerPage)
	end property
	public property let pageCount(value) ''[int] sets the pagecount. you can set it by your own then you dont need to provide recordcount and recordsperpage
		p_pageCount = cInt(value)
	end property
	
	public property get pages ''[array] array including all your pages. available after executing perform()
		pages = pagesArray
	end property
	
	public property get lastPage ''[int] gets the number of the last page
		lastPage = pageCount
	end property
	
	public property get firstPage ''[int] gets the number of the first page
		firstPage = 1
	end property
	
	public property get dataStartPosition ''[int] gets the start position for your data.
		if isOnFirstPage() then
			dataStartPosition = 1
		else
			dataStartPosition = ((currentPage - 1) * recordsPerPage) + 1
		end if
	end property
	
	public property get dataEndPosition ''[int] gets the end position for your data.
		if isOnLastPage() then
			dataEndPosition = recordCount
		else
			dataEndPosition = dataStartPosition + (recordsPerPage - 1)
		end if
	end property
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	performs the paging algorithm. you need to execute this at least once.
	'' @DESCRIPTION:	After executing you can access all the properties for your input and
	''					eg. draw a paging-bar
	'**********************************************************************************************************
	public sub perform()
		adjustCurrentPage()
		if pageCount > 1 then
			
			startIDX = 1
			if currentPage > p_numberOfPages then
				cnt = 0
				if pageCount - currentPage <= p_numberOfPages then cnt = p_numberOfPages - abs(pageCount - currentPage)
				tmp = currentPage - p_numberOfPages - cnt
				if tmp >= 1 then startIDX = tmp
			end if
			
			endIDX = pageCount
			if currentPage + p_numberOfPages  <= pageCount then
				cnt = 0
				if currentPage <= p_numberOfPages then cnt = abs(currentPage - p_numberOfPages) + 1
				tmp = currentPage + p_numberOfPages + cnt
				if tmp < pageCount then endIDX = tmp
			end if
			
			arrayFields = empty
			for i = startIDX to endIDX
				if i <> startIDX then arrayFields = arrayFields & ","
				arrayFields = arrayFields & i
			next
			
			if arrayFields <> empty then pagesArray = split(arrayFields, ",")
		else
			pagesArray = array(1)
		end if
	end sub
	
	'**********************************************************************************************************
	'' adjustCurrentPage 
	'' checks if the currentpage is possible or not. Thanks to SOLO for this idea
	'**********************************************************************************************************
	private function adjustCurrentPage()
		if currentPage > pageCount then
			currentPage = pageCount
		elseif currentPage < firstPage then
			currentPage = firstPage
		end if
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	returns the pagenumber for the wanted position in our pages
	'' @RETURN:			[int]
	'**********************************************************************************************************
	public function page(position)
		page = cInt(pagesArray(position))
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	returns true if there is an available next block with the amount of pages
	'' @RETURN:			[bool]
	'**********************************************************************************************************
	public function hasNextBlock()
		hasNextBlock = (cInt(pagesArray(uBound(pagesArray))) < pageCount)
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	returns true if there is an available previous block with the amount of pages
	'' @RETURN:			[bool]
	'**********************************************************************************************************
	public function hasPreviousBlock()
		hasPreviousBlock = (cInt(pagesArray(lBound(pagesArray))) > 1)
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	true if there is at least one next page
	'' @RETURN:			[bool]
	'**********************************************************************************************************
	public function hasNextPage()
		hasNextPage = (currentPage < pageCount)
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	true if the is at least one previous page
	'' @RETURN:			[bool]
	'**********************************************************************************************************
	public function hasPreviousPage()
		hasPreviousPage = (currentPage > 1)
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	true if the currentpage is the firstpage
	'' @RETURN:			[bool]
	'**********************************************************************************************************
	public function isOnFirstPage()
		isOnFirstPage = (currentPage = firstPage)
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	true if the currentpage is the lastpage
	'' @RETURN:			[bool]
	'**********************************************************************************************************
	public function isOnLastPage()
		isOnLastPage = (currentPage = lastPage)
	end function
	
	'**********************************************************************************************************
	'* ceil 
	'**********************************************************************************************************
	private function ceil(dividend, divider)
	    if (dividend mod divider) = 0 Then
	   	 	ceil = dividend / divider
	    else
		    ceil = Int(dividend / divider) + 1
	    end if
    end function

end class
lib.registerClass("Pageable")
%>