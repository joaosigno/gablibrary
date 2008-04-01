<!--#include file="sortMethod.asp"-->
<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		Sort
'' @CREATOR:		Michael Rebec
'' @CREATEDON:		27.11.2003
'' @CDESCRIPTION:	Sorts Dictionaries, 1 or 2 dimensional Array Objects and it is also possible to sort a
''					table by it`s columns.
'' @VERSION:		0.1

'**************************************************************************************************************
const TYPE_ARRAY 	= 1
const TYPE_ARRAY2	= 2
const TYPE_DICT		= 4
const TYPE_COLUMN	= 8
const TYPE_SHUFFLE	= 16
const SORT_BUBBLE	= 32
const SORT_ASC		= true
const SORT_DESC		= false

class sort

	private p_sortType
	private p_source
	private p_sourceType
	private p_sortMethod
	private p_column
	
	'Konstruktor => set the default values
	private sub Class_Initialize()
		p_sortType 				= SORT_BUBBLE
		p_sourceType 			= 0
		p_source				= empty
		set p_sortMethod 		= new sortMethod
		p_sortMethod.ascending	= true
	end sub
	
	'Destruktor
	private sub Class_Terminate()
		set p_sortMethod = nothing
	end sub
	
	'************
	public property get sortType() ''Gets the SortType
		sortType = p_sortType
	end property
	public property let sortType(value) ''Sets the SortType. Needs a value 1,2,4,8,etc.
		Select Case UCase(value)
			Case SORT_BUBBLE
				p_sortType = 0
		End Select
	end property
	
	'************
	public property get source() ''Returns the source
		source = p_source
	end property
	public property let source(value) ''Sets the source to sort
		p_source = value
	end property
	
	'************
	public property get column()
		column = p_column
	end property
	public property let column(value)
		column = value
		p_sortMethod.int_sortcolumn = value
	end property
	
	'************
	public property let sortOrder(value) ''[bool] - Ascending?
		p_sortOrder = value
		p_sortMethod.ascending = value
	end property
	
	'************
	public property get sourceType()
		sortType = p_sortType
	end property
	public property let sourceType(value)
		p_sourceType = UCase(value)
	end property
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	performs the sort
	'**************************************************************************************************************
	public function sort()
		select Case CInt(p_sourceType)
			case TYPE_ARRAY:
				sort = p_sortMethod.bubbleSort(p_source)
				exit function
			case TYPE_ARRAY2:
				sort = p_sortMethod.bubbleSort2(p_source)
				exit function
			case TYPE_COLUMN:
				sort = p_sortMethod.columnSort(p_source)
				exit function
			case TYPE_SHUFFLE:
				sort = p_sortMethod.shuffleSort(p_source)
				exit function
		end Select
		sort = p_source
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Orders a dictionary object by one or two dictionary attributes. For example, if you have
	''					a dictionary with the attributes .name, .title, .color you can say you want to order the
	''					dict by "title, name". The function will return a sorted dictionary (ordered by the key).
	''					Note: Dictionary key has to start with 0 !! and have to be increased by 1 !!!
	'' @PARAM:			- objDict [Scripting.Dictionary]: the dictionary you want to sort
	'' @PARAM:			- orderBez [string]: the attributes you want to order by; separeted by ","
	'' @PARAM:			- sortMode [string]: sorting Method
	'' @RETURN:			- [Scripting.Dictionary]
	'**************************************************************************************************************
	public function SortDict(byval objDict, orderBez, sortMode)
		set tmpDict = Server.createObject("Scripting.Dictionary")
		dim nCount, strKey, iTemp, jTemp, strTemp
		if inStr(orderBez, ",") > 1 then
			orderObj = split(orderBez, ",")
			nCount = 0
			' Redim the array to the number of keys we need
			redim TempArray(objDict.Count - 1)
			for each strKey in objDict.Items
				sortString = ""
				for i = 0 to UBound(orderObj)
					if i = UBound(orderObj) then
						sortString = sortString & "strKey." & trim(orderObj(i))
					else
						sortString = sortString & "strKey." & trim(orderObj(i)) & " & "
					end if
				next
				execute("test=" & sortString)
				tempArray(nCount) = test
				nCount = nCount + 1
			next
			
			sortedArray = p_sortMethod.bubbleSort(tempArray)
			
			for i = 0 to UBound(sortedArray)
				for j = 0 to UBound(sortedArray)
					tmp = "objDict(i)." & orderObj(UBound(orderObj))
					execute("first=" & tmp)
					tmp = "len(objDict(i)." & orderObj(0) & ")"
					execute("secon=" & tmp)
					if strComp(first, mid(sortedArray(j),secon + 1)) = 0 then
						tmpDict.Add j, objDict(i)
					end if
				next
			next
			
			set SortDict = tmpDict
		end if
		set tmpDict = nothing
	end function

end class
lib.registerClass("Sort")
%>