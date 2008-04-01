<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		arrayList
'' @CREATOR:		Michael Rebec / Michal Gabrukiewicz
'' @CREATEDON:		2005-01-03 12:36
'' @CDESCRIPTION:	OO representation of the array. Its like the arraylist in .net
'' @VERSION:		0.1

'**************************************************************************************************************
class arrayList

	private p_array		'the arrayList
	private temp		'temporarily used variable (e.g. for swap)
	private tempCount	'temporarily used counter (e.g. ubound(array)) to increase speed
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	private sub class_initialize()
		clear()
	end sub
	
	'**********************************************************************************************************
	'* destructor 
	'**********************************************************************************************************
	private sub class_terminate()
		p_array = null
	end sub
	
	public property get count ''[int] gets the amount of items beginning with 0. -1 = NULL
		count = uBound(p_array)
	end property
	
	public property get items ''[int] gets the number of items in the list. begins with 0 (0 items in the list)
		items = count + 1
	end property
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	Determines if a value exists in the arraylist
	'' @PARAM:			val [variant]: what value
	'' @RETURN:			[bool] true if it contains that value
	'**********************************************************************************************************
	public function contains(val)
		for i = 0 to count
			if cStr(p_array(i)) = cStr(val) then
				contains = true
				exit for
			end if
		next
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	Adds an item to the arrayList
	'' @PARAM:			val [variant] the item you want to add. 
	'**********************************************************************************************************
	public sub add(val)
		redim preserve p_array(count + 1)
		p_array(count) = val
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	Adds an array, or string to the arrayList
	'' @DESCRIPTION:	This procedure allows you to add an array to the arraylist. If you don't have an array
	''					but a string, which have "," as seperators you can add the string-items to the array
	'' @PARAM:			- obj [array/string]: the array you want to add; string seperated by ","
	'**********************************************************************************************************
	public sub addRange(obj)
		if isArray(obj) then
			for i = 0 to uBound(obj)
				add(obj(i))
			next
		else
			addRange(split(cStr(obj), ","))
		end if
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	Clears the array
	'' @DESCRIPTION:	This procedure deallocates the memory used by the arrayList
	'**********************************************************************************************************
	public sub clear()
		p_array = array()
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	Change the value of an existing item
	'' @DESCRIPTION:	You can change the value of an item with this function by setting the index and the 
	''					new value
	'' @PARAM:			- pos [int]: the position in the array
	'' @PARAM:			- val [variant]: the new value
	'**********************************************************************************************************
	public sub setItem(pos, val)
		p_array(pos) = val
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	Get the value of an existing item
	'' @DESCRIPTION:	This function returns the value of an existing item - you have to check outside if the
	''					selected index is in the array-bounds.
	'' @PARAM:			- pos [int]: the position in the array
	'' @RETURN:			[string]: the item at the position
	'**********************************************************************************************************
	public function item(pos)
		item = p_array(pos)
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	Returns the array-index of the first string occurency
	'' @DESCRIPTION:	This function returns the array-index of the first occurency of a given string in the
	''					array. It will only return an index higher then -1 if the whole string is in the array.
	''					Casesensitivity is enabled !!
	'' @PARAM:			- val [string]: the value you are looking for
	'' @RETURN:			[int]: the index of the first occurency in the array; -1 if not
	'**********************************************************************************************************
	public function indexOf(val)
		indexOf = -1
		for i = 0 to count
			if p_array(i) = val then
				indexOf = i
				exit function
			end if
		next
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	Reverse the items in the arrayList
	'' @DESCRIPTION:	This procedure reverses all items in the array, e.g. array(1,2,3,4) will be array(4,3,2,1)
	'**********************************************************************************************************
	public sub reverse()
		tempCount = count
		for i = 0 to cLng(tempCount / 2)
			call swap(i, tempCount - i)
		next
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	Swap two values in the arrayList
	'' @DESCRIPTION:	Swap two arrayList-values identified by two indexes
	'' @PARAM:			- indexFirst [int]: the first position in the arrayList
	'' @PARAM:			- indexSecond [int]: the second position in the arrayList
	'**********************************************************************************************************
	public sub swap(indexFirst, indexSecond)
		temp = p_array(indexFirst)
		p_array(indexFirst) = p_array(indexSecond)
		p_array(indexSecond) = temp
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	Returns the array
	'' @DESCRIPTION:	Returns the arrayList as a new Array
	'' @RETURN:			[array]: the array
	'**********************************************************************************************************
	public function toArray()
		toArray = p_array
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	Returns the array as string, seperated by the "seperator"
	'' @PARAM:			- seperator [string]: the seperator
	'' @RETURN:			[string]: the array formatted as string
	'**********************************************************************************************************
	public function toString(seperator)
		toString = str.arrayToString(p_array, seperator)
	end function

end class
%>