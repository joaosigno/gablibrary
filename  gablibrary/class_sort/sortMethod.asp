<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		SortMethod
'' @CREATOR:		Michael Rebec
'' @CREATEDON:		27.11.2003
'' @CDESCRIPTION:	Sorts Dictionaries, 1 or 2 dimensional Array Objects and it is also possible to sort a
''					table by it`s columns.
'' @VERSION:		0.1

'**************************************************************************************************************
class sortMethod
	
	public ascending
	public int_sortcolumn
	
	'Konstruktor => set the default values
	private sub Class_Initialize()
		ascending = true
	end sub
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	sorts an array with the bubble sort method.
	'' @PARAM:			- arrvar [Array]: the array you want to order
	'**************************************************************************************************************
	public function bubbleSort(byval arrvar)
		myArray = arrvar
		if ascending then
			Do
				bln_allesok = true
				For int_counter = 0 to UBound(myArray) - 1
					if myArray(int_counter) > myArray(int_counter+1) then
						str_help = myArray(int_counter)
						myArray(int_counter) = myArray(int_counter+1)
						myArray(int_counter+1) = str_help
						bln_allesok = false
					end if
				Next
			Loop While bln_allesok = false
		else
			Do
				bln_allesok = true
				For int_counter = 0 to UBound(myArray) - 1
					if myArray(int_counter) < myArray(int_counter+1) then
						str_help = myArray(int_counter)
						myArray(int_counter) = myArray(int_counter+1)
						myArray(int_counter+1) = str_help
						bln_allesok = false
					end if
				Next
			Loop While bln_allesok = false
		end if
		bubbleSort = myArray
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	sorts a 2 dimensional array with the bubble sort method
	'' @PARAM:			- arrvar [Array]: the array you want to order
	'' @PARAM:			- the column you want to order by
	'**************************************************************************************************************
	public function bubbleSort2(byval arrvar)
		ReDim arrhelp(UBound(arrvar,2)) 
		myArray = arrvar
		if ascending then
			Do
				bln_allesok = true
				For int_counter = 0 to UBound(myArray,1) - 2
					if myArray(int_counter,int_sortcolumn) > myArray(int_counter+1,int_sortcolumn) then
						For int_innercounter = 0 to UBound(arrhelp,1) - 1
							arrhelp(int_innercounter) = myArray(int_counter, int_innercounter)
							myArray(int_counter,int_innercounter) = myArray(int_counter+1,int_innercounter)
							myArray(int_counter+1,int_innercounter) = arrhelp(int_innercounter)
						Next
						bln_allesok = false
					end if
				Next
			Loop While bln_allesok = false
		else
			Do
				bln_allesok = true
				For int_counter = 0 to UBound(myArray,1) - 2
					if myArray(int_counter,int_sortcolumn) < myArray(int_counter+1,int_sortcolumn) then
						For int_innercounter = 0 to UBound(arrhelp,1) - 1
							arrhelp(int_innercounter) = myArray(int_counter, int_innercounter)
							myArray(int_counter,int_innercounter) = myArray(int_counter+1,int_innercounter)
							myArray(int_counter+1,int_innercounter) = arrhelp(int_innercounter)
						Next
						bln_allesok = false
					end if
				Next
			Loop While bln_allesok = false
		end if
		bubbleSort2 = myArray
	end function
	
		'**************************************************************************************************************
	'' @SDESCRIPTION:	this is a special sort alogrithm only for tables. It can be used to order tables 
	''					by their columns.
	''					For example, if you have a presorted array and you want to insert these values 
	''					into a table with "n" columns, this function will return you an ordered array where the
	''					values are sorted by the table columns
	'' @PARAM:			- arObj [Array]: the array you want to order
	'**************************************************************************************************************
	public function columnSort(byval arObj)
		if int_sortcolumn <= 1 then				' if you have max 1 int_sortcolumn the array won't be changed
			tableSort = arObj
			exit function
		end if
		anz = ubound(arObj)						' get the number of elements
		redim myArray(anz)						' create a new Array with the same number of elements
		cur = 0									' contains the current arrayindex
		counter = 0								' counts the columns
		x = (anz+1) Mod int_sortcolumn			' how much columns in the last row
		for i = 0 to anz
			myArray(i) = arObj(cur)
			if counter = int_sortcolumn then counter = 0	' reset counter if columns reached
			counter = counter + 1
			if counter <= x then							' controls the step size
				step = Fix((anz+1)/int_sortcolumn)+1
			else
				step = Fix((anz+1)/int_sortcolumn)
			end if
			if (step + cur) > anz then			' get the next arrayindex
				cur = (step + cur) - (anz)
			else
				cur = cur + step
			end if
		next
		columnSort = myArray
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	this is a special sort alogrithm for arrays, if you want to
	''					randomize the values (with repeating or without) of your array
	'' @PARAM:			- arObj [Array]: the array you want to shuffle
	'**************************************************************************************************************
	public function shuffleSort(byVal arObj)
		myArray = arObj
		max = UBound(arObj)
		randomize
		for i = 1 to max
			do
				randomValue = round(rnd * (max - 1) + 1)
				if arObj(randomValue) <> ":::" then
					myArray(i) = arObj(randomValue)
					arObj(randomValue) = ":::"
					exit do
				end if
			loop
		next
		
		postRnd = round(rnd * (max - 1) + 1) 'randomize the first element
		tmp = myArray(0)
		myArray(0) = myArray(postRnd)
		myArray(postRnd) = tmp
		
		shuffleSort = myArray
	end function
end class
%>