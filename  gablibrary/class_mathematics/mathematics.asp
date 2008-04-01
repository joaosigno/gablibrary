<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		mathematics
'' @CREATOR:		Michael Rebec
'' @CREATEDON:		19.12.2003
'' @CDESCRIPTION:	Various math operations.
''					See a demonstration in the demo-folder of this class.
'' @VERSION:		0.1.1

'**************************************************************************************************************
class mathematics

	'**************************************************************************************************************
	'' @SDESCRIPTION:	checks if a value lies between 2 bounds
	'' @PARAM:			- value [int]: the value which has to be checked
	'' @PARAM:			- lowerBound [int]: the lower bound
	'' @PARAM:			- upperBound [int]: the upper bound
	'' @RETURN:			- [bool] true if the value lies between the bounds
	'**************************************************************************************************************
	public function between(value, lowerBound, upperBound)
		between = cint(value) >= cint(lowerBound) and cint(value) <= cint(upperBound)
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Calculates PI
	'' @RETURN:			- [long] the number PI
	'**************************************************************************************************************
	public function Pi()
		epsilon = 0.00001
		a = 1
		b = sqr(0.5)
		t = 0.25
		x = 1
		do while ((a - b) > epsilon)
		    tempa = a
		    a = (a + b) / 2
		    b = sqr(tempa * b)
		    t = t - x * pow(a - tempa, 2)
		    x = 2 * x
		 	loop
		Pi = pow(a + b, 2.0) / (4 * t)
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Checks if a number is prime
	'' @RETURN:			- [bool] true if number is prime
	'**************************************************************************************************************
	public function IsPrime(ByRef pLngNumber)
	   	IsPrime = false
		if pLngNumber = 2 then
			IsPrime = true
			exit function
		end if
	    if pLngNumber < 2 then exit function
	    if pLngNumber Mod 2 = 0 then exit function
	    lLngSquare = Sqr(pLngNumber)
	    For lLngIndex = 3 To lLngSquare Step 2
		    if pLngNumber Mod lLngIndex = 0 then exit function
	    Next
	   	IsPrime = true
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Checks if a number is prime
	'' @RETURN:			- [bool] true if number is prime
	'**************************************************************************************************************
	'public function IsPrime(ByRef pLngNumber)
	   	'IsPrime = false
		'if pLngNumber = 2 then
			'IsPrime = true
			'exit function
		'end if
	    'if pLngNumber < 2 then exit function
	    'if modulo(pLngNumber, 2) = 0 then exit function
	    'lLngSquare = Sqr(pLngNumber)
	    'For lLngIndex = 3 To lLngSquare Step 2
		    'if modulo(pLngNumber, lLngIndex) = 0 then exit function
	    'Next
	   	'IsPrime = true
	'end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	
	'' @RETURN:			
	'**************************************************************************************************************
	public function modulo(a, b)
		a = cdbl(a)
		b = cdbl(b)
		c = math.floor(a / b)
		
		modulo = a - c * b
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	return the largest value when the dividend is divided by divider
	'' @RETURN:			- [int/long/double]
	'**************************************************************************************************************
	public function Ceil(dividend, divider)
	    if (dividend mod divider) = 0 Then
	   	 	ceil = dividend / divider
	    else
		    ceil = Int(dividend / divider) + 1
	    end if
    end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Floor returns the largest integer less than or equal to the given numeric expression
	'' @PARAM:			[double]
	'' @RETURN:			- [int] number
	'**************************************************************************************************************
	public function Floor(byVal n)
		n 		= cDbl(n)
		iTmp 	= round(n) ' rounds up
		if iTmp > n then iTmp = iTmp - 1 'test rounded value against the non rounded value; if greater sub 1
		Floor = iTmp
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	converts a binary number to integer
	'' @RETURN:			- [int] the converted number
	'**************************************************************************************************************
	public function Bin2Int(byVal Num)
		n = Len(Num) - 1
		a = n
		do while n > -1 
			x = Mid(Num, ((a + 1) - n), 1)
			Bin2Int = lib.IIf((x = 1), Bin2Int + (2 ^ (n)), Bin2Int) 
			n = n - 1 
		loop
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	converts a int number to binary
	'' @RETURN:			- [int/long] the binary number
	'**************************************************************************************************************
	public function Int2Bin(byVal Dec)
		do while Dec > 0
		 Int2Bin = (Dec MOD 2) & Int2Bin
		 Dec = Int(Dec/2)
		loop
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Returns the larger of two specified numbers
	'' @RETURN:			- [int/long/double] the larger of two specified numbers
	'**************************************************************************************************************
	public function Max(number1, number2)
		if number1 > number2 then
			max = number1
		else
			max = number2
		end if
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Returns the smaller of two specified numbers
	'' @RETURN:			- [int/long/double] the smaller of two specified numbers
	'**************************************************************************************************************
	public function Min(number1, number2)
		if number1 > number2 then
			min = number2
		else
			min = number1
		end if
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Returns a specified number raised to the specified power
	'' @RETURN:			- [int/long/double] specified number raised to the specified power
	'**************************************************************************************************************
	public function Pow(byval a, byval b)
		pow = 1
		for i = 1 to b
			pow = pow * a
		next
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Calculates the factorial number
	'' @PARAM:			- [int] the number you want to factorize
	'' @RETURN:			- [int/long/double] the factorial number
	'**************************************************************************************************************
	public function Factorial(byval number)
		if number > 1 then
			Factorial = number * Factorial(number - 1)
		else
			Factorial = 1
		end if
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Calculates the roman numeral for a given number
	'' @PARAM:			- [int] the number you want to format
	'' @RETURN:			- [string] the roman numeral for that number
	'' @CREDIT:			Dougie Lawson/Sock - Planet Source Code
	'**************************************************************************************************************
	public function Roman(byval number)
	    roman_unit = Array("","I","II","III","IV","V","VI","VII","VIII","IX")
	    roman_tens = Array("","X","XX","XXX","XL","L","LX","LXX","LXXX","XC")
	    roman_hund = Array("","C","CC","CCC","CD","D","DC","DCC","DCCC","CM")
	    roman_thou = Array("","M","MM","MMM","MMMM","MMMMM")
	    v = 0 : w = 0 : x = 0 : y = 0
	    v = ((number - (number mod 1000)) / 1000)
	    number = (number mod 1000)
	    w = ((number - (number mod 100)) / 100)
	    number = (number mod 100)
	    x = ((number - (number mod 10)) / 10)
	    y = (number mod 10)
		
	    roman = roman_thou(v) & roman_hund(w) & roman_tens(x) & roman_unit(y)
	end function
	
end class
lib.registerClass("Mathematics")
%>
