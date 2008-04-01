<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		gabLibCrypt
'' @CREATOR:		Michal Gabrukiewicz - gabru @ grafix.at
'' @CREATEDON:		16.09.2003
'' @CDESCRIPTION:	Thats a custom crypt for GabLib called GabLibCrypt :) It crypts strings using only letters
''					and numbers. (HEXcode). You can Crypt a string and decrypt back. 
''					Attention: this is a very simple algorithm, so don't use it for sensitive data.
'' @VERSION:		0.2

'**************************************************************************************************************
class gabLibCrypt
	private idSession
	
	public useSessionID
	
	'Konstruktor => set the default values
	private sub Class_Initialize()
		idSession = 0
		useSessionID = false
	end sub
	
	'**********************************************************************************************************************
	'' @SDESCRIPTION:	Encrypts a string with the GabLibCrypt algorithm
	'' @PARAM:			- str [string]: the string for encryption
	'' @RETURN:			- [string] encrypted string
	'**********************************************************************************************************************
	public function deCrypt(byval str)
		tmp = empty
		if right(str, 2) = "FF" then
			useSessionID = true
		else
			useSessionID = false
		end if
		str = trimend(str)
		id_step = 1
		for i = 1 to len(str) step 2
			if useSessionID then
				tmp = tmp & chr(clng("&H" & mid(str, i, 2)) XOR getSessionID(id_step))
				id_step = id_step + 1
			else
				tmp = tmp & chr(clng("&H" & mid(str, i, 2)) XOR 30)
			end if
		next
		deCrypt = tmp
	end function
	
	'**********************************************************************************************************************
	'' @SDESCRIPTION:	Encrypts a string with the GabLibCrypt algorithm
	'' @PARAM:			- str [string]: the string for encryption
	'' @RETURN:			- [string] encrypted string
	'**********************************************************************************************************************
	public function enCrypt(str)
		tmp = empty
		for i = 1 to len(str)
			if useSessionID then
				hx = cstr(hex(asc(mid(str, i, 1)) XOR getSessionID(i)))
				if len(hx) = 1 then
					hx = "0" & hx
				end if
				tmp = tmp & hx
			else
				tmp = tmp & hex(asc(mid(str, i, 1)) XOR 30)
			end if
		next
		if useSessionID then
			tmp = tmp & "FF"
		else
			tmp = tmp & "00"
		end if
		enCrypt = tmp
	end function
	
	'**********************************************************************************************************************
	'* getSessionID
	'**********************************************************************************************************************
	private function getSessionID(id_step)
		tmp = Session.SessionID
		aaa = len(tmp)
		getSessionID = CInt(Mid(tmp, (id_step mod aaa)+1, 1)) * 10
	end function
	
	'**********************************************************************************************************************
	'* trimEnd 
	'**********************************************************************************************************************
	private function trimEnd(str)
		if str <> empty then
			trimEnd = Left(str, len(str)-2)
		end if
	end function

end class
%>