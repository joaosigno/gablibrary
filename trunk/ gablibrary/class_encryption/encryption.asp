<!--#include virtual ="/gab_Library/class_encryption/class_md5Algorithm.asp"-->
<!--#include virtual ="/gab_Library/class_encryption/class_sha256Algorithm.asp"-->
<!--#include virtual ="/gab_Library/class_encryption/class_gabLibCrypt.asp"-->
<!--#include virtual ="/gab_Library/class_encryption/class_rijndael.asp"-->
<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		encryption
'' @CREATOR:		Michal Gabrukiewicz - gabru @ grafix.at
'' @CREATEDON:		15.09.2003
'' @CDESCRIPTION:	This class lets you encrypt a string with different algorithms. You can use this class
''					for encryption so you can fast change your algorithm. But you also can use an algorithm-class
''					directly if you want to decrease your bandwith on loading all classes.
''					MD5: Derived from the RSA Data Security, Inc. MD5 Message-Digest Algorithm,
''					as set out in the memo RFC1321.
''					Encryption-code taken from Web Site: http://www.frez.co.uk and modified to a class. Many thanks!
''					AES: needs a key to encode the provided string. Use the encryption.key property to set the key
'' @VERSION:		0.2

'**************************************************************************************************************
const MD5 			= 1
const SHA256 		= 2
const GABLIB	 	= 4
const AES			= 8

class encryption

	private p_algorithm
	private p_key
	private p_bytIn(), p_bytKey()
	
	'Construktor => set the default values
	private sub Class_Initialize()
		p_algorithm = MD5
		key = "AeiI38OsmNq29iOuI7"
	end sub
	
	public property let algorithm(value) ''Sets the algorithm for encryption. Following values accepted: MD5 (default), SHA256, GABLIB, AES
		p_algorithm = value
	end property
	public property get algorithm() ''Returns the algorithm
		algorithm = p_algorithm
	end property
	
	public property let key(value) ''Sets the key for encryption. Only used with AES - if not set, the default key will be used (WARNING: the default key is not secure - use your own for encryption purposes !!!)
		p_key = value
	end property
	public property get key() ''Returns the cipher key
		key = p_key
	end property
	
	'******************************************************************************************
	'* setAESByteArrays - sets the byte arrays p_bytIn and b_bytKey
	'******************************************************************************************
	private sub setAESByteArrays(str)
		lLength = len(str)
		redim p_bytIn(lLength - 1)
		for i = 1 to lLength
			p_bytIn(i - 1) = cByte(AscB(mid(str, i, 1)))
		next
		
		lLength = len(key)
		redim p_bytKey(lLength - 1)
		for i = 1 to lLength
			p_bytKey(i - 1) = cByte(AscB(mid(key, i, 1)))
		next
	end sub
	
	'**********************************************************************************************************************
	'' @SDESCRIPTION:	Encrypts a string with the configured algorithm
	'' @PARAM:			- str [string]: the string for encryption
	'' @RETURN:			- [string] encrypted string
	'**********************************************************************************************************************
	public function encrypt(str)
		select case me.algorithm
			case 1 'MD5
				set encryptObj = new md5Algorithm
				encrypt = encryptObj.MD5(str)
			case 2 'SHA256
				set encryptObj = new sha256Algorithm
				encrypt = encryptObj.SHA256(str)
			case 4 'GabLib
				set encryptObj = new gabLibCrypt
				encrypt = encryptObj.enCrypt(str)
				
			case 8 'Rijndael - AES
				setAESByteArrays(str)
				set encryptObj = new rijndael
				bytOut = encryptObj.encryptData(p_bytIn, p_bytKey)
			    for i = 0 to uBound(bytOut)
			        encrypt = encrypt & right("0" & hex(bytOut(i)), 2)
			    next
		end select
		set encryptObj = nothing
	end function
	
	'**********************************************************************************************************************
	'' @SDESCRIPTION:	Decrypts a string with the configured algorithm (The Cipher-method must be decryptable!)
	'' @PARAM:			- str [string]: the string for decryption
	'' @RETURN:			- [string] decrypted string
	'**********************************************************************************************************************
	public function decrypt(str)
		select case me.algorithm
			case 4 'GabLib
				set encryptObj = new gabLibCrypt
				decrypt = encryptObj.deCrypt(str)
			case 8 'Rijndael - AES
				setAESByteArrays(str)
				set encryptObj = new rijndael
				lLength = len(str)
		        redim bytOut((lLength \ 2) - 1)
		        for i = 1 to lLength step 2
		          bytOut(i \ 2) = cByte("&H" & mid(str, i, 2))
		        Next
				bytClear = encryptObj.decryptData(bytOut, p_bytKey)
				lLength = uBound(bytClear) + 1
				for i = 0 to lLength - 1
					decrypt = decrypt & chr(bytClear(i))
				next
		end select
		set encryptObj = nothing
	end function

end class
%>