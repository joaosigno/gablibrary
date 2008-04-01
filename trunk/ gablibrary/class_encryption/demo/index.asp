<!--#include virtual="/gab_Library/class_page/generatepage.asp"-->
<!--#include virtual="/gab_Library/class_encryption/encryption.asp"-->
<%
set page = new generatePage
with page
	.onlyWebDev = true
	.draw()
end with
set page = nothing

'******************************************************************************************
'* main - we start . DONT FORGET TO KILL ALL OBJECTS!! 
'******************************************************************************************
sub main()
	shouldBe = "b87c88f85edaee124b3995b4555579a7"
	stringToCrypt = "GabLib"
	
	set encrypt = new encryption
	
	str.writeln("<br><br>Asynchronious Hashing Methods (deCrypt not possible):<hr>")
	str.writeln "Encrypted with MD5: "
	encrypt.algorithm = MD5
	str.writeln encrypt.encrypt(stringToCrypt) 'we crypt the String
	
	str.writeln "<BR>"
	
	str.writeln "Encrypted with SHA256: "
	encrypt.algorithm = SHA256
	str.writeln encrypt.encrypt(stringToCrypt) 'we crypt the String
	
	str.writeln "<br><br>Synchronious Cipher Methods:<hr>"
	
	str.writeln "Encrypted with GABLIB: "
	encrypt.algorithm = GABLIB
	str.writeln encrypt.encrypt(stringToCrypt) 'we crypt the String
	str.writeln "<BR>Decrypted with GABLIB: "
	str.writeln encrypt.decrypt(encrypt.encrypt(stringToCrypt)) 'we decrypt the String
	
	str.writeln "<br><br>"
	
	str.writeln "Ecrypted with Rijndael - AES: "
	encrypt.algorithm = AES
	enc = encrypt.encrypt(stringToCrypt)
	str.writeln enc
	str.writeln "<BR>Decrypted with Rijndael - AES: "
	str.writeln encrypt.decrypt(enc)
	str.writeln("<br><em>For a more detailed demo look at rijndael.asp in the demo folder !</em>")
	
	set encrypt = nothing
end sub
%>