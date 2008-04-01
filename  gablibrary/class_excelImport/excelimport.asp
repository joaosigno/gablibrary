<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		Excelimport
'' @CREATOR:		Michael Rebec
'' @CREATEDON:		31.10.2003
'' @CDESCRIPTION:	Imports Data from an Excel file, e.g. "SELECT * FROM [Tabelle$A:F]"
'' @VERSION:		0.21

'**************************************************************************************************************
class excelimport

	private objConn						' Connection Object
	private objRS						' Recordset Object
	private strDriver					' Excel Driver
	private strDatabase					' Excel Database
	
	public sqlQuery						'' SQL Command, e.g. "SELECT * FROM [Tabelle$A:F]"
	public fileName						'' The Path and Filename of the Excel File
	
	'Konstruktor => set the default values
	private sub Class_Initialize()
		sqlQuery 			= empty
		fileName			= empty
	end sub
	
	public property get columnCount '' [int] Returns the number of Columns
		columnCount = objRS.fields.count
	end property
	
	'**************************************************************************************************************
	'' @DESCRIPTION: 	Initialize's the connection and returns the Recordset Object 
	''					which contains the data from the Excelfile
	'' @RETURN:			[recordset-object] Recordset Object with the data
	'**************************************************************************************************************
	public function getRecordset()
		if not sqlQuery = empty then
			'call initDriver()	- from version 0.1 -> now sConn
			set getRecordset = createConnection(sqlQuery)
		end if
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION: 	closes the Connection to the excelsheet and kills the recordset-object
	'**************************************************************************************************************
	public sub closeConnection()
		objRS.Close()
		Set objRS = Nothing
		objConn.Close()
		Set objConn = Nothing
	end sub
	
	'**************************************************************************************************************
	' initialize the Database Excel Driver
	'**************************************************************************************************************
	private sub initDriver()
		strDriver = "DRIVER=Microsoft Excel-Treiber (*.xls);"
		strDatabase = "DBQ=" & Server.MapPath(fileName) & ";"
	end sub
	
	'**************************************************************************************************************
	' creates a Excel DB connection and returns a Recordset Object 
	'**************************************************************************************************************
	private function createConnection(sqlCommand)
		On Error Resume Next
			'Create ADODB Connection
			Set objConn = Server.CreateObject("ADODB.Connection")
			'objConn.Open strDriver & strDatabase
			sConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(fileName) & ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"""
			objConn.Open sConn
			
			'Create ADODB Recordset
			Set objRS = Server.CreateObject("ADODB.Recordset")
			objRS.Open sqlCommand, objConn
			
			if Err.Number <> 0 then
				call errorHandling(Err)
			else
				'Return ADODB Recordset
				Set createConnection = objRS
			end if
	end function
	
	'**************************************************************************************************************
	' handles the errors
	'**************************************************************************************************************
	private sub errorHandling(errorMsg)
		const wrongSheet = -2147217865	' If you entered a xls file with the wrong format
		response.write("<BR><BR><div align=""center""><span class=""nosuccess"">")
			Select Case errorMsg.Number
				Case wrongSheet: 
					response.write("An error occured, please submit the correct file")
				Case Else
					response.write("An error occured: " & errorMsg.Number & " " & errorMsg.Description)
			End Select
		response.write("</span><BR><BR><BR><input class=button type=""Button"" value=""Back"" onclick=""javascript:history.back();"">")
		response.write("</div>")
		response.end
	end sub

end class
lib.registerClass("ExcelImport")
%>