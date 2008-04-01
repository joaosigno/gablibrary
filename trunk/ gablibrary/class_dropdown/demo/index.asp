<!--#include virtual="/gab_Library/class_page/generatepage.asp"-->
<!--#include virtual="/gab_Library/class_dropdown/dropdown.asp"-->
<%
'******************************************************************************************
'* Creator: 	Michal Gabrukiewicz
'* Created on: 	2005-02-01 16:56
'* Description: demonstration of dropdown-control
'******************************************************************************************

set page = new GeneratePage
with page
	.DBConnection = true
	.onlyWebDev = true
	.draw()
end with
set page = nothing

'******************************************************************************************
'* main - we start . DONT FORGET TO KILL ALL OBJECTS!! 
'******************************************************************************************
sub main()
	set dict = server.createObject("scripting.dictionary")
	dict.add lib.getUniqueID(), "michal"
	dict.add lib.getUniqueID(), "michael"
	dict.add lib.getUniqueID(), "axel"
	dict.add lib.getUniqueID(), "Axel Schweis"
	dict.add lib.getUniqueID(), "Franz Maier"
	
	set aDD = new Dropdown
	with aDD
		set .datasource = dict
		.selectedValue = 2
		.name = "myDropdown"
		.id = "aDD"
		.dataValuefield = "id_user"
		.cssClass = "css"
		.commonFieldText = "--- select ---"
		.size = 20
		.commonFieldValue = "220"
		.onItemCreated = "onDDItemCreated"
		.autoDrawItems = false
	end with
	
	set bDD = new Dropdown
	with bDD
		set .datasource = dict
		.selectedValue = 2
		.name = "myDropdown2"
		.size = 20
		.id = "bDD"
		.dataValuefield = "id_user"
		.onItemCreated = "onDDItemCreated2"
		set connector = .getNewConnector(aDD)
	end with
	
	set connector.page = page
	connector.draw 10, "300"
	
	set aDD = nothing
	set bDD = nothing
	
	str.writeln("<br><br><form name='frm' action='index.asp?' method='post'>")
	'Dropdown example with adding new items
	set DD = new Dropdown
	with DD
		set .datasource = dict
		.name = "adding"
		.enableAdding = true
		.attributes = "onChange='frm.submit()'"
		.selectedValue = lib.RF("adding")
		.style = "width:300px;"
		.draw()
	end with
	str.writeln("</form>")
	set dict = nothing
end sub

'******************************************************************************************
'* onDDItemCreated 
'******************************************************************************************
sub onDDItemCreated2(eventArgs)
	with eventArgs
		if .index = 2 then
			.text = "<img src=/images/icon_pdf.gif align=absmiddle>&nbsp;text"
		end if
	end with
end sub

'******************************************************************************************
'* onDDItemCreated 
'* demonstrates how to add an item during runtime
'******************************************************************************************
sub onDDItemCreated(eventArgs)
	with eventArgs
		if .index = 1 then
			.style = "background-color:red;color:white"
			.text = "haxn"
			.title = "test"
			set it = eventArgs.dropdown.getNewItem("check", "check")
			eventArgs.draw()
			it.draw()
		else
			eventArgs.draw()
		end if
	end with
end sub
%>