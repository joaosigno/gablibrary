<!--#include virtual="/gab_Library/class_page/generatepage.asp"-->
<!--#include virtual="/gab_Library/class_validateable/validateable.asp"-->
<!--#include virtual="/gab_Library/class_form/form.asp"-->
<%
'******************************************************************************************
'* Creator: 	David Rankin
'* Created on: 	2006-10-23 11:55
'* Description: Demo for the form class
'******************************************************************************************

set aForm = new Form
set page = new generatepage
with page
	.onlyWebDev = true
	.loadToolTips = true
	.draw()
end with
set page = nothing
set aForm = nothing

'******************************************************************************************
'* main 
'******************************************************************************************
sub main()
	if page.isPostBack() then 
		Name = lib.RF("name")
		if isValid(name) and aForm.validator.isValid() Then
			lib.execJS("alert('Valid!')")
		end if
	end if
	content()
end sub

'******************************************************************************************
'* function 
'******************************************************************************************
function isValid(name)
	isValid = true
	if trim(name & "") = "" then
		aForm.validator.addInvalidField "Name", "Please enter a valid name"
		IsValid = false
	end if
end function

'******************************************************************************************
'* content 
'******************************************************************************************
sub content() %>
	
	
	<form name="frm" method="post" action="<%= aForm.action %>" class="form">
	<h1>Name</h1>
	<fieldset>
		<legend><div>The Name </div></legend>
		<div class="content">
		<table>
			<tr>
				<td class="label">Name:</td>
				<td>
					<input type="Text" name="name" value="" size="40"> 
					<% aForm.drawError "name" %>
				</td>
			</tr>
		</table>
	</fieldset>
	
	<div class="endline">
		<input class="button" type="Submit" name="save" value="Save">
		<% aForm.drawPrintButton() %>
		<% aForm.drawCancelButton("/index.asp") %>
	</div>
	
	</form>

<% end sub %>