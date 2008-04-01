<!--#include virtual="/gab_Library/class_page/generatepage.asp"-->
<!--#include virtual="/jobApply/lib.asp"-->
<!--#include virtual="/gab_Library/class_drawtable/drawtable.asp"-->
<!--#include virtual="/gab_Library/class_tabs/tabs.asp"-->
<%
'******************************************************************************************
'* Creator: 	Michal Gabrukiewicz, David Rankin
'* Created on: 	2006-07-00 15:12
'* Description: Lists all jobs
'* Input:		-
'******************************************************************************************

set page = new GeneratePage
set applicantTable = new Drawtable
set tabStrip = new tabs
with page
	.onlyWebDev = true
	.draw()
end with
set tabStrip = nothing
set page = nothing
set applicantTable = nothing

'******************************************************************************************
'* main 
'******************************************************************************************
sub main()
	
	
	tabStrip.defaultStylesheet = false
	
	set aTab = new tab
	aTab.caption = "Current Jobs"
	aTab.Procedure = "SetDrawTableSQL"
	tabStrip.AddTab(aTab)
	tabStrip.default aTab
	set aTab = nothing
	
	set aTab = new tab
	aTab.caption = "Future Jobs"
	aTab.Procedure = "SetDrawTableSQL"
	tabStrip.AddTab(aTab)
	set aTab = nothing
	
	set aTab = new tab
	aTab.caption = "Expired Jobs"
	aTab.Procedure = "SetDrawTableSQL"
	tabStrip.AddTab(aTab)
	set aTab = nothing
	
	set aTab = new tab
	aTab.caption = "All Jobs"
	aTab.Procedure = "SetDrawTableSQL"
	tabStrip.AddTab(aTab)
	set aTab = nothing
	
	SetDrawTableSQL()
	
	with applicantTable
		.fullsearch = true
		.fieldLinkID = "id"
		.showFilterBar = true
		.allowAdding = true
		.paging = true
		.recsPerPage = 20
		.excelExport = true
		.commonSort = "createdOn DESC"
		.addurl = "addJob.asp"
		.addTxt = "Add new job"
		.height = 350
		
		'Columns
		set aColumn = new columnCommon
		aColumn.fieldName = "title_de"
		aColumn.displayString = "German Title"
		aColumn.displayFunction = "loadFrameset"
		.addColumn(aColumn)
		set aColumn = nothing
		
		set aF = new columnFilter
		with aF
			.fieldName = "title_en"
		end with
		.addNewFilter(aF)
		set aF = new columnFilter
		with aF
			.fieldName = "title_de"
		end with
		.addNewFilter(aF)
		set aF = new columnFilter
		with aF
			.fieldName = "startdate"
		end with
		.addNewFilter(aF)
		set aF = new columnFilter
		with aF
			.fieldName = "enddate"
		end with
		.addNewFilter(aF)
		set aF = new columnFilter
		with aF
			.fieldName = "createdOn"
		end with
		.addNewFilter(aF)
		set aF = new columnFilter
		with aF
			.fieldName = "Applicants"
		end with
		.addNewFilter(aF)
		set aF = new columnFilter
		with aF
			.fieldName = "name_en"
			.showDropdownSQL = "SELECT * FROM location"
			.primaryKey = "id"
			.displayField = "name_en"
			.fieldToMatch = "fk_location"
		end with
		.addNewFilter(aF)
		
		set aColumn = new columnCommon
		aColumn.fieldName = "title_en"
		aColumn.displayString = "English Title"
		.addColumn(aColumn)
		set aColumn = nothing
		
		set aColumn = new columnCommon
		aColumn.fieldName = "name_en"
		aColumn.displayString = "Location"
		.addColumn(aColumn)
		set aColumn = nothing
		
		set aColumn = new columnCommon
		aColumn.fieldName = "startdate"
		aColumn.displayString = "Start Date"
		.addColumn(aColumn)
		set aColumn = nothing
		
		set aColumn = new columnCommon
		aColumn.fieldName = "enddate"
		aColumn.displayString = "End Date"
		.addColumn(aColumn)
		set aColumn = nothing
		
		set aColumn = new columnCommon
		aColumn.fieldName = "createdOn"
		aColumn.displayString = "Created on"
		.addColumn(aColumn)
		set aColumn = nothing
		
		set aColumn = new columnCommon
		aColumn.fieldName = "Applicants"
		aColumn.isNumber = true
		aColumn.displayString = "Applicants"
		.addColumn(aColumn)
		set aColumn = nothing
		
	end with
	
	content()
end sub

'******************************************************************************************
'* loadFrameset
'******************************************************************************************
function loadFrameset(x)
	applicantTable.addurl = "jobFrameset.asp"
	loadFrameset = x
end function

'******************************************************************************************
'* GetSelectedTab
'******************************************************************************************
function GetSelectedTab()
	GetSelectedTab = 0
	if isNumeric(request.querystring("activeTab")) then
		GetSelectedTab = cint(request.querystring("activeTab"))
	end if
end function

'******************************************************************************************
'* SetDrawTableSQL
'******************************************************************************************
function SetDrawTableSQL()
	SelectedTab = GetSelectedTab()
	Select Case SelectedTab
		Case 1
			applicantTable.SQLQuery = "SELECT * FROM ApplicantViewCurrent WHERE 1 = 1"
		Case 2
			applicantTable.SQLQuery = "SELECT * FROM ApplicantViewFuture WHERE 1 = 1"
		Case 3
			applicantTable.SQLQuery = "SELECT * FROM ApplicantViewExpired WHERE 1 = 1"
		Case 4
			applicantTable.SQLQuery = "SELECT * FROM ApplicantViewAll WHERE 1 = 1"
		Case Else
			applicantTable.SQLQuery = "SELECT * FROM ApplicantViewCurrent WHERE 1 = 1"
	End Select
end function

'******************************************************************************************
'* content 
'******************************************************************************************
sub content() %>
	<%
	tabStrip.Draw()
	applicantTable.draw()
	%>
<% end sub %>