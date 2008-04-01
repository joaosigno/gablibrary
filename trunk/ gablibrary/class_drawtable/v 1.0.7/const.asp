<%
'******************************************************************************************************************************
' CONSTANTS FOR OUR TABLE TEMPLATE 
'******************************************************************************************************************************
const DEFAULTDATABASE				= "oracle"																		'whats the defaultdatabase
const SELECTEDCOLOR					= "#CCCCCC"																		'The color is we select a record-line
const HEADERPERROWS					= 30																			'show tablerowheaders every X line?
const RECORDSPERPAGE				= 50																			'how many records should be displayed per page
const CLASSLOCATION					= "/gab_Library/class_drawtable/"												'we need the location of our table
const PAGING_AMOUNT_OF_NUMBERS		= 10

'TEXT
const TXTSTRPRINT					= "Print this page"
const TXTRECORDSPLURAL				= "records"																		'word for records. Plural!
const TXTRECORDSSINGULAR			= "record"																		'word for records. Singular!
const TXTNORECSAVAILABLE			= "No records found!"															'text if there are no records to display
const TXTPAGINGNEXT					= ">>"																			'something to make a next link if paging allowed
const TXTPAGINGPREV					= "<<"																			'something to make a prev link if paging allowed
const TXTPAGINGALL					= "ALL"																			'a text for "show all"
const TXTPAGINGSEPERATOR			= ""																			'seperator between pagenumbers of paging
const TXTADDNEWRECORD				= "Add record"																	'Text for adding a new record
const TXTEXPORTEXCEL				= "Export to Excel"																'Text for exporting excel
const TXTFASTDELETEBUTTON			= "&nbsp;del&nbsp;"																'text for the fast delete button
const TXTFASTDELETEQUESTION			= "Are you sure you want to delete?"											'text for asking before using fast delete
const TXTCOMMONFILTER				= "- Filter -"																	'common text for a filter dropdown
const TXTSEARCHALL					= "Search all > "
const TXTSELECTALLLASTSLONG			= "ATTENTION:\n\nThis action can last very long if there are a lot of records to update.\nPlease be sure to wait as long as the action has been finished.\n\nDo you want to proceed?"
const TXTSELECTALL					= "select all"

const TOOLTIP_TXT_SELECTALL			= "Click here to select all radiobuttons for this columns"
const TOOLTIP_TXT_SORTBY			= "Click here to sort by"
const TOOLTIP_TXT_FILTERRESET		= "Reset the filter"
const TOOLTIP_TXT_FASTDELETE		= "Click here to delete this record immediately!"
const TOOLTIP_TXT_SHOWALL			= "show all records"
const TOOLTIP_TXT_NEXTPAGE			= "next page"
const TOOLTIP_TXT_PREVPAGE			= "previous page"
const TOOLTIP_TXT_FILTER_TEXTFIELD	= "Enter a text and press ENTER to use filter on '{fieldname}'.  Allowed number operators: <,>,=,<>" 'Description for filter. {fieldname} will be replaced by fieldname
const TOOLTIP_TXT_SEARCHALL			= "Enter a keyword to search the whole data. All filters will be reseted automatically!"

dim ALLOWED_FILTER_OPERATORS(3)		'Here you can define the operators you want to allow on the filters. example <100 .. will return all recordsets smaller 100
ALLOWED_FILTER_OPERATORS(0) = "<>"
ALLOWED_FILTER_OPERATORS(1) = "<"
ALLOWED_FILTER_OPERATORS(2) = ">"
ALLOWED_FILTER_OPERATORS(3) = "="
%>