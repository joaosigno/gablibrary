<!--#include file="languages/en.asp"-->
<%
'whats the defaultdatabase
const DEFAULTDATABASE = "mssql"
'show tablerowheaders every X line?
const HEADERPERROWS = 30
'how many records should be displayed per page
const RECORDSPERPAGE = 50
'we need the location of our table
const DT_CLASSLOCATION	= "/gab_Library/class_drawtable/"
'stylesheet which should be loaded to format the drawtable. leave it empty to use the default stylesheet.
const DRAWTABLE_CSS_LOCATION = "/styles/drawtable.css"
'how many page numbers will be displayed at once
const PAGING_AMOUNT_OF_NUMBERS = 10
'The backgroundcolor when a radiobutton has been clicked
const SELECTEDBGCOLOR = "#CCCCCC"
'The color when a radiobutton has been clicked
const SELECTEDCOLOR	= "#000000"
const ROW_COLOR_1 = "#FFFFFF"
const ROW_COLOR_2 = "#FFFFFF"
const ROW_COLOR_HOVER = "#FEFDE0"

'Here you can define the operators you want to allow on the filters.
'example <100 .. will return all recordsets smaller 100
dim ALLOWED_FILTER_OPERATORS(3)
ALLOWED_FILTER_OPERATORS(0) = "<>"
ALLOWED_FILTER_OPERATORS(1) = "<"
ALLOWED_FILTER_OPERATORS(2) = ">"
ALLOWED_FILTER_OPERATORS(3) = "="
%>