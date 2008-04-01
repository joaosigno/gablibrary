
<%
'we need the location of our table
const DT_CLASSLOCATION	= "/gab_Library/class_drawtable/"
'The backgroundcolor when a radiobutton has been clicked
const SELECTEDBGCOLOR = "#CCCCCC"
'The color when a radiobutton has been clicked
const SELECTEDCOLOR	= "#000000"

'Here you can define the operators you want to allow on the filters.
'example <100 .. will return all recordsets smaller 100
dim ALLOWED_FILTER_OPERATORS(3)
ALLOWED_FILTER_OPERATORS(0) = "<>"
ALLOWED_FILTER_OPERATORS(1) = "<"
ALLOWED_FILTER_OPERATORS(2) = ">"
ALLOWED_FILTER_OPERATORS(3) = "="
%>