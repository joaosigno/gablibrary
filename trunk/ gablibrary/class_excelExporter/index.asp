<!--#include virtual="/gab_Library/class_page/generatepage.asp"-->
<!--#include virtual="/gab_Library/class_excelExporter/excelExporter.asp" -->
<%
'working DEMO of the ExcelExporter which can be used as a default
'for your excel-exports.
set page = new GeneratePage
set eExporter = new ExcelExporter
if consts.UTF8 then str.writeln("<meta http-equiv=""content-type"" content=""text/html; charset=utf-8""/>")
eExporter.export()
set eExporter = nothing
%>