<!--#include virtual="/gab_Library/class_page/generatepage.asp"-->
<!--#include virtual="/gab_Library/class_calculator/calculator.asp"-->
<%
set page = new generatePage
with page
	.DBConnection 	= false
	.title 			= "Wyeth custom online calculator" & str.clone("&nbsp;", 100)
	.debugMode		= false
	.loginRequired	= false
	.frameSetter	= false
	.showFooter		= false
	.bodyAttribute	= "topmargin=0 lefmargin=0 onkeyup=""handleKeys();"""
	.devWarning		= false
	.isModalDialog	= false
	.draw()
end with
set page = nothing

'******************************************************************************************
'* main - we start . DONT FORGET TO KILL ALL OBJECTS!! 
'******************************************************************************************
sub main()
	str.writeln("<base target=_self>")
	set calci = new calculator
	with calci
		.JSTarget = request.queryString("JSTarget")
		.displayedValue = request.queryString("displayedValue")
		
		set btn = new calculatorButton
		with btn
			.toolTip = "Euro"
			.caption = "EUR"
			.value = "13,7461234567890"
		end with
		.addCustomButton(btn)
		set btn = nothing
		
		set btn = new calculatorButton
		with btn
			.toolTip = "US-Dollar"
			.caption = "USD"
			.value = "10,011"
		end with
		.addCustomButton(btn)
		set btn = nothing
		
		set btn = new calculatorButton
		with btn
			.toolTip = "Polnische"
			.caption = "PLN"
			.value = "4,45"
		end with
		.addCustomButton(btn)
		set btn = nothing
		
		.draw()
		
	end with
	set calci = nothing
end sub
%>