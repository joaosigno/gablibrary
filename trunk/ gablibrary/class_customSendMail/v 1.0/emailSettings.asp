<%
'******************************************************************************************************************
'* emailHeader 
'* returns the custom email-header as string. 
'* please return an empty string if you dont want to add a header to your mails. 
'* PARAM: logoContentID is the ID of the attached LOGO to the mail. 
'******************************************************************************************************************
private function emailHeader(logoContentID)
	emailHeader = _
		"<div style=""position:absolute;right:20px;top:20px;font-size:8pt;color:#FF0000;font-family:verdana;text-align:right;"">" &_
			"<img src=""cid:" & logoContentID & """ border=0 alt=""" & consts.company_name & """>" & _
		"</div><div style=""font:12px verdana;"">"
end function

'******************************************************************************************************
'* emailFooter 
'* returns the custom email-footer as string. 
'* return empty string if you dont need a footer. Footer will be added only for HTML mails
'******************************************************************************************************
private function emailFooter()
	emailFooter = _
	"</div><div style=""background-color:#ccc; text-align:right; margin-top:30px; width:410px;font-family:verdana;"">" &_
		"<div style=""font-size:8pt; color:#000; background-color:#fff; padding:7px; text-align:left; width:400px;"">" &_
			"This is an automatically generated email<br>" &_
			"sent through " & consts.company_name & " Intranet on " & now() & "<br>" &_
			str.trimStart(consts.domain, 7) &_
		"</div>" &_
	"</div>"
end function
%>