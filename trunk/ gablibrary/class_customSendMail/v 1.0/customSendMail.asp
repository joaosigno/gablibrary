<!--#include virtual="/gab_LibraryConfig/_customSendMail.asp"-->
<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		customSendMail
'' @CREATOR:		Michal Gabrukiewicz - gabru @ grafix.at
'' @CREATEDON:		20.11.2003
'' @CDESCRIPTION:	This class is just like an interface for email-sending because of a lot of different
''					components for sending emails. This class must be written by the user himself.
''					The needed methods and properties should be available. There is also another
''					pro using this class: if you ever change your mail component you just need to implement
''					it to this class and all you applications are working with the new component.
'' @VERSION:		1.0

'**************************************************************************************************************
class CustomSendMail

	private mailerObject, p_errorMessage
	private loggerIdentification, mailServerUsername, mailServerPassword
	
	public subject				''[string] Email Subject
	public fromEmail			''[string] Senders Email. default = auto-bot-email from the constants
	public fromName				''[string] Senders Name. default = auto-bot-name from the constants
	public body					''[string] Body of the email. If htmlEmail true then it should be html-code
	public htmlEmail			''[bool] Should the body be HTML encoded? - Note: HTML and BODY-Tags are added to the content automatically,
								''so you dont need to provide these things. Start directly with your real content.
	public companyLayout		''[bool] Add the company layout to HTML emails? Header and footer is in emailSettings.asp. Default = true
	
	'******************************************************************************************************************
	'* constructor 
	'******************************************************************************************************************
	private sub class_Initialize()
		loggerIdentification 	= lib.init(GL_CMAIL_LOGGERIDENTIFICATION, "emails")
		set mailerObject 		= Server.CreateOBject("JMail.Message")
		p_errorMessage			= empty
		subject					= empty
		fromEmail				= lib.init(GL_CMAIL_FROMEMAIL, consts.automated_bot_email)
		fromName				= lib.init(GL_CMAIL_FROMNAME, consts.automated_bot_name)
		body					= empty
		htmlEmail				= lib.init(GL_CMAIL_HTMLEMAIL, true)
		companyLayout			= lib.init(GL_CMAIL_COMPANYLAYOUT, true)
		mailServerUsername		= lib.init(GL_CMAIL_MAILSERVERUSERNAME, empty)
		mailServerPassword		= lib.init(GL_CMAIL_MAILSERVERPASSWORD, empty)
	end sub
	
	'******************************************************************************************************************
	'* destructor 
	'******************************************************************************************************************
	private sub class_terminate()
		Set mailerObject = Nothing
	end sub
	
	public property get errorMessage() ''[string] get the errormessage (available after send())
		errormessage = p_errorMessage
	end property
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Sends the email according to the set properties
	'' @DESCRIPTION:	if sending failed use the errorMessage property to get the detailed error.
	'' @RETURN:			[bool] wheather the email was sent successfully or not. 
	'******************************************************************************************************************
	public function send()
		with mailerObject
			if not isEmpty(mailServerUsername) then .mailServerUserName = mailServerUsername
			if not isEmpty(mailServerPassword) then .mailServerPassword = mailServerPassword
			if consts.UTF8 then .charset = "utf-8"
			.logging = true
			.silent = true
			.ISOEncodeHeaders = false 'iso-8859-1 error in subject
			.from = fromEmail
			.fromName = fromName
			.subject = subject
			if htmlEmail then
				if companyLayout then
					'we attach the company-logo.
					logoID = addAttachment(server.mappath(consts.logo), true, empty)
					.HTMLBody = HTMLHeader() & emailHeader(logoID) & body & emailFooter() & HTMLFooter()
				else
					'Html-content incl. HTML header and footer (HTML and BODY-tags)
					.HTMLBody = HTMLHeader() & body & HTMLFooter()
				end if
			else
				.body = body
			end if
			
			logRecipients = replace(ucase(mailerObject.recipientsString), vbcrlf, "; ")
			if .send(consts.smtp_server) then
				call lib.logAndForget(loggerIdentification, """" & subject & """ sent from " & ucase(fromEmail) & " to " & logRecipients)
				send = true
			else
				lib.logAndForget loggerIdentification, "FAILED(" & mailerObject.errorsource & "; " & mailerObject.errormessage & ")! """ & subject & """ from " & ucase(fromEmail) & " to " & logRecipients
				p_errorMessage = mailerObject.errormessage
				send = false
			end if
			'str.writeln("<textarea>" & .log & "</textarea>")
			'str.end
		end with
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Clears the recipients list
	'******************************************************************************************************************
	public sub clearRecipients()
		mailerObject.clearRecipients()
	end sub
	
	'******************************************************************************************************************
	'* HTMLHeader 
	'******************************************************************************************************************
	private function HTMLHeader()
		HTMLHeader = "<html><body>"
	end function
	
	'******************************************************************************************************************
	'* HTMLFooter 
	'******************************************************************************************************************
	private function HTMLFooter()
		HTMLFooter = "</body></html>"
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	returns all recipients
	'' @RETURN:			[collection]
	'**********************************************************************************************************
	public function getRecipients()
		set getRecipients = mailerObject.recipients
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Adds a recipient to the email-object. Use this method as often as you want to add recipients
	''					To your email.
	'' @PARAM:			- email [string]: Recipients email
	'' @PARAM:			- name [string]: Recipients name
	'' @PARAM:			- toWhat [string]: Define the type of to. E.g. CC, BCC, TO, etc.
	'' @RETURN:			[bool] wheather the email was sent successfully or not. 
	'******************************************************************************************************************
	public sub addRecipient(toWhat, email, name)
		select case ucase(toWhat)
			case "TO"
				mailerObject.addRecipient email, name
			case "CC"
				mailerObject.addRecipientCC email, name
			case "BCC"
				mailerObject.addRecipientBCC email, name
		end select
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	adds more than one recipient seperated by ";". Only email values. Name will be the email
	'' @PARAM:			recipientType [string]: type of the recipient. TO, CC or BCC
	'' @PARAM:			emails [string]: recipients emails seperated by ";"
	'**********************************************************************************************************
	public sub addRecipients(recipientType, emails)
		if inStr(emails, ";") > 0 then
			arrEmails = split(emails, ";")
			for i = 0 to ubound(arrEmails)
				addRecipient recipientType, arrEmails(i), arrEmails(i)
			next
		else
			addRecipient recipientType, emails, emails
		end if
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Adds an attachment to the email-object. Use this method as often as you want to add attachments
	''					to your email.
	'' @PARAM:			- filename [string]: the name of your file (absolute)
	'' @PARAM:			- isInline [bool]: if true the attachment will be added as an inline attachment
	'' @PARAM:			- contentType [string]: Define the content type of the attachment (empty if unknown)
	'' @RETURN:			[int] content.id if isInline is true
	'******************************************************************************************************************
	public function addAttachment(fileName, isInline, contentType)
		if contentType = empty then
			addAttachment = mailerObject.addAttachment(fileName, isInline)
		else
			addAttachment = mailerObject.addAttachment(fileName, isInline, contentType)
		end if
	end function
	
	%><!--#include virtual="/gab_Library/class_customSendMail/emailSettings.asp"--><%

end class
lib.registerClass("CustomSendMail")
%>
