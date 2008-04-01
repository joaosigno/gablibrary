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
''					This Class has been updated to support ASPemail as well as JMail.
'' @VERSION:		1.5

'**************************************************************************************************************
class CustomSendMail

	private mailerObject, p_errorMessage
	private loggerIdentification, mailServerUsername, mailServerPassword, recipients
	private component
	
	public subject				''[string] Email Subject
	public fromEmail			''[string] Senders Email. default = auto-bot-email from the constants
	public fromName				''[string] Senders Name. default = auto-bot-name from the constants
	public body					''[string] Body of the email. If htmlEmail true then it should be html-code
	public htmlEmail			''[bool] Should the body be HTML encoded? - Note: HTML and BODY-Tags are added to the content automatically,
								''so you dont need to provide these things. Start directly with your real content.
	public companyLayout		''[bool] Add the company layout to HTML emails? Header and footer is in emailSettings.asp. Default = true
	
	public property get errorMessage() ''[string] get the errormessage (available after send())
		errormessage = p_errorMessage
	end property
	
	'******************************************************************************************************************
	'* constructor 
	'******************************************************************************************************************
	private sub class_Initialize()
		loggerIdentification 	= lib.init(GL_CMAIL_LOGGERIDENTIFICATION, "emails")
		p_errorMessage			= empty
		subject					= empty
		fromEmail				= lib.init(GL_CMAIL_FROMEMAIL, consts.automated_bot_email)
		fromName				= lib.init(GL_CMAIL_FROMNAME, consts.automated_bot_name)
		body					= empty
		htmlEmail				= lib.init(GL_CMAIL_HTMLEMAIL, true)
		companyLayout			= lib.init(GL_CMAIL_COMPANYLAYOUT, true)
		mailServerUsername		= lib.init(GL_CMAIL_MAILSERVERUSERNAME, empty)
		mailServerPassword		= lib.init(GL_CMAIL_MAILSERVERPASSWORD, empty)
		component				= lib.init(GL_CMAIL_COMPONENT, "JMAIL")		
		if component = "JMAIL" then
			set mailerObject = Server.CreateOBject("JMail.Message")
		elseif component = "ASPEMAIL" then
			set mailerObject = Server.CreateObject("Persits.MailSender")
		else
			lib.error("E-Mail component not configured")
		end if
	end sub 
	
	'******************************************************************************************************************
	'* destructor 
	'******************************************************************************************************************
	private sub class_terminate()
		Set mailerObject = Nothing
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Sends the email according to the set properties
	'' @DESCRIPTION:	if sending failed use the errorMessage property to get the detailed error.
	'' @RETURN:			[bool] wheather the email was sent successfully or not. 
	'******************************************************************************************************************
	public function send()
		send = false
		if component = "ASPEMAIL" then
			with mailerObject
				if not isEmpty(mailServerUsername) then .UserName = mailServerUsername
				if not isEmpty(mailServerPassword) then .Password = mailServerPassword
				.Queue = True
				if consts.UTF8 then .charset = "utf-8"
				.from = fromEmail
				.fromName = fromName
				.subject = subject
				if htmlEmail then
					.isHTML = true
					if companyLayout then
						logoID = addAttachment(server.mappath(consts.logo), true, empty)
						.Body = HTMLHeader() & emailHeader(logoID) & body & emailFooter() & HTMLFooter()
					else
						.Body = HTMLHeader() & body & HTMLFooter()
					end if
				else
					.body = body
				end if
				if .send() then
					lib.logAndForget loggerIdentification, """" & subject & """ sent from " & ucase(fromEmail) & " to " & ucase(recipients)
					send = true
				else
					lib.logAndForget loggerIdentification, "FAILED: " & subject & """ from " & ucase(fromEmail) & " to " & ucase(recipients)
					send = false
				end if
			end with
		elseif component = "JMAIL" then
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
				
				if .send(consts.smtp_server) then
					lib.logAndForget loggerIdentification, """" & subject & """ sent from " & ucase(fromEmail) & " to " & ucase(recipients)
					send = true
				else
					lib.logAndForget loggerIdentification, "FAILED(" & mailerObject.errorsource & "; " & mailerObject.errormessage & ")! """ & subject & """ from " & ucase(fromEmail) & " to " & ucase(recipients)
					p_errorMessage = mailerObject.errormessage
					send = false
				end if
			end with
		end if
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	JMAIL ONLY: Clears the recipients list
	'******************************************************************************************************************
	public sub clearRecipients()
		recipients = ""
		if component = "JMAIL" then mailerObject.clearRecipients() else lib.error(component & " does not support this method. This call is for JMail only")
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	ASPMAIL ONLY: Clears all address and attachment lists so that a new message can be sent.
	'******************************************************************************************************************
	public sub Reset()
		recipients = ""
		if component = "ASPEMAIL" then mailerObject.Reset() else lib.error(component & " does not support this method. This call is for ASPEMail only")
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
	'' @RETURN:			[string] Colon seperated email addresses
	'**********************************************************************************************************
	public function getRecipients()
		set getRecipients = recipients
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Adds a recipient to the email-object. Use this method as often as you want to add recipients
	''					To your email.
	'' @PARAM:			- email [string]: Recipients email
	'' @PARAM:			- name [string]: Recipients name
	'' @PARAM:			- toWhat [string]: Define the type of to. E.g. CC, BCC, TO, etc.
	'******************************************************************************************************************
	public sub addRecipient(toWhat, email, name)
		recipients = email & "; " & recipients ' we want to keep the recipients addresses
		if component = "ASPEMAIL" then
			select case ucase(toWhat)
				case "TO"
					mailerObject.AddAddress email, name
				case "CC"
					mailerObject.AddCC email, name
				case "BCC"
					mailerObject.AddBCC email, name
			end select
		elseif component = "JMAIL" then
			select case ucase(toWhat)
				case "TO"
					mailerObject.addRecipient email, name
				case "CC"
					mailerObject.addRecipientCC email, name
				case "BCC"
					mailerObject.addRecipientBCC email, name
			end select
		end if
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
	'' @PARAM:			- filename [string]: the name and path of your file (absolute)
	'' @PARAM:			- isInline [bool]: if true the attachment will be added as an inline attachment
	'' @RETURN:			[int] A unique ID that can be used to identify this attachment. This is useful if you are 
	''					embedding images in the email, body eg: IMG SRC="cid:xxxx"
	'******************************************************************************************************************
	public function addAttachment(fileName, isInline, contentType)
		addAttachment = lib.getUniqueID
		if 	component = "ASPEMAIL" then
			if isInline then
				 mailerObject.AddEmbeddedImage fileName, addAttachment
			else
				mailerObject.addAttachment(fileName)
			end if
		elseif component = "JMAIL" then
			if contentType = empty then
				addAttachment = mailerObject.addAttachment(fileName, isInline)
			else
				addAttachment = mailerObject.addAttachment(fileName, isInline, contentType)
			end if
		end if
	end function
	
	%><!--#include virtual="/gab_Library/class_customSendMail/emailSettings.asp"--><%

end class
lib.registerClass("CustomSendMail")
%>
