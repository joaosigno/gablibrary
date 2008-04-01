<%
'**************************************************************************************************************

'' @CLASSTITLE:		Spinner
'' @CREATOR:		Michael Rebec / Michal Gabrukiewicz
'' @CREATEDON:		13.07.2004
'' @CDESCRIPTION:	Just spin everything you want ;D. It's the .Net NumericUpDown / DomainUpDown Control in one.
''					If you want to use your own Items, change the ControlType to CONTROL_CUSTOM
'' @VERSION:		0.2

'**************************************************************************************************************

const ORIENTATION_LEFTRIGHT 	= 1
const ORIENTATION_TOPDOWN		= 2
const CONTROL_NUMERIC			= 1
const CONTROL_CUSTOM			= 2
const SKIN_DEFAULT				= 1
const SKIN_TWO					= 2

class spinner

	public firstInstance	'' [bool] determines if this instance is the first on the page.
							'' Its needed for the javascripts & stylesheet initialization
	public orientation		'' [enum-orientation] what orientation of the spinner? ORIENTATION_LEFTRIGHT or ORIENTATION_TOPDOWN
	public controlType		'' [enum-controltype] is it a numeric spinner or a custom spinner (e.g. with strings, etc.)
	public controlItems		'' [array] an array with your custom items which should be spinned around. works only if controlType is set to custom
	public skin				'' [enum-skin] which skin should be used for the control
	public minimum			'' [int] the minimum-value
	public maximum			'' [int] the maximum-value
	public step				'' [int/float(string)] the increment value; if you use a floatingpoint number type it as string
	public decimalPlaces	'' [int] rounds to the number of places and fills up with 0 if needed
	public integerPlaces	'' [int] fills the number from left with zero's, e.g. 3 => 000, 050, 075, 100
	public looping			'' [bool] allow repeating if true (works only if min or max != 0). default = on
	public styleAttribute	'' [string] modify the style attributes
	public readOnly			'' [bool] disable editing of the spinBox. default = false
	public width			'' [int] width in pixel of the control. default = 100
	public maxLength		'' [int] the maximum allowed length of the value-field
	public onchange			'' [string] here you can add a javascript call which will be executed on-spinner-change
	
	private classLocation
	private imgSrcUp
	private imgSrcDown
	private imgSrcLeft
	private imgSrcRight
	private BUTTON_WIDTH
	private p_contentData
	
	'**************************************************************************************************************
	'* Class_Initialize
	'**************************************************************************************************************
	sub Class_Initialize()
		firstInstance	= true
		looping			= true
		readOnly		= false
		orientation		= ORIENTATION_TOPDOWN
		controlType		= CONTROL_NUMERIC
		skin			= SKIN_TWO
		classLocation	= "/gab_Library/class_spinner/"
		minimum			= 0
		maximum			= 0
		step			= 1
		decimalPlaces	= 0
		integerPlaces	= 0
		width 			= 100
		maxLength		= 0
		BUTTON_WIDTH	= 14
		styleAttribute	= empty
		controlItems	= empty
		onchange		= empty
		p_contentData	= empty
	end sub
	
	'**************************************************************************************************************
	'' @SDECRIPTION:	draws the control
	'' @PARAM:			- spin_ID [string]: name of the control. the textfield will be named like this
	'' @PARAM:			- startValue [variant]: what value should be the startvalue
	'**************************************************************************************************************
	public sub draw(spin_ID, startValue)
		call includeScripts(spin_ID, startValue)
		call drawSpinnerControl(spin_ID, startValue)
	end sub
	
	'**************************************************************************************************************
	'* includeScripts - includes the javascript and stylesheet blocks
	'**************************************************************************************************************
	private sub includeScripts(spinID, startValue)
		dim tmp : tmp = empty
		dim pStartID : pStartID = 0
		
		if firstInstance then
			writeln("<Script language=""JavaScript"" src=""" & classLocation & "javascript.js""></SCRIPT>")
			if skin = SKIN_TWO then
				writeln("<link rel=stylesheet type=text/css href=""" & classLocation & "styles/two.css"">")
				writeln("<script language=""JavaScript"" src=""" & classLocation & "styles/two.js""></script>")
			else
				writeln("<link rel=stylesheet type=text/css href=""" & classLocation & "styles/default.css"">")
				writeln("<script language=""JavaScript"" src=""" & classLocation & "styles/default.js""></script>")
			end if
			firstInstance = false
		end if
		
		if skin = SKIN_TWO then
			imgSrcUp 	= classLocation & "images/two/up.gif"
			imgSrcDown	= classLocation & "images/two/down.gif"
			imgSrcLeft	= classLocation & "images/two/left.gif"
			imgSrcRight	= classLocation & "images/two/right.gif"
		else
			imgSrcUp 	= classLocation & "images/default/up.gif"
			imgSrcDown	= classLocation & "images/default/down.gif"
			imgSrcLeft	= classLocation & "images/default/left.gif"
			imgSrcRight	= classLocation & "images/default/right.gif"
		end if
		
		writeln("<Script language=""JavaScript"">")
		writeln("preloadImages();")
		writeln("var vdo" & spinID & " = " & orientation & ";" & _
				"var vdd" & spinID & " = " & decimalPlaces & ";" & _
				"var vdi" & spinID & " = " & integerPlaces & ";")
		
		if IsArray(controlItems) and controltype = CONTROL_CUSTOM then
			writeln("var vdCustomMode_" & spinID & " = 1;")
			writeln("var vdCustomValues_" & spinID & " = new Array(")
			for i = 0 to UBound(controlItems)
				tmp = tmp & "'" & controlItems(i) & "'"
				if not i = UBound(controlItems) then 
					tmp = tmp & ", "
				else
					tmp = tmp & ");"
				end if
				if controlItems(i) = startValue then
					pStartID = i
				end if
			next
			writeln(tmp & vbcrlf & "var vdArrayCount_" & spinID & " = " & pStartID & ";")
		else
			writeln("var vdCustomMode_" & spinID & " = 0;")
		end if
		
		writeln("</Script>")
	end sub
	
	'**************************************************************************************************************
	'* getToolTip - returns the tooltips
	'**************************************************************************************************************
	private function getToolTip(mtype)
		select case mtype:
			case "spinbox": 
				if minimum = 0 and maximum = 0 then
					getToolTip = "Minimum: inf - Maximum: inf - Step: " & step
				elseif controlType = CONTROL_CUSTOM and IsArray(controlItems) then
					getToolTip = "Minimum: " & controlItems(0) & " - Maximum: " & controlItems(Ubound(controlItems))
				else
					getToolTip = "Minimum: " & minimum & " - Maximum: " & maximum & " - Step: " & step
				end if
					getToolTip = getToolTip & vbcrlf & vbcrlf & "Hint: You can change the value by using the mousewheel"
			case "spin_up":
				getToolTip = "Increment value by " & step
			case "spin_down":
				getToolTip = "Decrease value by " & step
		end select
	end function
	
	'**************************************************************************************************************
	'* checkReadonly - makes the spinbox readonly if true
	'**************************************************************************************************************
	private function checkReadonly()
		if readonly then
			checkReadonly = " readonly"
		else
			checkReadonly = empty
		end if
	end function
	
	'**************************************************************************************************************
	'* drawSpinnerTextfield - draws the spinner-textbox
	'**************************************************************************************************************
	private sub drawSpinnerTextfield(spinID, startValue)
		if orientation = ORIENTATION_TOPDOWN then
			callCountParam = "topdown"
			cssClass = "spinnerBoxTD"
			calculatedWidth = width - BUTTON_WIDTH
			marginCorrection = empty
		else
			callCountParam = "leftright"
			cssClass = "spinnerBoxLR"
			calculatedWidth = width - 2 * BUTTON_WIDTH
			marginCorrection = "margin-left:" & BUTTON_WIDTH & "px"
		end if
		
		onMouseWheel = "shiftUpDown(this, '" & step & "', " & minimum & ", " & maximum & ", '" & looping & "');" & onchange
		onKeyDown = "callCount('" & callCountParam & "', " & spinID & ", " & step & ", " & minimum & ", " & maximum & ", '" & looping & "');" & onchange
		onBlur	= "checkCorrectInput(this, " & minimum & ", " & maximum & ", " & step & ");" & onchange
		
		writeln("<input type=""text"" id=""" & spinID & """ name=""" & spinID & """ " & checkReadonly() & " value=""" & startValue & """ " &_
					"title=""" & getToolTip("spinbox") & """ " &_
					"onmousewheel=""" & onMouseWheel & """ " &_
					"onkeydown=""" & onKeyDown & """ " &_
					"onblur=""" & onBlur & """ " &_
					"class=""xx " & cssClass & """ " &_
					"style=""width:" & calculatedWidth & "px;" & marginCorrection & """" & getMaxLength() & ">")
	end sub
	
	'**************************************************************************************************************
	'* getMaxLength 
	'**************************************************************************************************************
	private function getMaxLength()
		getMaxLength = empty
		'we check if maxlength is larger than 0 and is larger than the maximum and minimum allowed value
		if maxLength > 0 then
			if controlType = CONTROL_NUMERIC then
				if maxLength >= len(minimum) and maxLength >= len(maximum) then
					getMaxLength = " maxLength=""" & maxLength & """"
				end if
			end if
		end if
	end function
	
	'**************************************************************************************************************
	'* drawUpButton 
	'**************************************************************************************************************
	private sub drawIncreaseButton(spinID)
		if orientation = ORIENTATION_TOPDOWN then
			cssClass = "spinnerControlButtonUp"
			imgSrc = imgSrcUp
		else
			cssClass = "spinnerControlButtonRight"
			imgSrc = imgSrcRight
		end if
		
		onClick = "callCount('up'," & spinID & ", '" & step & "', " & minimum & ", " & maximum & ", '" & looping & "');" & onchange
		writeln("<button class=""spinnerControlButton " & cssClass & """ id=""up_" & spinID & """ onclick=""" & onClick & """ tabindex=""-1"">" &_
						"<img src=""" & imgSrc & """ id=""img" & spinID & "_up"" alt=""" & getToolTip("spin_up") & """>" &_
					"</button>")
	end sub
	
	'**************************************************************************************************************
	'* drawDecreaseButton 
	'**************************************************************************************************************
	private sub drawDecreaseButton(spinID)
		if orientation = ORIENTATION_TOPDOWN then
			cssClass = "spinnerControlButtonDown"
			imgSrc = imgSrcDown
		else
			cssClass = "spinnerControlButtonLeft"
			imgSrc = imgSrcLeft
		end if
		
		onClick = "callCount('down'," & spinID & ", '" & step & "', " & minimum & ", " & maximum & ", '" & looping & "');" & onchange
		writeln("<button class=""spinnerControlButton " & cssClass & """ id=""down_" & spinID & """ onclick=""" & onClick & """ tabindex=""-1"">" &_
						"<img src=""" & imgSrc & """ id=""img" & spinID & "_down"" alt=""" & getToolTip("spin_down") & """>" &_
					"</button>")
	end sub
	
	'**************************************************************************************************************
	'* drawSpinnerButtons - draw the spinner-control at the correct place ;)
	'**************************************************************************************************************
	private sub drawSpinnerControl(spinID, startValue)
		'we need to recalculate the width because of placement with stylesheet. we need to subtract
		'the button width from the wanted width and then we get the wanted width.
		recalculatedWidth = width - BUTTON_WIDTH
		
		'we need to give a right margin because the buttons are placed absolute and so the whole control would be smaller
		writeln("<span style=""margin-right:" & BUTTON_WIDTH & "px;width:" & recalculatedWidth & "px;" & styleAttribute & """>")
		if orientation = ORIENTATION_TOPDOWN then
			call drawSpinnerTextfield(spinID, startValue)
			writeln("<span class=spinnerControlButtons>")
			call drawIncreaseButton(spinID)
			call drawDecreaseButton(spinID)
			writeln("</span>")
		else
			writeln("<span class=spinnerControlButtons>")
			call drawDecreaseButton(spinID)
			writeln("</span>")
			call drawSpinnerTextfield(spinID, startValue)
			writeln("<span class=spinnerControlButtons>")
			call drawIncreaseButton(spinID)
			writeln("</span>")
		end if
		writeln("</span>")
	end sub
	
	'**************************************************************************************************************
	'* writeln - is more a sub than a function, but so i dont have to use call :)
	'**************************************************************************************************************
	private function writeln(msg)
		'p_contentData = p_contentData & msg
		response.write(msg)
	end function
	
	public property get toString ''wont work yet
		toString = p_contentData
	end property
end class
%>