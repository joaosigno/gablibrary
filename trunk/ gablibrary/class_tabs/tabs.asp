<!--#include file="const.asp"-->
<!--#include virtual="/gab_libraryConfig/_tabs.asp"-->
<!--#include file="class_tab.asp"-->
<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		tabs
'' @CREATOR:		Michal Gabrukiewicz - gabru @ grafix.at
'' @CREATEDON:		03.09.2003
'' @CDESCRIPTION:	Generates a complete & nice tabs-control. Handles the queryString automatically.
''					Every tab gets a procedure which will be executed when clicking on the tab. 
'' @VERSION:		1.0

'**************************************************************************************************************
class tabs

	private dictObj						'We need a dictonary object to store our tabs
	private dictObjCount				'we need a key for every item of the dictonary object
	private defaultObj					'which tab is selected by default?
	private p_defaultSelectedTabIndex
	
	public attribute					''[string] An attribute for everyTab. e.g. style="width:200px;"
	public addProcedure					''[string] Name of procedure. Do you want to exectue a procedure directly beside the tabs?
	public QSnotUse						''[string] Querystring-fields you dont want to take with you on postback. if you have more than one then just seperate with comma
	public printable					''[bool] should the tabs appear on printouts? default = false
	public CSSLocation					''[string] location to the stylesheetfile.
	
	'Konstruktor => set the default values
	private sub Class_Initialize()
		set dictObj						= Server.CreateObject("Scripting.Dictionary")
		dictObjCount					= 1
		set defaultObj					= new Tab
		attribute						= empty
		addProcedure					= empty
		QSnotUse						= empty
		p_defaultSelectedTabIndex		= -1
		printable						= false
		CSSLocation = lib.init(GL_TABS_CSSLOCATION, TABS_CLASSLOCATION & "styles/standard.css")
	end sub
	
	public property let defaultSelectedTabIndex(val) ''[int] index of the selected tab-index by default
		p_defaultSelectedTabIndex = val
	end property
	
	public property get selectedTab ''[int] index of the selected tab. first = 1. 0 if no selected
		selectedTab = 0
		if request.queryString("activeTab") <> "" then
			selectedTab = cInt(request.queryString("activeTab"))
		end if
	end property
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Adds a tab-object to your tabs-control
	'' @DESCRIPTION:	This method needs an instance of "tab"
	'' @PARAM:			- tabObject [TAB]: your tab-object (tab-instance)
	'******************************************************************************************************************
	public function addTab(tabObject)
		'put the tabobject to the dictonary
		dictObj.add dictObjCount, tabObject
		'increase the key
		dictObjCount = dictObjCount + 1
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Specifies the default tab. Optional!
	'' @PARAM:			- defObject [TAB]: Your default object (tab-instance)
	'******************************************************************************************************************
	public function default(defObject)
		'put the tabobject to the dictonary
		set defaultObj = defObject
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Draws the tabs-control
	'******************************************************************************************************************
	public function draw()
		beginning()
		callSub = main()
		ending()
		'callsub has the info about the sub we need to execute
		if not callSub = empty then execute("call " & callSub)
	end function
	
	'**************************************************************************************************************
	' header 
	'**************************************************************************************************************
	private function main()
		'we got through all tabs
		for each k in dictObj.keys
			'we check if there is an active tab
			if selectedTab = cint(k) then
				isActive = true
				main = dictObj(k).procedure
			'here we  check if maybe there is an default tab
			elseif selectedTab = 0 and (dictObj(k).caption = defaultObj.caption or cInt(k) = cInt(p_defaultSelectedTabIndex)) then
				isActive = true
				main = dictObj(k).procedure
			else
				isActive = false
			end if
			
			'lets draw this nice litte tabby
			dictObj(k).draw isActive,k, attribute, QSnotUse
		next
		set dictObj = nothing
		set defaultObj = nothing
	end function
	
	'**************************************************************************************************************
	' beginning 
	'**************************************************************************************************************
	private sub beginning()
		'We load the default css
		str.write("<link rel=""stylesheet"" type=""text/css"" href=""" & CSSLocation & """>")
		%><!--#include file="display_beginning.asp"--><%
	end sub
	
	'**************************************************************************************************************
	' ending 
	'**************************************************************************************************************
	private sub ending()
		%><!--#include file="display_ending.asp"--><%
	end sub

end class
lib.registerClass("Tabs")
%>
