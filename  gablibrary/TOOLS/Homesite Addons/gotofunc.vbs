'GotoFunc for HomeSite 4.5.1
'Author: Corey Haines (corey@elender.hu)
'Description: 	This script is intended to mimic the Shift-F2 functionality in the Visual Basic IDE. 
'Basically, this will inspect the word around the cursor, and search for a function or sub that has that name. 
'It will then scroll the editor pane to	the beginning of that method. 
'It is pretty hacked together right now, as I don't have too much experience programming the editor, itself.</p>
'Updates:
'
' This script has been released to the public domain. If you make a change, please let me know, so I can add it,
' if necessary, to the original. You are free to distribute this script, as long as you leave this header in place.
'
'--- MODIFICATIONS
'--- DATE:   Wednesday, January 19, 2005
'--- AUTHOR: William Morris, Kansas City, MO  USA (william@seritas.com)
'--- * Script now searches all open files.
'--- * Checks for an opeing parenthesis ("bracket" in the UK), line break, or whitespace after the word being
'---   searched on, to avoid false finds.  For instance, assume a file with two procedures
'---   with similar names, appearing in this order in the file:
'---	  CreateConnectionEx
'---	  CreateConnection
'---   Searching just for "CreateConnection" would find "CreateConnectionEX" and stop.
'---   The script now finds the correct entry.
'--- * The script now highlights the search term when found.

option explicit

sub Main()
	dim doc
	dim sCurWord
	dim curIndex
	dim app
	dim docCounter
	dim foundIt
	
	Const cNotFound = 0
	Const cFoundInCurrentDocument = 1
	Const cFoundInOtherDocument = 2
	Const cNotFoundNotSearched = 3
	
	foundIt = cNotFound
	set app = application
	set doc = app.ActiveDocument
	curIndex = app.documentIndex

	sCurWord = getCurrentWord(doc)

	if len(trim(sCurWord)) > 0 then
		if not moveToWord(doc, sCurWord) then
			for docCounter = 0 to app.DocumentCount - 1
				'--- don't check the document we were already in, eh?
				if docCounter <> curIndex then
					Application.DocumentIndex = docCounter
					set doc = app.activeDocument
					if moveToWord(doc, sCurWord) then
						foundit = cFoundInOtherDocument
						exit for
					end if
				end if
			next
		else
			foundIt = cFoundInCurrentDocument
		end if
	end if

	if foundit = cNotFound then
		application.documentIndex = curIndex
		msgbox "Not found.", vbexclamation, "Procedure Finder"
	end if
	
	set doc = Nothing
	
end sub

'______________________________________________________________________________

function moveToWord(doc, sWord)
	dim lTotLines
	dim lCurLine, sCurLine
	dim lLoc
	dim docCounter
	dim foundIt
	dim arrPossibles
	
	'--- expand this list as needed
	redim arrPossibles(9)
	arrPossibles(0) = "function " & sWord & " "
	arrPossibles(1) = "function " & sWord & "("
	arrPossibles(2) = "function " & sWord & chr(13)
	arrPossibles(3) = "sub " & sWord & " "
	arrPossibles(4) = "sub " & sWord & "("
	arrPossibles(5) = "sub " & sWord & chr(13)
	arrPossibles(6) = "function " & sWord & chr(10)
	arrPossibles(7) = "sub " & sWord & chr(10)
	arrPossibles(8) = "function " & sWord & vbcrlf
	arrPossibles(9) = "sub " & sWord & vbcrlf

	foundIt = 0
	lLoc = 0
	lCurLine = 1
	lTotLines = doc.LineCount
	
	moveToWord = false
	
	do until (lLoc > 0) or (lCurLine > lTotLines)
		sCurLine = trim(doc.Lines(lCurLine)) & vbcrlf
		'--- look to see if the current line contains the desired word
		'--- and save off the location
		lLoc = Instr(1, sCurLine, sWord, 1)
		if lLoc > 0 then
			'--- okay, we have the word...is it in a sub or function declaration?
			for docCounter = 0 to ubound(arrPossibles)
				if instr(1, sCurLine, arrPossibles(docCounter), 1) > 0 then
					foundIt = 1
					exit for
				end if
			next
			if foundIt = 0 then
				'--- if not, reset the location
				lLoc = 0
			else
				'--- if we do find it, there's no point in checking the rest of the file, right?
				lCurLine = lCurLine + 1
				exit do
			end if
		end if
		lCurLine = lCurLine + 1
	loop
	if lLoc > 0 then
		doc.setCaretPos lLoc, lCurLine
		'--- WM
		'--- now, highlight the word
		'--- doing a WordRight can cause multiple lines to be selected,
		'--- so better to highlight one character at a time, to the
		'--- length of the word
		for docCounter = 1 to len(sWord)
			doc.CursorRight true
		next
		moveToWord = true
	end if
end function


'______________________________________________________________________________

function getCurrentWord(doc)
	dim sCurWord
	dim bGotName
	dim sNext, sPrev

	sPrev = doc.getPreviousChar
	doc.cursorright false
	sNext = doc.getNextChar
	doc.cursorleft false

	select case true
		case asc(sPrev) = 13, asc(sPrev) = 9, len(trim(sPrev)) = 0
			'
		case asc(sNext) = 13, asc(sNext) = 9, len(trim(sNext)) = 0
			doc.cursorwordright false
		case else
			doc.cursorwordleft false
	end select
	bGotName = False
	sCurWord = ""
	do until bGotName
		select case true
			case asc(doc.getCurrentChar) = 13,asc(doc.getCurrentChar) = 9, _
				len(trim(doc.getCurrentChar)) = 0, doc.getCurrentChar="("
				bGotName = True
			case else
				sCurWord = sCurWord & doc.getCurrentChar
				doc.CursorRight false
		end select
	loop

	getCurrentWord = sCurWord
	
end function

