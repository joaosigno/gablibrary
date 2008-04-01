function cursorLeft(times,n){ 
	for (i = 1; i <= times; i++)
		app.ActiveDocument.CursorLeft(n)
}

function cursorRight(times,n){ 
	for (i = 1; i <= times; i++)
		app.ActiveDocument.CursorRight(n)
}

function cursorUp(times,n){ 
	for (i = 1; i <= times; i++)
		app.ActiveDocument.CursorUp(n)
}

function cursorDown(times,n){ 
	for (i = 1; i <= times; i++)
		app.ActiveDocument.CursorDown(n)
}