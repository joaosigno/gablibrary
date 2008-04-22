function gablib() {}

gablib.doctypeEnabled = function() {
	return document.compatMode != "BackCompat";
}

gablib.getStyle = function(element, style) {
	if (element.currentStyle) return element.currentStyle[style];
	else {
		var css = document.defaultView.getComputedStyle(element, null);
		return css ? css[style] : null;
	}
}

gablib.setStyle = function(element, styles) {
	for (var property in styles) {
		element.style[property] = styles[property];
	}
}

//toggles the visibility of an element (eID)
//displayType is optional and by default block (inline possible)
//returns true if it is visible
gablib.toggle = function(eID, displayType) {
	el = byID(eID);
	if (!el) return false;	
	if (el.style.display.toLowerCase() == 'none') {
		el.style.display = (arguments.length > 1) ? displayType : 'block';
		return true;
	} else {
		el.style.display = 'none';
	}
}

//optional: onComplete, url (because of bug in iis5 http://support.microsoft.com/kb/216493)
gablib.callback = function(theAction, func, params, onComplete, url) {
	if (params) {
		params = $H(params);
	} else {
		if ($('frm')) {
			params = $H($('frm').serialize(true));
		} else {
			params = new Hash();
		}
	} 
	params = params.merge({gabLibPageAjaxed: 1, gabLibPageAjaxedAction: theAction});
	uri = window.location.href;
	if (uri.endsWith('/') && url) uri += url;
	if (url) uri = url;
	new Ajax.Request(uri, {
		method: 'post',
		parameters: params,
		requestHeaders: {Accept: 'application/json'},
		onSuccess: function(trans) {
			//alert(trans.responseText);
			if (!trans.responseText.startsWith('{ "root":')) {
				gablib.callbackFailure(trans);
			} else {
				if (func) func(trans.responseText.evalJSON(true).root);
			}
		},
		onFailure: gablib.callbackFailure,
		onComplete: onComplete
	}); 
}
gablib.callbackFailure = function(transport) {
	friendlyMsg = transport.responseText;
	friendlyMsg = friendlyMsg.replace(new RegExp("(<head[\\s\\S]*?</head>)|(<script[\\s\\S]*?</script>)", "gi"), "");
	friendlyMsg = friendlyMsg.stripTags();
	friendlyMsg = friendlyMsg.replace(new RegExp("[\\s]+", "gi"), " ");
	alert(friendlyMsg);
}

//opens a centered modaldialog with the wanted parameters. 
function openCenteredModal(url, width, height, scrolling) {
	strScroll = (scrolling) ? "Yes" : "No";
	return window.showModalDialog(url, window, 'dialogHeight: ' + height + 'px; dialogWidth: ' + width + 'px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No; scroll: ' + strScroll);
}

//opens a modaldialog right-bottom to the mouse. if no place on the bottom then to the top.
function openModal(url, width, height, scrolling) {
	strScroll = (scrolling) ? "Yes" : "No";
	var mouseX 	= window.event.screenX;
	var mouseY = window.event.screenY;
	return window.showModalDialog(url, window, 'center:no; dialogHeight: ' + height + 'px; dialogTop:' + mouseY + '; dialogLeft:' + mouseX + 'px; dialogWidth: ' + width + 'px; edge: Raised; help: No; resizable: No; status: No; scroll: ' + strScroll);
}

function handleModalEscape() {
	currentKeyCode = event.keyCode;
	if (typeof(window) != undefined) {
	
		if (currentKeyCode == 27) window.close();
	}
}

function byID(ID) {
	return document.getElementById(ID);
}