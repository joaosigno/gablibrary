function showDetails(ident, description, file) {
	window.showModalDialog("details.asp?description="+ description + "&file="+ file +"&id="+ ident +"", "", "dialogHeight:400px; dialogWidth:600px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No; scroll: Auto");
}