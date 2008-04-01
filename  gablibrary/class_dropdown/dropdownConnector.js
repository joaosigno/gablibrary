/*****************************************************************************************************
* DROPDOWN OBJECT
* CREATED BY: Michal Gabrukiewicz - gabru @ grafix.at
* mainly for moving items from one to another dropdown, selecting all options, etc.
* Example: 	create a dropdown and select all options in the onclick-event of a button:
			onclick="(new Dropdown(frm.dd)).selectAll()". sort = 0 (no sort), 1 = asc, 2 = desc
*****************************************************************************************************/
function Dropdown(dropdown, sort) {
	if (!dropdown.id) dropdown = document.getElementById(dropdown);
	this.dropdown = dropdown;
	this.sorting = 1;
	if (arguments.length > 1) this.sorting = sort;
}

/*****************************************************************************************************
* @DESCRIPTION:		selects all options of a given dropdown
*****************************************************************************************************/
Dropdown.prototype.selectAll = function() {
	if (!this.dropdown) return;
	for (var i = 0; i < this.dropdown.options.length; i++) {
		this.dropdown.options[i].selected = true;
	}
}

/*****************************************************************************************************
* @DESCRIPTION:		selects all which are not selected and viceversa
*****************************************************************************************************/
Dropdown.prototype.invert = function() {
	if (!this.dropdown) return;
	for (var i = 0; i < this.dropdown.options.length; i++) {
		this.dropdown.options[i].selected = (!this.dropdown.options[i].selected);
	}
}

/*****************************************************************************************************
* @DESCRIPTION:		moves all items from one dropdown to another one
* @PARAM:			to [select-element, ID]: the target dropdown. e.g. frm.to
*****************************************************************************************************/
Dropdown.prototype.moveAll = function(to) {
	if (!to.id) to = document.getElementById(to);
	if (!this.dropdown || !to) return;
	this.moveOptions(to, false);
}

/*****************************************************************************************************
* @DESCRIPTION:		moves just the selected items from one dropdown to another one
* @PARAM:			to [select-element, ID]: the target dropdown. e.g. frm.to
* @PARAM:			onlySelected [bool]: OPTIONAL! move just the selected options? default = true
*****************************************************************************************************/
Dropdown.prototype.moveOptions = function(to, onlySelected) {
	if (!to.id) to = document.getElementById(to);
	if (!this.dropdown || !to) return;
	if (this.dropdown.selectedIndex == -1 && (onlySelected == true || arguments.length < 2)) return;
	
	to.selectedIndex = -1;
	
	//move them to the target dropdown
	for (i = 0; i < this.dropdown.options.length; i++) {
		var o = this.dropdown.options[i];
		if ((o.selected && (onlySelected == true || arguments.length < 2)) || onlySelected == false) {
			to.options[to.options.length] = new Option(o.text, o.value, false, true);
		}
	}
	//delete them from the source dropdown
	for (i = (this.dropdown.options.length - 1); i >= 0; i--) {
		o = this.dropdown.options[i];
		if ((o.selected && (onlySelected == true || arguments.length < 2)) || onlySelected == false)
			this.dropdown.options[i] = null;
	}
	this.sort(this.dropdown);
	this.sort(to);
}

/*****************************************************************************************************
* @DESCRIPTION:		takes all the selected from one to another and the other way round
* @PARAM:			to [select-element, dd]: the dropdown you want swap with
*****************************************************************************************************/
Dropdown.prototype.swapOptions = function(to) {
	if (!to.id) to = document.getElementById(to);
	if (!to) return;
	
	oldLength = to.options.length;
	for (i = 0; i < this.dropdown.options.length; i++) {
		o = this.dropdown.options[i];
		if (o.selected) to.options[to.options.length] = new Option(o.text, o.value, false, true);
	}
	
	for (i = (this.dropdown.options.length - 1); i >= 0; i--) {
		o = this.dropdown.options[i];
		if (o.selected) this.dropdown.options[i] = null;
	}
	
	for (i = 0; i < oldLength; i++) {
		o = to.options[i];
		if (o.selected) this.dropdown.options[this.dropdown.length] = new Option(o.text, o.value, false, true);
	}
	
	for (i = (oldLength - 1); i >= 0; i--) {
		if (to.options[i].selected) to.options[i] = null;
	}
	this.sort(this.dropdown);
	this.sort(to);
}

/*****************************************************************************************************
* @DESCRIPTION:		sorts a given dropdowns. STATIC method!
* @PARAM:			dd [select-element]: the dropdown you want to sort
*****************************************************************************************************/
Dropdown.prototype.sort = function(dd) {
	if (!dd || this.sorting == 0) return;
	
	var o = new Array();
	for (i = 0; i < dd.options.length; i++) {
		current = dd.options[i];
		o[o.length] = new Option(current.text, current.value, current.defaultSelected, current.selected);
	}
	
	asc = (this.sorting == 1);
	o = o.sort( 
		function(a, b) { 
			if ((a.text + "").toUpperCase() < (b.text + "").toUpperCase()) return ((asc) ? -1 : 1);
			if ((a.text + "").toUpperCase() > (b.text + "").toUpperCase()) return ((asc) ? 1 : -1);
			return 0;
		} 
	);

	for (i = 0; i < o.length; i++) {
		dd.options[i] = new Option(o[i].text, o[i].value, o[i].defaultSelected, o[i].selected);
	}
}

