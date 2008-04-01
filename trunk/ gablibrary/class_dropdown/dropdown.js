function toggleNewItem(sender, checkedField, dd, newField) {
	ddWidth = dd.offsetWidth;
	checkedField.value = (checkedField.value == "") ? "1" : "";
	checked = (checkedField.value == "1");
	dd.disabled = checked;
	dd.style.display = (checked) ? "none" : "inline";
	newField.disabled = !checked;
	newField.style.display = (!checked) ? "none" : "inline";
	if (checked) {
		newField.style.width = ddWidth;
		newField.value = "";
		newField.focus();
	} else {
		dd.focus();
	}
	sender.className = (checked) ? "dropdownNewActive" : "dropdownNewInActive";
}