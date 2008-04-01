/***************************************************************************************************************
* creates a XML file for the smartSense for gabLibrary
* Michal Gabrukiewicz
***************************************************************************************************************/
function Main() {	
	var app = Application;
	gabLibDir = getGablibDir(app);
	xmlFile = gabLibDir + "TOOLS\\Homesite Addons\\smartSense\\langDef\\gabLibrary.xml";
	
	if (app.MessageBox("Do you want to update the smartSense definitions for gabLibrary?", "Update smartSense?", 4) == 6) {
		xml = new ActiveXObject("Msxml2.DOMDocument");
		pinf = xml.createProcessingInstruction("xml", "version=\"1.0\"");
		xml.insertBefore(pinf, xml.childNodes(0));
		xml.appendChild(getNewNode("root", "", xml));
		elementsNode = getNewNode("elements", "", xml);
		xml.documentElement.appendChild(elementsNode);
		typesNode = getNewNode("types", "", xml);
		xml.documentElement.appendChild(typesNode);
		
		fso = new ActiveXObject("Scripting.FileSystemObject");
		var root = fso.GetFolder(gabLibDir);
		en = new Enumerator(root.subfolders);
		
		//parse the gablibconfig
		parseFile(fso.GetFile(root.parentFolder + "\\gab_LibraryConfig.asp"), xml, typesNode, elementsNode);
		parseFile(fso.GetFile(root.parentFolder + "\\gab_LibraryConfig\\class_customlib.asp"), xml, typesNode, elementsNode);
		
		//and now parse through all folders in the gablibrary..
		for (; !en.atEnd(); en.moveNext()) {
			en2 = new Enumerator(en.item().files);
			//loop through files
			for (; !en2.atEnd(); en2.moveNext()) {
				parseFile(en2.item(), xml, typesNode, elementsNode);
			}
		}
		
		xml.save(xmlFile);
		app.MessageBox("Successfully finished. Enjoy!", "", 0);
	}
}

function parseFile(file, xml, typesNode, elementsNode) {
	//we only use the .asp files
	if (!endsWith(file.name, ".asp")) return;
	
	//if the file has no contents then we skip it.
	if (file.Size == 0) return;
	
	strm = file.OpenAsTextStream(1);
	
	content = strm.readAll();
	//we look if its a class inside!
	if (content.indexOf("'' @CLASSTITLE:") > -1) {
		lines = content.split("\r");

		//get the class-details							
		title = "";
		description = "";
		staticname = "";
		postfix = "";
		for (i = lines.length; i > 0; i--) {
			line = trim(lines[i] + "");
			if (startsWith(line, "'' @CLASSTITLE:")) {
				title = capitalize(trim(line.replace("'' @CLASSTITLE:", "")));
			} else if (startsWith(line, "'' @STATICNAME:")) {
				staticname = trim(line.replace("'' @STATICNAME:", ""));
			} else if (startsWith(line, "'' @POSTFIX:")) {
				postfix = trim(line.replace("'' @POSTFIX:", ""))
			} else if (startsWith(line, "'' @CDESCRIPTION:")) {
				description = trim(line.replace("'' @CDESCRIPTION:", ""));
				//get the description if its over more lines ...
				for (j = i + 1; j < lines.length; j++) {
					aline = trim(lines[j]);
					if (startsWith(aline, "''") && !startsWith(aline, "'' @")) {
						description += "\r" + trim(aline.replace("''", ""));
					} else {
						break;
					}
				}
			}
		}
		
		//if no class title was found then we skip this file..
		if (title == "") return;
		
		classNode = getNewNode("el", "", xml);
		classNode.appendChild(getNewNode("display", title, xml));
		//if there is a staticname then the object is available everywhere.
		//otherwise there is a need to include it, so we create an include for it
		if (staticname == "") {
			val = "<!--#include virtual=\"/gab_Library/"+ file.parentFolder.name + "/" + file.name + "\"-->";
			classNode.appendChild(getNewNode("insert", val, xml));
			//add the type...
			if (postfix != "") {
				elNode = getNewNode("el", "", xml);
				elNode.setAttribute("name", postfix);
				elNode.appendChild(getNewNode("type", val, xml));
				typesNode.appendChild(elNode);
			}
		} else {
			classNode.appendChild(getNewNode("insert", staticname, xml));
		}
		classMembersNode = getNewNode("childElements", "", xml);
		classNode.appendChild(classMembersNode);
		elementsNode.appendChild(classNode);
		
		//we iterate through all lines in the file from the back in order to get the method first
		//and then be able to fetch the description details for it
		for (i = lines.length; i > 0; i--) {
			line = trim(lines[i] + "");
			//if its something public then we need to do something with it
			if (startsWith(line, "public")) {
				//cut the public keyword
				line = line.replace("public ", "");
				if (startsWith(line, "function ") || startsWith(line, "sub ")) {
					line = line.replace("function ", "");
					line = line.replace("sub ", "");
					lineUpper = line.toUpperCase();
					//if its the constructor or destructor then skip this line...
					if (startsWith(lineUpper, "CLASS_TERMINATE") || startsWith(lineUpper, "CLASS_INITIALIZE")) continue;
					//replace some keywords from the params.
					name = line.replace("byVal ", "");
					name = name.replace("byval ", "");
					name = name.replace("byRef ", "");
					name = name.replace("byref ", "");
					description = "";
					//get the details about the method
					for (j = i - 1; j >= 0; j--) {
						aline = trim(lines[j]);
						if (startsWith(aline, "'")) {
							if (startsWith(aline, "'' @SDESCRIPTION:")) {
								description = trim(aline.replace("'' @SDESCRIPTION:", ""));
								//get all lines for the description, if the description is breaked up in new lines
								for (k = j + 1; k < i; k++) {
									addLine = trim(lines[k]);
									if (startsWith(addLine, "''") && !startsWith(addLine, "'' @")) {
										description += "\r" + trim(lines[k].replace("''", ""));
									} else {
										break;
									}
								}
							}
						} else {
							break;
						}
					}
					obsolete = (description.indexOf("OBSOLETE!") > -1);
					
					addElement("m", name, name, description, obsolete, classMembersNode, xml);
				} else if (startsWith(line, "property get ") || startsWith(line, "property let ")) {
					description = "";
					isLetProperty = startsWith(line, "property let ");
					line = line.replace("property get ", "");
					line = line.replace("property let ", "");
					propertyInfo = line.split("''");
					name = trim(propertyInfo[0]);
					if (propertyInfo.length > 0) {
						if (isLetProperty) {
							description = "SET: " + trim(propertyInfo[1]);
						} else {
							description = "GET: " + trim(propertyInfo[1]);
						}
					}
					obsolete = (description.indexOf("OBSOLETE!") > -1);
					//we replace the brackets in the name because some properties have this by mistake
					name = name.replace(/\(.*\)/, "");
					//we check if a get or set property is already here with this name.
					//when yes then we update its description
					pNode = classMembersNode.selectSingleNode("el/display[.='" + name + "']");
					if (pNode) {
						pNode.parentNode.selectSingleNode("help").text += "\r" + description;
					} else {
						addElement("p", name, name, description, obsolete, classMembersNode, xml);
					}
				} else {
					memberInfo = line.split("''");
					name = trim(memberInfo[0]);
					if (memberInfo.length > 0) {
						description = trim(memberInfo[1]);
						//we check if there are any comments for the membervariable in the next lines
						//so the comment is breaked into new lines...
						for (j = i + 1; j < lines.length; j++) {
							aline = trim(lines[j]);
							if (startsWith(aline, "''")) {
								description += "\r" + aline.replace("''", "");
							} else {
								break;
							}
						}
					}
					//common member variable
					obsolete = (description.indexOf("OBSOLETE!") > -1);
					addElement("p", name, name, description, obsolete, classMembersNode, xml);
				}
			}
		}
	}
}

function getNewNode(name, value, xml) {
	node = xml.createElement(name);
	node.appendChild(xml.createTextNode(value));
	return node;
}

function getGablibDir(app) {
	var configFilePath = app.AppPath + "\gabLibrarySettings.txt";
	fso = new ActiveXObject("Scripting.FileSystemObject");
	var configFile = fso.OpenTextFile(configFilePath, 1, true);
	var found = false;
	while (!configFile.AtEndOfStream)  {
		line = configFile.ReadLine();
		parts = line.split("|");
		if (parts[0] == "GablibDir" && parts[1] != "") {
			dir = parts[1];
			found = true;
			break;
		}
	}
	configFile.Close();
	
	if (!found) {
		var configFile = fso.OpenTextFile(configFilePath, 8, true);
		dir = "";
		while (dir.substring(dir.length -1, dir.length) != "\\") {
			dir = app.InputBox("Path", "Path of the GabLibrary on your Development-machine: (U:\\gab_Library\\)", "");
		}
		configFile.WriteLine("GablibDir|" + dir);
		configFile.Close();
	}
	
	configFile = null;
	fso = null;
	
	return dir;
}

function endsWith(str, val) {
	return (str.substring(str.length - val.length, str.length) == val);
}

function startsWith(str, val) {
	return (str.substring(0, val.length) == val);
}

function trim(str) {
	str += "";
   return str.replace(/^\s*|\s*$/g, "");
}

function capitalize(str) {
	return str.substring(0, 1).toUpperCase() + str.substring(1, str.length);
}

function addElement(type, display, insert, help, obsolete, toNode, xml) {
	anode = getNewNode("el", "", xml);
	anode.setAttribute("type", type);
	if (obsolete) anode.setAttribute("obsolete", 1);
	//some keywords are reserved in ASP, but when putting in brackets you can still use them. remove the brackets for display.
	d = display.replace("[", "");
	d = d.replace("]", "");
	anode.appendChild(getNewNode("display", d, xml));
	anode.appendChild(getNewNode("insert", insert, xml));
	anode.appendChild(getNewNode("help", help, xml));
	toNode.appendChild(anode);
}