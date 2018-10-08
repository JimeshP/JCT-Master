var mainHeight = $(".main_contain").height();

/**
 * Function calls when the page is loaded 
 * to populate the mapping values
 */
$(document).ready(function() {	
	document.getElementById('addHeaderDiv').innerHTML = "Add Strength";
	disableAllfield();
	fetchMappingList('S', 0, "STR");
	/***************** To Disable the paste functionality ***************/
  /*  $("#attrDescInptId").bind("paste",function(e) {
        e.preventDefault();
    });*/
});


/**
 * Function to view all the mapping values
 * from database while landing to the page
 * @param objval
 */
function fetchMappingList(objVal, profileId, relatedPage) {	
	var userProf = Spine.Model.sub();
	userProf.configure("/admin/contentconfig/populateExistingMapping", "profileId", "relatedPage");
	userProf.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userProf({  
		profileId: profileId,
		relatedPage: relatedPage
	});
	modelPopulator.save(); //POST
	userProf.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	userProf.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if(statusCode == 200) {
			//populateQtnDropDown(obj.noOfSubQtn);
			if(profileId == 0){
				populateDropDown(obj.userProfileMap);	
			}
					
			if( null == objVal|| objVal == "" || objVal == 'S') {
				populateStrengthTable(obj.existingMappingList, profileId);
			} else if (objVal == 'V') {
				populateValueTable(obj.existingMappingList, profileId);
			} else if(objVal == 'P') {
				populatePassionTable(obj.existingMappingList, profileId);
			} 
		} else if(statusCode == 201) {
			document.getElementById("existingMappingListId").innerHTML = "<div align='center'>"+obj.statusDesc+"</div>";
		} else if(statusCode == 500) {
			//Show error message
		}
	});
}

/**
 * Function to populate the Strength values 
 * on table view
 * @param mappingListData
 * @param action
 */
function populateStrengthTable(mappingListData, profileIdVal) {
	document.getElementById("existing_data_list").innerHTML = "Existing Strengths";
	var attributeOrder = 0;
	var tableDiv = document.getElementById("existingMappingListId");
	if (null != mappingListData) {
		var headingString = "<table width='94%' id='mappingTableSTRId' border=1 align='center' bordercolor='#78C0D3' class='tablesorter'><thead class='tab_header'><tr><th width='11%'><b>SL. No.</b></th><th width='15%'><b>Profile Name</b></th><th width='20%'><b>Attribute Name</b></th><th><b>Attribute Description</b></th><th width='12%'><b>Action</b></th></tr></thead><tbody>";
		var splitRec = mappingListData.split("$$$");
		if(splitRec.length <= 2){
			document.getElementById("chgOrderBtnId").disabled = true;
		}
		var counter = 1;
		for(var outerIndex = 0; outerIndex < splitRec.length-1; outerIndex++){						
			var objSplit = splitRec[outerIndex].split("~");
			var rowColor = "";			
			if(objSplit[0] == 'STR') {	
				if(counter%2 == 0){
					rowColor = "#F1F1F1";
				} else {
					rowColor = "#D2EAF0";
				}
				headingString = headingString + "<tr align='center' bgcolor='"+rowColor+"' class='row_width'><td>&nbsp; "+counter+".</td>";
				headingString = headingString + "<td align='center'>"+objSplit[4]+"</td>";
				//headingString = headingString + "<td align='center'>"+objSplit[6]+"</td>";
				headingString = headingString + "<td align='center'>&nbsp;&nbsp; "+objSplit[1]+"</td>";
				headingString = headingString + "<td align='center'>&nbsp;&nbsp; "+objSplit[5]+"</td>";
				//headingString = headingString + "<td class='table_col_txt_style'>&nbsp;&nbsp;&nbsp;<a class='delete_style' href='#' onClick='deleteStrength(\""+objSplit[2]+"\", \""+objSplit[1]+"\")'>Delete</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a class='edit_style' href='#' onClick='editStrength(\""+objSplit[2]+"\", \""+objSplit[1]+"\", \""+objSplit[4]+"\", \""+objSplit[5]+"\", \""+objSplit[6]+"\")'>Edit</a></td>";				
				//headingString = headingString + "<td class='table_col_txt_style'>&nbsp;&nbsp;&nbsp;<a class='delete_style' href='#' id='deleteSTR_"+counter+"' onClick='deleteStrength(\""+objSplit[2]+"\", \""+objSplit[1]+"\", \""+objSplit[7]+"\")'><img src = '../img/delete.png' /></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a class='edit_style' href='#' onClick='editStrength(\""+objSplit[2]+"\", \""+objSplit[1]+"\", \""+objSplit[4]+"\", \""+objSplit[5]+"\", \""+objSplit[7]+"\", \""+counter+"\")'><img src = '../img/edit.png' /></a></td>";	
				headingString = headingString + "<td class='table_col_txt_style' align='center'><a class='edit_style' href='#' onClick='editStrength(\""+objSplit[2]+"\", \""+objSplit[1]+"\", \""+objSplit[4]+"\", \""+objSplit[5]+"\", \""+objSplit[7]+"\", \""+counter+"\")'><img src = '../img/edit.png' /></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a class='delete_style' href='#' id='deleteSTR_"+counter+"' onClick='deleteStrength(\""+objSplit[2]+"\", \""+objSplit[1]+"\", \""+objSplit[7]+"\")'><img src = '../img/delete.png' /></a></td>";
				headingString = headingString + "</tr>";
				counter = counter + 1;
			}
			//alert(objSplit[6]);
			if(attributeOrder < objSplit[6]) {
				
				attributeOrder = objSplit[6];
				//alert("aaaa"+attributeOrder);
			}
		}
		/*if(profileIdVal != 0){
			document.getElementById('attrOrderInptId').value = parseInt(attributeOrder)+1;		
		}*/
		headingString = headingString + "</tbody></table>";
		tableDiv.innerHTML = headingString;
		//new SortableTable(document.getElementById('mappingTableSTRId'), 1);
		 $("table").tablesorter(); 
	} else {
		/*if(profileIdVal != 0){
			document.getElementById('attrOrderInptId').value = 1;		
		}*/
		document.getElementById("chgOrderBtnId").disabled = true;
		tableDiv.innerHTML = "<div align='center'><img src='../img/no-record.png'><br />No Mapping Attribute To Display</div>";
	}
}


/**
 * Function to populate the Values 
 * on table view
 * @param mappingListData
 * @param action
 */
function populateValueTable(mappingListData, profileIdVal) {
	document.getElementById("existing_data_list").innerHTML = "Existing Values";
	var attributeOrder = 0;
	var tableDiv = document.getElementById("existingMappingListId");
	if (null != mappingListData) {
		var headingString = "<table width='94%' id='mappingTableVALId' border=1 align='center' bordercolor='#78C0D3' class='tablesorter'><thead class='tab_header'><tr><th width='11%'><b>SL. No.</b></th><th width='15%'><b>Profile Name</b></th><th width='20%'><b>Attribute Name</b></th><th><b>Attribute Description</b></th><th width='12%'><b>Action</b></th></tr></thead><tbody>";
		var splitRec = mappingListData.split("$$$");
		if(splitRec.length <= 2){
			document.getElementById("chgOrderBtnId").disabled = true;
		}
		var counter = 1;
		for(var outerIndex = 0; outerIndex < splitRec.length-1; outerIndex++){						
			var objSplit = splitRec[outerIndex].split("~");
			var rowColor = "";			
			if(objSplit[0] == 'VAL') {	
				if(counter%2 == 0){
					rowColor = "#F1F1F1";
				} else {
					rowColor = "#D2EAF0";
				}
				headingString = headingString + "<tr align='center' bgcolor='"+rowColor+"' class='row_width'><td>&nbsp; "+counter+".</td>";
				headingString = headingString + "<td align='center'>"+objSplit[4]+"</td>";
				//headingString = headingString + "<td align='center'>"+objSplit[6]+"</td>";
				headingString = headingString + "<td align='center'>&nbsp;&nbsp; "+objSplit[1]+"</td>";
				headingString = headingString + "<td align='center'>&nbsp;&nbsp; "+objSplit[5]+"</td>";
				//headingString = headingString + "<td class='table_col_txt_style'>&nbsp;&nbsp;&nbsp;<a class='delete_style' href='#' id='deleteVAL_"+counter+"' onClick='deleteValue(\""+objSplit[2]+"\", \""+objSplit[1]+"\", \""+objSplit[7]+"\")'><img src = '../img/delete.png' /></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a class='edit_style' href='#' onClick='editValue(\""+objSplit[2]+"\", \""+objSplit[1]+"\", \""+objSplit[4]+"\", \""+objSplit[5]+"\", \""+objSplit[7]+"\", \""+counter+"\")'><img src = '../img/edit.png' /></a></td>";								
				headingString = headingString + "<td class='table_col_txt_style' align='center'><a class='edit_style' href='#' onClick='editValue(\""+objSplit[2]+"\", \""+objSplit[1]+"\", \""+objSplit[4]+"\", \""+objSplit[5]+"\", \""+objSplit[7]+"\", \""+counter+"\")'><img src = '../img/edit.png' /></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a class='delete_style' href='#' id='deleteVAL_"+counter+"' onClick='deleteValue(\""+objSplit[2]+"\", \""+objSplit[1]+"\", \""+objSplit[7]+"\")'><img src = '../img/delete.png' /></a></td>";
				headingString = headingString + "</tr>";
				counter = counter + 1;
			}
			if(attributeOrder < objSplit[6]) {
				attributeOrder = objSplit[6];
			}
		}
		/*if(profileIdVal != 0){
			document.getElementById('attrOrderInptId').value = parseInt(attributeOrder)+1;		
		}*/
		headingString = headingString + "</tbody></table>";
		tableDiv.innerHTML = headingString;
		//new SortableTable(document.getElementById('mappingTableVALId'), 1);
		 $("table").tablesorter(); 
	} else {
		/*if(profileIdVal != 0){
			document.getElementById('attrOrderInptId').value = 1;		
		}*/
		document.getElementById("chgOrderBtnId").disabled = true;
		tableDiv.innerHTML = "<div align='center'><img src='../img/no-record.png'><br />No Mapping Attribute To Display</div>";
	}
}


/**
 * Function to populate the passion values 
 * on table view
 * @param mappingListData
 * @param action
 */
function populatePassionTable(mappingListData, profileIdVal) {
	document.getElementById("existing_data_list").innerHTML = "Existing Passions";
	var attributeOrder = 0;
	var tableDiv = document.getElementById("existingMappingListId");
	if (null != mappingListData) {
		var headingString = "<table width='94%' id='mappingTablePASId' border=1 align='center' bordercolor='#78C0D3' class='tablesorter'><thead class='tab_header'><tr><th width='11%'><b>SL. No.</b></th><th width='15%'><b>Profile Name</b></th><th width='20%'><b>Attribute Name</b></th><th><b>Attribute Description</b></th><th width='12%'><b>Action</b></th></tr></thead><tbody>";
		var splitRec = mappingListData.split("$$$");
		if(splitRec.length <= 2){
			document.getElementById("chgOrderBtnId").disabled = true;
		}
		var counter = 1;
		for(var outerIndex = 0; outerIndex < splitRec.length-1; outerIndex++){						
			var objSplit = splitRec[outerIndex].split("~");
			var rowColor = "";
			if(objSplit[0] == 'PAS') {	
				if(counter%2 == 0){
					rowColor = "#F1F1F1";
				} else {
					rowColor = "#D2EAF0";
				}
				headingString = headingString + "<tr align='center' bgcolor='"+rowColor+"' class='row_width'><td>&nbsp; "+counter+".</td>";
				headingString = headingString + "<td align='center'>"+objSplit[4]+"</td>";
				//headingString = headingString + "<td align='center'>"+objSplit[6]+"</td>";
				headingString = headingString + "<td align='center'>&nbsp;&nbsp; "+objSplit[1]+"</td>";
				headingString = headingString + "<td align='center'>&nbsp;&nbsp; "+objSplit[5]+"</td>";
				headingString = headingString + "<td class='table_col_txt_style' align='center'><a class='edit_style' href='#' onClick='editPassion(\""+objSplit[2]+"\", \""+objSplit[1]+"\", \""+objSplit[4]+"\", \""+objSplit[5]+"\", \""+objSplit[7]+"\", \""+counter+"\")'><img src = '../img/edit.png' /></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a class='delete_style' href='#' id='deleteVAL_"+counter+"' onClick='deletePassion(\""+objSplit[2]+"\", \""+objSplit[1]+"\", \""+objSplit[7]+"\")'><img src = '../img/delete.png' /></a></td>";				
				headingString = headingString + "</tr>";
				counter = counter + 1;
			}
			if(attributeOrder < objSplit[6]) {
				attributeOrder = objSplit[6];
			}
		}
		/*if(profileIdVal != 0){
			document.getElementById('attrOrderInptId').value = parseInt(attributeOrder)+1;		
		}*/
		headingString = headingString + "</tbody></table>";
		tableDiv.innerHTML = headingString;
		//new SortableTable(document.getElementById('mappingTablePASId'), 1);
		$("table").tablesorter(); 
	} else {
		/*if(profileIdVal != 0){
			document.getElementById('attrOrderInptId').value = 1;		
		}*/
		document.getElementById("chgOrderBtnId").disabled = true;
		tableDiv.innerHTML = "<div align='center'><img src='../img/no-record.png'><br />No Mapping Attribute To Display</div>";
	}

}


/**
 * Function add to check the  
 * selected radio button
 * @param obj
 */
function checkSelectedOption(obj) {	
	var strengthAddDiv = document.getElementById('strengthAddDiv');
	var valueAddDiv = document.getElementById('valueAddDiv');
	var passionAddDiv = document.getElementById('passionAddDiv');
	var textLabel = document.getElementById('addHeaderDiv');
	if(obj.value == 'S'){
				
		strengthAddDiv.setAttribute("style", "display:block");	
		valueAddDiv.setAttribute("style", "display:none");
		passionAddDiv.setAttribute("style", "display:none");
		textLabel.innerHTML = "Add Strength";
		document.getElementById("addStrengthBtnId").value = "Add";
		//fetchMappingList(obj);
		disableAllfield();
		fetchMappingList('S', 0, "STR");
		
	} else if(obj.value == 'V') {		
		
		strengthAddDiv.setAttribute("style", "display:none");	
		valueAddDiv.setAttribute("style", "display:block");
		passionAddDiv.setAttribute("style", "display:none");
		textLabel.innerHTML = "Add Value";
		document.getElementById("addValueBtnId").value = "Add";		
		//fetchMappingList(obj);
		disableAllfield();
		fetchMappingList('V', 0, "VAL");
	} else {
		strengthAddDiv.setAttribute("style", "display:none");	
		valueAddDiv.setAttribute("style", "display:none");
		passionAddDiv.setAttribute("style", "display:block");
		textLabel.innerHTML = "Add Passion";
		document.getElementById("addPassionBtnId").value = "Add";
		//fetchMappingList(obj);
		disableAllfield();
		fetchMappingList('P', 0, "PAS");
	}
	
}

/**
 * Function add to populate
 * the user profile drop-down value
 * @param userProfileMap
 */
function populateDropDown(userProfileMap){
	document.getElementById("inputUserProfileInpt").innerHTML = "";	
	var userProfileSelect = document.getElementById("inputUserProfileInpt");
	var defaultOpn = document.createElement("option");
	defaultOpn.text = "Select User Profile";
	defaultOpn.value = "";
	defaultOpn.className = "form-control-general";
	userProfileSelect.add(defaultOpn, null);

	for (var key in userProfileMap) {				
		var option = document.createElement("option");
		option.text = userProfileMap[key];
	    option.value = key;
	    option.className = "form-control-general";
	    try {
	    	userProfileSelect.add(option, null); //Standard 
	    }catch(error) {
	    	//regionSelect.add(option); // IE only
	    }
	}
}

/**
 * Function to add the
 * Strengths to the database
 * @param null
 */
function saveStrength() {
	var attrDescText = document.getElementById("attrDescInptId").value;
	var attrFullDescText = document.getElementById("attrFullDescInptId").value;
	//var attributeOrder = document.getElementById("attrOrderInptId").value;
	var userProfile = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value;
	var userProfileText = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].text;
	if(validateValue(attrDescText, attrFullDescText, "STR")) {
		var newUserProf = Spine.Model.sub();
		newUserProf.configure("/admin/contentconfig/saveStrength", "attrDescTextVal", "userProfileVal", "userProfileId", "attrFullDescText" ,"createdBy");
		newUserProf.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new newUserProf({  
			attrDescTextVal: attrDescText,
			userProfileVal: userProfileText,
			userProfileId: userProfile,
			attrFullDescText: attrFullDescText,
			createdBy: sessionStorage.getItem("jctEmail")
		});
		modelPopulator.save(); //POST
		newUserProf.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
		});
		newUserProf.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			var statusCode = obj.statusCode;
			if(statusCode == 200) {
				//alertify.set({ delay: 2700 });
				alertify.alert("Strength has been successfully created.");
				populateStrengthTable(obj.existingMappingList, "ADD");	
				document.getElementById("attrDescInptId").value = "";
				document.getElementById("attrFullDescInptId").value = "";
				//disableAllfield();
			} else if(statusCode == 600) {
				alertify.alert("Strength already exists.");
			}  else if(statusCode == 602) {
				alertify.alert("One profile should have maximum of 24 Strengths.");
			}  else if(statusCode == 500) {
				// Error Message
			}
		});
	}
}


/**
 * Function to add the
 * Values to the database
 * @param null
 */
function saveValue() {
	var attrDescText = document.getElementById("attrDescInptId").value;
	var attrFullDescText = document.getElementById("attrFullDescInptId").value;
	//var attributeOrder = document.getElementById("attrOrderInptId").value;
	var userProfile = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value;
	var userProfileText = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].text;
	if(validateValue(attrDescText, attrFullDescText, "VAL")) {
		var newUserProf = Spine.Model.sub();
		newUserProf.configure("/admin/contentconfig/saveValue", "attrDescTextVal", "userProfileVal", "userProfileId", "attrFullDescText", "createdBy");
		newUserProf.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new newUserProf({  
			attrDescTextVal: attrDescText,
			userProfileVal: userProfileText,
			userProfileId: userProfile,
			attrFullDescText: attrFullDescText,			
			createdBy: sessionStorage.getItem("jctEmail")
		});
		modelPopulator.save(); //POST
		newUserProf.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
		});
		newUserProf.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			var statusCode = obj.statusCode;
			if(statusCode == 200) {
				//alertify.set({ delay: 2700 });
				alertify.alert("Value has been successfully created.");
				populateValueTable(obj.existingMappingList, "ADD");		
				document.getElementById("attrDescInptId").value = "";
				document.getElementById("attrFullDescInptId").value = "";
				//disableAllfield();
			} else if(statusCode == 600) {
				alertify.alert("Value already exists.");
			} else if(statusCode == 602) {
				alertify.alert("One profile should have maximum of 24 Values.");
			} else if(statusCode == 500) {
				// Error Message
			}
		});
	}
}

/**
 * Function to add the
 * passion to the database
 * @param null
 */
function savePassion() {
	var attrDescText = document.getElementById("attrDescInptId").value;
	var attrFullDescText = document.getElementById("attrFullDescInptId").value;
	//var attributeOrder = document.getElementById("attrOrderInptId").value;
	var userProfile = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value;
	var userProfileText = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].text;
	if(validateValue(attrDescText, attrFullDescText, "PAS")) {
		var newUserProf = Spine.Model.sub();
		newUserProf.configure("/admin/contentconfig/savePassion", "attrDescTextVal", "userProfileVal", "userProfileId", "attrFullDescText", "createdBy");
		newUserProf.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new newUserProf({  
			attrDescTextVal: attrDescText,
			userProfileVal: userProfileText,
			userProfileId: userProfile,
			attrFullDescText: attrFullDescText,
			createdBy: sessionStorage.getItem("jctEmail")
		});
		modelPopulator.save(); //POST
		newUserProf.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
		});
		newUserProf.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			var statusCode = obj.statusCode;
			if(statusCode == 200) {
				//alertify.set({ delay: 2700 });
				alertify.alert("Passion has been successfully created.");
				populatePassionTable(obj.existingMappingList, "ADD");	
				document.getElementById("attrDescInptId").value = "";
				document.getElementById("attrFullDescInptId").value = "";
				//disableAllfield();
			} else if(statusCode == 600) {
				alertify.alert("Passion already exists.");
			} else if(statusCode == 602) {
				alertify.alert("One profile should have maximum of 24 Passions.");
			} else if(statusCode == 500) {
				// Error Message
			}
		});
	}
}

/**
 * Function to edit the strength
 * @param tablePkId
 * @param questnDesc
 * @param dropDownVal
 */
function editStrength(tablePkId, attrDesc, dropDownVal, attrFullDesc, profileId, counter){
	fetchMappingList('S', profileId, "STR");
	/******* ADDED FOR JCT PUBLIC VERSION ********/
	if(dropDownVal == "Default Profile") {
		var globalProfileChgDiv = document.getElementById('globalProfileChgDiv');
		globalProfileChgDiv.setAttribute("style", "display:block");	
	}	
	/********** ENDED **********/
	$("#strengthAddDiv").addClass("right_allgn_edit");
	var chgOrderDiv = document.getElementById("chgOrderDiv");	
	chgOrderDiv.setAttribute("style", "display:none");
	$("#deleteSTR_"+counter).addClass('disable_delete');
	enableAllfield();
	document.getElementById('attrDescInptId').value = attrDesc;
	document.getElementById('attrFullDescInptId').value = attrFullDesc;
	document.getElementsByName("mappingValFld")[0].disabled = true;
	document.getElementsByName("mappingValFld")[1].disabled = true;
	document.getElementsByName("mappingValFld")[2].disabled = true;
	document.getElementById('inputUserProfileInpt').disabled = true;		
	//document.getElementById('qtnDescInptId').setAttribute("readonly", "true");
	document.getElementById('addHeaderDiv').innerHTML = "Update Strength";	
	document.getElementById("updateId").innerHTML = "Update Mapping Values";
	document.getElementById("addStrengthBtnId").value = "Update";
	document.getElementById("addStrengthBtnId").setAttribute("onclick", "updateStrength("+tablePkId+",'"+attrDesc+"', '"+attrFullDesc+"')");
	var updateSection = document.getElementById("createMappingValUpdate");
	updateSection.className = updateSection.className + " otherclass";
	//populateUserProfile
	var usrGrp = Spine.Model.sub();
	usrGrp.configure("/admin/contentconfig/populateUserProfileRefQtn", "none");
	usrGrp.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new usrGrp({  
		none: ""
	});
	modelPopulator.save(); //POST
	usrGrp.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	usrGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if(statusCode == 200) {
			var userProfileSelect = document.getElementById("inputUserProfileInpt");
			userProfileSelect.innerHTML = "";
			var defaultOpn = document.createElement("option");
			defaultOpn.text = "Select User Profile";
			defaultOpn.value = "";
			defaultOpn.className = "form-control-general";
			userProfileSelect.add(defaultOpn, null);
			for (var key in obj.userProfileMap) {				
				var option = document.createElement("option");
				option.text = obj.userProfileMap[key];
			    option.value = key;
			    option.className = "form-control-general";
			    if(option.text == dropDownVal) {
					option.setAttribute("selected", "selected");
				}
			    try {
			    	userProfileSelect.add(option, null); //Standard 
			    }catch(error) {
			    	//regionSelect.add(option); // IE only
			    }
			}
		} else if(statusCode == 500) {
			// Error Message
		}
	});
}


/**
 * Function to update the
 * strength values
 * @param tablePkId
 * @param prevVal
 */
function updateStrength(tablePkId, prevVal, prevValDesc) {
	var atttrDesc = document.getElementById("attrDescInptId").value;
	var attrFullDesc = document.getElementById("attrFullDescInptId").value;
	var userProfile = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value;
	var userProfileText = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].text;
	var chechBoxVal = $("#globalProfileCheckBoxId").is(":checked");
	
	var cnfrmMsg = "";
	if (chechBoxVal == true) {
		cnfrmMsg = "Are you sure you want to make global profile change?";
	} else {
		cnfrmMsg = "Are you sure you want to update?";
	}
	alertify.set({ buttonReverse: true });
	alertify.confirm(cnfrmMsg, function (e) {
		if (e) {
			if(validateValue(atttrDesc, attrFullDesc)) {
				var newUserProf = Spine.Model.sub();
				newUserProf.configure("/admin/contentconfig/updateStrength", "atttrDescVal", "userProfileVal", "userProfileId", "tablePkId", "attrFullDesc", "chechBoxVal");
				newUserProf.extend( Spine.Model.Ajax );
				//Populate the model with data to transfer
				var modelPopulator = new newUserProf({  
					//userGroupVal: userGroup,
					//reflectnQtnVal: prevVal,
					atttrDescVal: atttrDesc,
					userProfileVal: userProfileText,
					userProfileId: userProfile,
					tablePkId: tablePkId,
					attrFullDesc: attrFullDesc,
					chechBoxVal: chechBoxVal
				});
				modelPopulator.save(); //POST
				newUserProf.bind("ajaxError", function(record, xhr, settings, error) {
					alertify.alert("Unable to connect to the server.");
				});
				newUserProf.bind("ajaxSuccess", function(record, xhr, settings, error) {
					var jsonStr = JSON.stringify(xhr);
					var obj = jQuery.parseJSON(jsonStr);
					var statusCode = obj.statusCode;
					if(statusCode == 200) {
						
						//alertify.alert("Strength has been edited successfully.");
						alertify.alert(obj.statusDesc);
						//populateStrengthTable(obj.existingMappingList, "UPDATE");
						fetchMappingList('S', userProfile, "STR");
						
						/** Attribute Description set null **/
						document.getElementById('attrDescInptId').value = "";	
						document.getElementById('attrFullDescInptId').value = "";	
						/** Mapping Value radio button become enable **/
						document.getElementsByName("mappingValFld")[0].disabled = false;
						document.getElementsByName("mappingValFld")[1].disabled = false;
						document.getElementsByName("mappingValFld")[2].disabled = false;				
						/** User Profile dropdown become enable and set null **/
						document.getElementById('inputUserProfileInpt').disabled = false;	
						/*document.getElementById('inputUserProfileInpt').value = "";	*/
						/*** Header description changed **/
						document.getElementById("updateId").innerHTML = "Mapping Values";
						document.getElementById('addHeaderDiv').innerHTML = "Add Strength";	
						/*** Button description changed **/
						document.getElementById("addStrengthBtnId").value = "Add";
						document.getElementById("addStrengthBtnId").setAttribute("onclick", "saveStrength()");
						/*** remove the class **/
						var updateSection = document.getElementById("createMappingValUpdate");
						updateSection.className = updateSection.className - " otherclass";
						
						var chgOrderDiv = document.getElementById("chgOrderDiv");	
						chgOrderDiv.setAttribute("style", "display:block");
						document.getElementById("chgOrderBtnId").disabled = false;
						
						/******* ADDED FOR JCT PUBLIC VERSION ********/
						var globalProfileChgDiv = document.getElementById('globalProfileChgDiv');
						globalProfileChgDiv.setAttribute("style", "display:none");	
						
					} else if(statusCode == 600) {
						alertify.alert("Strength already exists.");
					} else if(statusCode == 500) {
						fetchMappingList('S', userProfile, "STR");
						alertify.alert(obj.statusDesc);
					}
				});
			}			
		}
	});		
}


/**
 * Function to edit the values
 * @param tablePkId
 * @param questnDesc
 * @param dropDownVal
 */
//function editValue(tablePkId, attrDesc, dropDownVal, attrFullDesc, attrOrder){
function editValue(tablePkId, attrDesc, dropDownVal, attrFullDesc, profileId, counter){
	fetchMappingList('V', profileId, "VAL");
	if(dropDownVal == "Default Profile") {
		var globalProfileChgDiv = document.getElementById('globalProfileChgDiv');
		globalProfileChgDiv.setAttribute("style", "display:block");	
	}
	$("#valueAddDiv").addClass("right_allgn_edit");
	var chgOrderDiv = document.getElementById("chgOrderDiv");	
	chgOrderDiv.setAttribute("style", "display:none");
	$("#deleteVAL_"+counter).addClass('disable_delete');
	
	enableAllfield();
	document.getElementById('attrDescInptId').value = attrDesc;
	document.getElementById('attrFullDescInptId').value = attrFullDesc;
	document.getElementsByName("mappingValFld")[0].disabled = true;
	document.getElementsByName("mappingValFld")[1].disabled = true;
	document.getElementsByName("mappingValFld")[2].disabled = true;
	document.getElementById('inputUserProfileInpt').disabled = true;		
	//document.getElementById('qtnDescInptId').setAttribute("readonly", "true");
	document.getElementById('addHeaderDiv').innerHTML = "Update Value";	
	document.getElementById("updateId").innerHTML = "Update Mapping Values";
	document.getElementById("addValueBtnId").value = "Update";
	document.getElementById("addValueBtnId").setAttribute("onclick", "updateValue("+tablePkId+",'"+attrDesc+"', '"+attrFullDesc+"')");
	var updateSection = document.getElementById("createMappingValUpdate");
	updateSection.className = updateSection.className + " otherclass";
	//populateUserProfile
	var usrGrp = Spine.Model.sub();
	usrGrp.configure("/admin/contentconfig/populateUserProfileRefQtn", "none");
	usrGrp.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new usrGrp({  
		none: ""
	});
	modelPopulator.save(); //POST
	usrGrp.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	usrGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if(statusCode == 200) {
			var userProfileSelect = document.getElementById("inputUserProfileInpt");
			userProfileSelect.innerHTML = "";
			var defaultOpn = document.createElement("option");
			defaultOpn.text = "Select User Profile";
			defaultOpn.value = "";
			defaultOpn.className = "form-control-general";
			userProfileSelect.add(defaultOpn, null);
			for (var key in obj.userProfileMap) {				
				var option = document.createElement("option");
				option.text = obj.userProfileMap[key];
			    option.value = key;
			    option.className = "form-control-general";
			    if(option.text == dropDownVal) {
					option.setAttribute("selected", "selected");
				}
			    try {
			    	userProfileSelect.add(option, null); //Standard 
			    }catch(error) {
			    	//regionSelect.add(option); // IE only
			    }
			}
		} else if(statusCode == 500) {
			// Error Message
		}
	});
}


/**
 * Function to update the
 * value attributes
 * @param tablePkId
 * @param prevVal
 */
function updateValue(tablePkId, prevVal, prevDescVal) {
	var atttrDesc = document.getElementById("attrDescInptId").value;
	var attrFullDesc = document.getElementById("attrFullDescInptId").value;
	var userProfile = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value;
	var userProfileText = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].text;
	var chechBoxVal = $("#globalProfileCheckBoxId").is(":checked");
	
	var cnfrmMsg = "";
	if (chechBoxVal == true) {
		cnfrmMsg = "Are you sure you want to make global profile change?";
	} else {
		cnfrmMsg = "Are you sure you want to update?";
	}
	alertify.set({ buttonReverse: true });
	alertify.confirm(cnfrmMsg, function (e) {
		if (e) {			
			if(validateValue(atttrDesc, attrFullDesc)) {
				var newUserProf = Spine.Model.sub();
				newUserProf.configure("/admin/contentconfig/updateValue", "atttrDescVal", "userProfileVal", "userProfileId", "tablePkId" , "attrFullDesc" , "chechBoxVal");
				newUserProf.extend( Spine.Model.Ajax );
				//Populate the model with data to transfer
				var modelPopulator = new newUserProf({  
					//userGroupVal: userGroup,
					//reflectnQtnVal: prevVal,
					atttrDescVal: atttrDesc,
					userProfileVal: userProfileText,
					userProfileId: userProfile,
					tablePkId: tablePkId,
					attrFullDesc: attrFullDesc,
					chechBoxVal: chechBoxVal 
				});
				modelPopulator.save(); //POST
				newUserProf.bind("ajaxError", function(record, xhr, settings, error) {
					alertify.alert("Unable to connect to the server.");
				});
				newUserProf.bind("ajaxSuccess", function(record, xhr, settings, error) {
					var jsonStr = JSON.stringify(xhr);
					var obj = jQuery.parseJSON(jsonStr);
					var statusCode = obj.statusCode;
					if(statusCode == 200) {
						alertify.alert(obj.statusDesc);
						//alertify.alert("Value has been edited successfully.");
						//populateValueTable(obj.existingMappingList, "UPDATE");
						fetchMappingList('V', userProfile, "VAL");
						
						/** Attribute Description set null **/
						document.getElementById('attrDescInptId').value = "";		
						document.getElementById('attrFullDescInptId').value = "";		
						/** Mapping Value radio button become enable **/
						document.getElementsByName("mappingValFld")[0].disabled = false;
						document.getElementsByName("mappingValFld")[1].disabled = false;
						document.getElementsByName("mappingValFld")[2].disabled = false;				
						/** User Profile dropdown become enable and set null **/
						document.getElementById('inputUserProfileInpt').disabled = false;	
						/*document.getElementById('inputUserProfileInpt').value = "";	*/
						/*** Header description changed **/
						document.getElementById("updateId").innerHTML = "Mapping Values";
						document.getElementById('addHeaderDiv').innerHTML = "Add Value";	
						/*** Button description changed **/
						document.getElementById("addValueBtnId").value = "Add";
						document.getElementById("addValueBtnId").setAttribute("onclick", "saveValue()");
						/*** remove the class **/
						var updateSection = document.getElementById("createMappingValUpdate");
						updateSection.className = updateSection.className - " otherclass";
						
						var chgOrderDiv = document.getElementById("chgOrderDiv");	
						chgOrderDiv.setAttribute("style", "display:block");
						document.getElementById("chgOrderBtnId").disabled = false;
						
						/******* ADDED FOR JCT PUBLIC VERSION ********/
						var globalProfileChgDiv = document.getElementById('globalProfileChgDiv');
						globalProfileChgDiv.setAttribute("style", "display:none");	
						
					} else if(statusCode == 600) {
						alertify.alert("Value already exists.");
					} else if(statusCode == 500) {
						fetchMappingList('V', userProfile, "VAL");
						alertify.alert(obj.statusDesc);
					}
				});
			}
		}
	});	
}


/**
 * Function to edit the passion attributes
 * @param tablePkId
 * @param questnDesc
 * @param dropDownVal
 */
function editPassion(tablePkId, attrDesc, dropDownVal, attrFullDesc, profileId, counter){
	fetchMappingList('P', profileId, "PAS");
	if(dropDownVal == "Default Profile") {
		var globalProfileChgDiv = document.getElementById('globalProfileChgDiv');
		globalProfileChgDiv.setAttribute("style", "display:block");	
	}
	$("#passionAddDiv").addClass("right_allgn_edit");
	var chgOrderDiv = document.getElementById("chgOrderDiv");	
	chgOrderDiv.setAttribute("style", "display:none");
	$("#deletePAS_"+counter).addClass('disable_delete');
	
	enableAllfield();
	document.getElementById('attrDescInptId').value = attrDesc;
	document.getElementById('attrFullDescInptId').value = attrFullDesc;
	document.getElementsByName("mappingValFld")[0].disabled = true;
	document.getElementsByName("mappingValFld")[1].disabled = true;
	document.getElementsByName("mappingValFld")[2].disabled = true;
	document.getElementById('inputUserProfileInpt').disabled = true;		
	//document.getElementById('qtnDescInptId').setAttribute("readonly", "true");
	document.getElementById('addHeaderDiv').innerHTML = "Update Passion";	
	document.getElementById("updateId").innerHTML = "Update Mapping Values";
	document.getElementById("addPassionBtnId").value = "Update";
	document.getElementById("addPassionBtnId").setAttribute("onclick", "updatePassion("+tablePkId+",'"+attrDesc+"', '"+attrFullDesc+"')");
	var updateSection = document.getElementById("createMappingValUpdate");
	updateSection.className = updateSection.className + " otherclass";
	//populateUserProfile
	var usrGrp = Spine.Model.sub();
	usrGrp.configure("/admin/contentconfig/populateUserProfileRefQtn", "none");
	usrGrp.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new usrGrp({  
		none: ""
	});
	modelPopulator.save(); //POST
	usrGrp.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	usrGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if(statusCode == 200) {
			var userProfileSelect = document.getElementById("inputUserProfileInpt");
			userProfileSelect.innerHTML = "";
			var defaultOpn = document.createElement("option");
			defaultOpn.text = "Select User Profile";
			defaultOpn.value = "";
			defaultOpn.className = "form-control-general";
			userProfileSelect.add(defaultOpn, null);
			for (var key in obj.userProfileMap) {				
				var option = document.createElement("option");
				option.text = obj.userProfileMap[key];
			    option.value = key;
			    option.className = "form-control-general";
			    if(option.text == dropDownVal) {
					option.setAttribute("selected", "selected");
				}
			    try {
			    	userProfileSelect.add(option, null); //Standard 
			    }catch(error) {
			    	//regionSelect.add(option); // IE only
			    }
			}
		} else if(statusCode == 500) {
			// Error Message
		}
	});
}

/**
 * Function to update the
 * passion attributes
 * @param tablePkId
 * @param prevVal
 */
function updatePassion(tablePkId, prevVal, prevDescVal) {
	var atttrDesc = document.getElementById("attrDescInptId").value;
	var attrFullDesc = document.getElementById("attrFullDescInptId").value;
	var userProfile = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value;
	var userProfileText = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].text;
	var chechBoxVal = $("#globalProfileCheckBoxId").is(":checked");
	var cnfrmMsg = "";
	if (chechBoxVal == true) {
		cnfrmMsg = "Are you sure you want to make global profile change?";
	} else {
		cnfrmMsg = "Are you sure you want to update?";
	}
	alertify.set({ buttonReverse: true });
	alertify.confirm(cnfrmMsg, function (e) {
		if (e) {
			if(validateValue(atttrDesc, attrFullDesc)) {
				var newUserProf = Spine.Model.sub();
				newUserProf.configure("/admin/contentconfig/updatePassion", "atttrDescVal", "userProfileVal", "userProfileId", "tablePkId" , "attrFullDesc", "chechBoxVal");
				newUserProf.extend( Spine.Model.Ajax );
				//Populate the model with data to transfer
				var modelPopulator = new newUserProf({  
					//userGroupVal: userGroup,
					//reflectnQtnVal: prevVal,
					atttrDescVal: atttrDesc,
					userProfileVal: userProfileText,
					userProfileId: userProfile,
					tablePkId: tablePkId,
					attrFullDesc: attrFullDesc,
					chechBoxVal: chechBoxVal
				});
				modelPopulator.save(); //POST
				newUserProf.bind("ajaxError", function(record, xhr, settings, error) {
					alertify.alert("Unable to connect to the server.");
				});
				newUserProf.bind("ajaxSuccess", function(record, xhr, settings, error) {
					var jsonStr = JSON.stringify(xhr);
					var obj = jQuery.parseJSON(jsonStr);
					var statusCode = obj.statusCode;
					if(statusCode == 200) {
						alertify.alert(obj.statusDesc);
						//alertify.alert("Passion has been edited successfully.");
						//populatePassionTable(obj.existingMappingList, "UPDATE");
						fetchMappingList('P', userProfile, "PAS");
						
						/** Attribute Description set null **/
						document.getElementById('attrDescInptId').value = "";		
						document.getElementById('attrFullDescInptId').value = "";		
						/** Mapping Value radio button become enable **/
						document.getElementsByName("mappingValFld")[0].disabled = false;
						document.getElementsByName("mappingValFld")[1].disabled = false;
						document.getElementsByName("mappingValFld")[2].disabled = false;				
						/** User Profile dropdown become enable and set null **/
						document.getElementById('inputUserProfileInpt').disabled = false;	
						/*document.getElementById('inputUserProfileInpt').value = "";	*/
						/*** Header description changed **/
						document.getElementById("updateId").innerHTML = "Mapping Values";
						document.getElementById('addHeaderDiv').innerHTML = "Add Passion";	
						/*** Button description changed **/
						document.getElementById("addPassionBtnId").value = "Add";
						document.getElementById("addPassionBtnId").setAttribute("onclick", "savePassion()");
						/*** remove the class **/
						var updateSection = document.getElementById("createMappingValUpdate");
						updateSection.className = updateSection.className - " otherclass";
						
						var chgOrderDiv = document.getElementById("chgOrderDiv");	
						chgOrderDiv.setAttribute("style", "display:block");
						document.getElementById("chgOrderBtnId").disabled = false;
						
						/******* ADDED FOR JCT PUBLIC VERSION ********/
						var globalProfileChgDiv = document.getElementById('globalProfileChgDiv');
						globalProfileChgDiv.setAttribute("style", "display:none");	
					} else if(statusCode == 600) {
						alertify.alert("Passion already exists.");
					} else if(statusCode == 500) {
						fetchMappingList('P', userProfile, "PAS");
						alertify.alert(obj.statusDesc);
					}
				});
			}			
		}
	});			
}


/**
 * Function to validate the
 * attribute description text
 * @param input
 * @returns {Boolean}
 */
function validateValue(attrNameText, attrDescText){
	var re=/^\s*$/;
	if(document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value == ""){
		alertify.alert("Please select User Profile.");
		return false;
	}
	if (re.test(attrNameText)){
		alertify.alert("Please enter attribute name.");
		return false;
	} else if(attrNameText.search(/[.,]{2,}/) > 0) {
		alertify.alert("Consecutive special character is not allowed");
		return false;
	}
	if (re.test(attrDescText)){
		alertify.alert("Please enter attribute description.");
		return false;
	} else if(attrDescText.search(/[.,]{2,}/) > 0) {
		alertify.alert("Consecutive special character is not allowed");
		return false;
	}
	
	/****** Validate attribute order******/
/*	if (re.test(attributeOrder)){
		alertify.alert("Please enter attribute order.");
		return false;
	} else if(attributeOrder == 0){
		alertify.alert("Attribute order cannot be 0.");
		return false;
	} else if(relatedPage == "STR"){
		var tbl = document.getElementById("mappingTableSTRId");
		var numRows = tbl.rows.length;
		var array = new Array();
		for (var i = 1; i < numRows; i++) {
		    var cells = tbl.rows[i].getElementsByTagName('td');
		    array.push(cells[2].innerHTML);
		}
		    if(array.indexOf(attributeOrder) > -1){		    	
		    	alertify.alert("Attribute order already exists.");
				return false;
		    }
		
	} else if (relatedPage == "VAL") {
		var tbl = document.getElementById("mappingTableVALId");
		var numRows = tbl.rows.length;
		var array = new Array();
		for (var i = 1; i < numRows; i++) {
		    var cells = tbl.rows[i].getElementsByTagName('td');
		    array.push(cells[2].innerHTML);
		}
		    if(array.indexOf(attributeOrder) > -1){		    	
		    	alertify.alert("Attribute order already exists.");
				return false;
		}
	} else if (relatedPage == "PAS") {
		var tbl = document.getElementById("mappingTablePASId");
		var numRows = tbl.rows.length;
		var array = new Array();
		for (var i = 1; i < numRows; i++) {
		    var cells = tbl.rows[i].getElementsByTagName('td');
		    array.push(cells[2].innerHTML);
		}
		    if(array.indexOf(attributeOrder) > -1){		    	
		    	alertify.alert("Attribute order already exists.");
				return false;
		}
	}*/
	return true;
}


/**
 * Function to validate the
 * attribute description text
 * @param input
 * @returns {Boolean}
 */
function validateValueForUpdate(attrNameText, attrDescText, attributeOrder){
	var re=/^\s*$/;
	if(document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value == ""){
		alertify.alert("Please select User Profile.");
		return false;
	}
	if (re.test(attrNameText)){
		alertify.alert("Please enter attribute name.");
		return false;
	} else if(attrNameText.search(/[.,]{2,}/) > 0) {
		alertify.alert("Consecutive special character is not allowed");
		return false;
	}
	if (re.test(attrDescText)){
		alertify.alert("Please enter attribute description.");
		return false;
	} else if(attrDescText.search(/[.,]{2,}/) > 0) {
		alertify.alert("Consecutive special character is not allowed");
		return false;
	}
	
	/****** Validate attribute order******/
	if (re.test(attributeOrder)){
		alertify.alert("Please enter attribute order.");
		return false;
	} else if(attributeOrder == 0){
		alertify.alert("Attribute order cannot be 0.");
		return false;
	} 
	return true;
}

/**
 * Function to delete the
 * Strength attributes
 * @param tablePkId
 * @param attrDesc
 */
function deleteStrength(tablePkId, attrDesc, profileId){
	alertify.set({ buttonReverse: true });
	alertify.confirm("Are you sure you want to delete the strength?", function (e) {
		if (e) {	    	
	    	var newUserProf = Spine.Model.sub();
	    	newUserProf.configure("/admin/contentconfig/deleteStrength", "tablePkId", "attrDesc","profileId");
	    	newUserProf.extend( Spine.Model.Ajax );
	    	//Populate the model with data to transfer
	    	var modelPopulator = new newUserProf({  
	    		tablePkId: tablePkId,
	    		attrDesc: attrDesc,
	    		profileId: profileId
	    	});
	    	modelPopulator.save(); //POST
	    	newUserProf.bind("ajaxError", function(record, xhr, settings, error) {
	    		alertify.alert("Unable to connect to the server.");
	    	});
	    	newUserProf.bind("ajaxSuccess", function(record, xhr, settings, error) {
	    		var jsonStr = JSON.stringify(xhr);
	    		var obj = jQuery.parseJSON(jsonStr);
	    		var statusCode = obj.statusCode;
	    		if(statusCode == 200) {
	    			//alertify.set({ delay: 2700 });
					alertify.alert("Strength has been deleted successfully.");
					fetchMappingList('S', profileId, "STR");
					//populateStrengthTable(obj.existingMappingList, "DELETE");
	    		} else if(statusCode == 500) {
	    			// Error Message
	    		}
	    	});
	    
		}
	});	
}

/**
 * Function to delete the
 * Value attributes
 * @param tablePkId
 * @param attrDesc
 */
function deleteValue(tablePkId, attrDesc, profileId){
	alertify.set({ buttonReverse: true });
	alertify.confirm("Are you sure you want to delete the value?", function (e) {
		if (e) {	    	
	    	var newUserProf = Spine.Model.sub();
	    	newUserProf.configure("/admin/contentconfig/deleteValue", "tablePkId", "attrDesc", "profileId");
	    	newUserProf.extend( Spine.Model.Ajax );
	    	//Populate the model with data to transfer
	    	var modelPopulator = new newUserProf({  
	    		tablePkId: tablePkId,
	    		attrDesc: attrDesc,
	    		profileId: profileId
	    	});
	    	modelPopulator.save(); //POST
	    	newUserProf.bind("ajaxError", function(record, xhr, settings, error) {
	    		alertify.alert("Unable to connect to the server.");
	    	});
	    	newUserProf.bind("ajaxSuccess", function(record, xhr, settings, error) {
	    		var jsonStr = JSON.stringify(xhr);
	    		var obj = jQuery.parseJSON(jsonStr);
	    		var statusCode = obj.statusCode;
	    		if(statusCode == 200) {
	    			//alertify.set({ delay: 2700 });
					alertify.alert("Value has been deleted successfully.");
					fetchMappingList('V', profileId, "VAL");
					//populateValueTable(obj.existingMappingList, "DELETE");
	    		} else if(statusCode == 500) {
	    			// Error Message
	    		}
	    	});
	    
		}
	});	
}

/**
 * Function to delete the
 * Passion attributes
 * @param tablePkId
 * @param attrDesc
 */
function deletePassion(tablePkId, attrDesc, profileId){
	alertify.set({ buttonReverse: true });
	alertify.confirm("Are you sure you want to delete the passion?", function (e) {
		if (e) {	    	
	    	var newUserProf = Spine.Model.sub();
	    	newUserProf.configure("/admin/contentconfig/deletePassion", "tablePkId", "attrDesc", "profileId");
	    	newUserProf.extend( Spine.Model.Ajax );
	    	//Populate the model with data to transfer
	    	var modelPopulator = new newUserProf({  
	    		tablePkId: tablePkId,
	    		attrDesc: attrDesc,
	    		profileId: profileId
	    	});
	    	modelPopulator.save(); //POST
	    	newUserProf.bind("ajaxError", function(record, xhr, settings, error) {
	    		alertify.alert("Unable to connect to the server.");
	    	});
	    	newUserProf.bind("ajaxSuccess", function(record, xhr, settings, error) {
	    		var jsonStr = JSON.stringify(xhr);
	    		var obj = jQuery.parseJSON(jsonStr);
	    		var statusCode = obj.statusCode;
	    		if(statusCode == 200) {
	    			//alertify.set({ delay: 2700 });
					alertify.alert("Passion has been deleted successfully.");
					fetchMappingList('P', profileId, "PAS");
					//populatePassionTable(obj.existingMappingList, "DELETE");
	    		} else if(statusCode == 500) {
	    			// Error Message
	    		}
	    	});
	    
		}
	});	
}


/**
 * Function call when user click on reset button
 * to reset the page
 * @param null
 */
function resetValue() {
	disableAllfield();
	$("#strengthAddDiv").removeClass("right_allgn_edit");
	$("#valueAddDiv").removeClass("right_allgn_edit");
	$("#passionAddDiv").removeClass("right_allgn_edit");
	document.getElementById('inputUserProfileInpt').disabled = false;
	document.getElementById('inputUserProfileInpt').value = "";	
	document.getElementsByName("mappingValFld")[0].disabled = false;
	document.getElementsByName("mappingValFld")[1].disabled = false;
	document.getElementsByName("mappingValFld")[2].disabled = false;
	//document.getElementsByName("mappingValFld")[0].checked=true;
	
	var globalProfileChgDiv = document.getElementById('globalProfileChgDiv');
	globalProfileChgDiv.setAttribute("style", "display:none");	
	
	var updateSection = document.getElementById("createMappingValUpdate");
	updateSection.className = updateSection.className - " otherclass";
	document.getElementById("updateId").innerHTML = "Mapping Values";
	var selectedVal = $('input[name=mappingValFld]:checked').val();
	if(selectedVal == "S"){
		document.getElementById('addHeaderDiv').innerHTML = "Add Strength";	
		document.getElementById("addStrengthBtnId").setAttribute("onclick", "saveStrength()");
	} else if(selectedVal == "V"){
		document.getElementById('addHeaderDiv').innerHTML = "Add Value";
		document.getElementById("addValueBtnId").setAttribute("onclick", "saveValue()");
	} else if(selectedVal == "P"){
		document.getElementById('addHeaderDiv').innerHTML = "Add Passion";	
		document.getElementById("addPassionBtnId").setAttribute("onclick", "savePassion()");
	}	
	document.getElementById("addStrengthBtnId").value = "Add";
	document.getElementById("addValueBtnId").value = "Add";		
	document.getElementById("addPassionBtnId").value = "Add";
	
	var chgOrderDiv = document.getElementById("chgOrderDiv");	
	chgOrderDiv.setAttribute("style", "display:block");
	
	var selectedRadioButton = $('input[name="mappingValFld"]:checked').val();
	if(selectedRadioButton == "P"){
		fetchMappingList(selectedRadioButton, 0, "PAS");
	} else if (selectedRadioButton == "V") {
		fetchMappingList(selectedRadioButton, 0, "VAL");	
	} else {
		fetchMappingList(selectedRadioButton, 0, "STR");	
	}
	
}

/**
 * Function add to allow only alphabets and numbers.
 * @param evt
 */
function alphaNumericInput(event) {	
	 var c = event.which || event.keyCode;	
	  if((c > 32 && c < 44) || ( c > 45 && c < 48 && c!= 46) || (c > 57 && c < 65) || (c > 90 && c < 97) ||
	     (c > 122 && c !== 127))
	    return false;
	  
}


function equalHeight(group) {
	   tallest = 0;
	   mmh = 15;
	   group.each(function() {
	      thisHeight = $(this).height();
	     // alert(thisHeight);
	      if(thisHeight > tallest) {
	         tallest = thisHeight;
	      }
	   });
	   group.height(tallest);
	}


function equalHeightAddDelete(group, heightVal) {
	//alert(heightVal);
	   tallest = 0;
	   mmh = 15;
	   group.each(function() {
	      thisHeight = $(this).height();
	     // alert(thisHeight);
	      tallest = thisHeight + heightVal;
	   });
	   group.height(tallest);
	}
	/*$(document).ready(function() {
	   equalHeight($(".col"));
	});*/

/**
 * Function call against the profile changed
 * @param obj
 */
function changeUserProfile(obj) {
	var relatedPage = "";
	var selectedRadioButton = $('input[name="mappingValFld"]:checked').val();
	if(selectedRadioButton == "P"){
		relatedPage = "PAS";
	} else if(selectedRadioButton == "V"){
		relatedPage = "VAL";
	} else {
		relatedPage = "STR";
	}
	if(null == obj.value || obj.value == "") {
		document.getElementById('attrDescInptId').setAttribute("disabled", "true");
		document.getElementById('attrFullDescInptId').setAttribute("disabled", "true");
		document.getElementById("chgOrderBtnId").disabled = true;
		fetchMappingList(selectedRadioButton, 0, relatedPage);
	} else {		
		document.getElementById('attrDescInptId').removeAttribute("disabled", "false");
		document.getElementById('attrFullDescInptId').removeAttribute("disabled", "false");
		document.getElementById("chgOrderBtnId").disabled = false ;
		fetchMappingList(selectedRadioButton, obj.value, relatedPage);
	}
}

/**
 * Function add to allow only numbers
 * as a input except 0
 * @param event
 */
function numberOnly(event) {	
	 var c = event.which || event.keyCode;
	 if ( (c >= 48  && c <= 57) || c == 43) 
	        return true;
	    return false;	  
}

/**
 * Function add to disable all the fields while reset
 * @param null
 */
function disableAllfield() {
	document.getElementById("attrDescInptId").value = "";
	document.getElementById("attrFullDescInptId").value = "";
	//document.getElementById("attrOrderInptId").value = "";
	
	document.getElementById('attrDescInptId').setAttribute("disabled", "true");
	document.getElementById('attrFullDescInptId').setAttribute("disabled", "true");
	//document.getElementById('attrOrderInptId').setAttribute("disabled", "true");
	document.getElementById("chgOrderBtnId").disabled = true;
}

/**
 * Function add to disable all the fields while reset
 * @param null
 */
function enableAllfield() {
	document.getElementById('attrDescInptId').removeAttribute("disabled", "true");
	document.getElementById('attrFullDescInptId').removeAttribute("disabled", "true");
	//document.getElementById('attrOrderInptId').removeAttribute("disabled", "true");
}

/**
 * Function add to restrict 0 as a input
 * and calculate the order value
 * @param event
 */
function changeAttributeOrder(obj) {	
	if(obj.value != "" && obj.value == 0){
		document.getElementById('qtnOrderInptId').value = "";	
		document.getElementById('inputNoOFSubQtn').value = "";	
		create("");
		alertify.alert("Question order cannot be 0.");
		return false;
	} else {
		var tbl = document.getElementById("actionPlanTableId");
		var numRows = tbl.rows.length;
		var array = new Array();
		for (var i = 1; i < numRows; i++) {
		    var cells = tbl.rows[i].getElementsByTagName('td');
		    array.push(cells[2].innerHTML);
		    if(array.indexOf(obj.value) > -1){
		    	document.getElementById('qtnOrderInptId').value = "";	
				document.getElementById('inputNoOFSubQtn').value = "";	
				create("");
		    	alertify.alert("Question order already exists.");
				return false;
		    }
		}
	}
	document.getElementById('inputNoOFSubQtn').value = "";	
	create("");
}


/**
 * Function add to disable all the fields while reset
 * @param null
 */
function setFieldOnUpdate(dist) {
	/** Attribute Description set null **/
	document.getElementById('attrDescInptId').value = "";	
	document.getElementById('attrFullDescInptId').value = "";	
	//document.getElementById('attrOrderInptId').value = "";
	
	document.getElementById('attrDescInptId').setAttribute("disabled", "true");
	document.getElementById('attrFullDescInptId').setAttribute("disabled", "true");
	//document.getElementById('attrOrderInptId').setAttribute("disabled", "true");
	
	/** Mapping Value radio button become enable **/
	document.getElementsByName("mappingValFld")[0].disabled = false;
	document.getElementsByName("mappingValFld")[1].disabled = false;
	document.getElementsByName("mappingValFld")[2].disabled = false;				
	/** User Profile dropdown become enable and set null **/
	document.getElementById('inputUserProfileInpt').disabled = false;	
	document.getElementById('inputUserProfileInpt').value = "";	
	document.getElementById("updateId").innerHTML = "Mapping Values";
	/*** Header description changed **/
	/*** Button description changed **/
	if( dist == "STR") {			
		document.getElementById('addHeaderDiv').innerHTML = "Add Strength";			
		document.getElementById("addStrengthBtnId").value = "Add";
		document.getElementById("addStrengthBtnId").setAttribute("onclick", "saveStrength()");
	} else if (dist == "VAL") {
		document.getElementById('addHeaderDiv').innerHTML = "Add Value";
		document.getElementById("addValueBtnId").value = "Add";
		document.getElementById("addValueBtnId").setAttribute("onclick", "saveValue()");
	} else if ( dist = "PAS") {
		document.getElementById('addHeaderDiv').innerHTML = "Add Passion";
		document.getElementById("addPassionBtnId").value = "Add";
		document.getElementById("addPassionBtnId").setAttribute("onclick", "savePassion()");
	}	
	/*** remove the class **/
	var updateSection = document.getElementById("createMappingValUpdate");
	updateSection.className = updateSection.className - " otherclass";
}

function changeOrder() {
	var attrCode = "";
	var selectedRadioButton = $('input[name="mappingValFld"]:checked').val();
	if(selectedRadioButton == "S"){
		attrCode = "STR";
		document.getElementById('headerDescMainId').innerHTML = "Change Strength Order";
	} else if (selectedRadioButton == "V") {
		attrCode = "VAL";
		document.getElementById('headerDescMainId').innerHTML = "Change Value Order";
	} else {
		attrCode = "PAS";
		document.getElementById('headerDescMainId').innerHTML = "Change Passion Order";
	}
	var profileId = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value;
	var userProfileText = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].text;
	
	populateMappingList(profileId, attrCode);	
}

/**
 * Function to fetch all the mapping attribute values
 * from database while landing to the page
 * @param attrCode
 */
function populateMappingList(profileId, attrCode) {	
	var userProf = Spine.Model.sub();
	userProf.configure("/admin/contentconfig/populateMappingList","profileId" , "attrCode");
	userProf.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userProf({  
		profileId: profileId,
		attrCode: attrCode
	});
	modelPopulator.save(); //POST
	userProf.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	userProf.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if(statusCode == 200) {						
			populateAttributeTable(obj.existingMappingList);	
		} else if(statusCode == 201) {
			document.getElementById("existingUserRoleListId").innerHTML = "<div align='center'>"+obj.statusDesc+"</div>";
		} else if(statusCode == 500) {
			//Show error message
		}
	});
}

function populateAttributeTable(listData) {		
	var tableDiv = document.getElementById("mainAreaListId");
	if (null != listData) {
		var headingString = "<table width='90%' id='table-1' border='1' align='center' class='tablesorter'><thead align='center' class='tab_header'><tr><th><b>Attribute Name</b></th></tr></thead><tbody>";
		var splitRec = listData.split("$$$");
		var counter = 1;
		for(var outerIndex = 0; outerIndex < splitRec.length-1; outerIndex++){			
			var objSplit = splitRec[outerIndex].split("~");		
			//if(objSplit[2] == 'BS') {		
				headingString = headingString + "<tr align='center' id="+counter+" class='custom_table_row' onmousedown='highlight()'><td align='center'>"+objSplit[0]+"</td>";	
				/*headingString = headingString + "<td align='center'>BB</td>";*/	
				headingString = headingString + "</tr>";
				counter = counter + 1;
			//} 
							
		}		
		headingString = headingString + "</tbody></table>";
		tableDiv.innerHTML = headingString;
		
	} 
	 $("#table-1").tableDnD();
	 $("table").tablesorter(); 
}
function highlight() {
	$("tr").mousedown(function(){
		$(this).addClass("selected").siblings().removeClass("selected");
	});
}

/**
 * Function to save attribute 
 * by attribute order
 * @param null
 */
function saveAttributeOrder() {
	var attrCode = "";
	var selectedRadioButton = $('input[name="mappingValFld"]:checked').val();
	if(selectedRadioButton == "S"){
		attrCode = "STR";
		document.getElementById('headerDescMainId').innerHTML = "Change Strength Order";
	} else if (selectedRadioButton == "V") {
		attrCode = "VAL";
		document.getElementById('headerDescMainId').innerHTML = "Change Value Order";
	} else {
		attrCode = "PAS";
		document.getElementById('headerDescMainId').innerHTML = "Change Passion Order";
	}
	var userProfile = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value;
	
	var builder = "";
	$('#table-1 tbody td').each(function() {
	    var cellText = $(this).html();  
	    builder = builder + cellText +"~~";	
	});
	var newUserProf = Spine.Model.sub();
	newUserProf.configure("/admin/contentconfig/saveAttributeOrder", "builder", "userProfile", "attrCode");
	newUserProf.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new newUserProf({  
		builder: builder,
		userProfile: userProfile,
		attrCode: attrCode
	});
	modelPopulator.save(); //POST
	newUserProf.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	newUserProf.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if(statusCode == 200) {
			alertify.alert("Attribute Order has been successfully changed.");
			populateAttributeTable(obj.mainQtnList);
			fetchMappingList(selectedRadioButton, userProfile , attrCode);
			//fetchMainQuestionList(selectedRadioButton, userProfile, relatedPage);		
		}  else if(statusCode == 500) {
			// Error Message
		}
	});
}