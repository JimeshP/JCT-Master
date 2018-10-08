/**
 * Function calls when the page is loaded 
 * to populate the user group list
 */
$(document).ready(function() {
	if (sessionStorage.getItem("userRole") == 2) {
		disableAllfield();
		fetchUserGroupList(0);
		/***************** To Disable the paste functionality ***************/
		/* $("#userGroupInptId").bind("paste",function(e) {
        	e.preventDefault();
    	});*/
		} else {
			document.getElementById("ugForm").style.display = "none";
			document.getElementById("progessDataDiv").innerHTML = sessionStorage.getItem("authMsg");
		}
});

/**
 * Function to populate the user group list
 * @param null
 */
function fetchUserGroupList( profileId ) {
	var userGrp = Spine.Model.sub();
	userGrp.configure("/admin/manageuser/populateExistingUserGroup", "profileId", "customerId", "roleId");
	userGrp.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userGrp({  
		profileId: profileId,
		customerId: sessionStorage.getItem("customerId"),
		roleId: 2
	});
	modelPopulator.save(); //POST
	userGrp.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	userGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if(statusCode == 200) {
			if(profileId == 0){
				populateDropDown(obj.userProfileMap);	
			}
			populateUserGroupTable(obj.existingUserGroupList);
		} else if(statusCode == 201) {
			document.getElementById("existingUserGroupListId").innerHTML = "<div align='center'><img src='../img/no-record.png'><br />"+obj.statusDesc+"</div>";
		} else if(statusCode == 500) {
			//Show error message
		}
	});
}

/**
 * Function call to save the user group in database
 * @param null
 */
function saveUserGroup() {
	var userGroup = document.getElementById("userGroupInptId").value;
	var userProfile = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value;
	var userProfileText = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].text;
	if(validateUserGroup(userGroup)) {
		var newUserProf = Spine.Model.sub();
		newUserProf.configure("/admin/manageuser/saveUserGroup", "userGroupVal", "userProfileVal", 
						"userProfileId", "createdBy", "customerId" , "roleId");
		newUserProf.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new newUserProf({  
			userGroupVal: userGroup,
			userProfileVal: userProfileText,
			userProfileId: userProfile,
			createdBy: sessionStorage.getItem("jctEmail"),
			customerId: sessionStorage.getItem("customerId"),
			roleId: 2
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
				alertify.alert("User Group has been successfully created.");
				document.getElementById("userGroupInptId").value = "";	
				//document.getElementById("inputUserProfileInpt").value = "";				
				//populateUserGroupTable(obj.existingUserGroupList);
				fetchUserGroupList(userProfile);
			} else if(statusCode == 600) {
				alertify.alert("User Group already exists.");
			} else if(statusCode == 500) {
				// Error Message
			}
		});
	}
}

/**
 * Function to validate the user group data
 * @param userGroup
 */
function validateUserGroup(userGroup){
	var re=/^\s*$/;
	if(document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value == ""){
		alertify.alert("Please select User Profile.");
		return false;
	}
	if (re.test(userGroup)){
		alertify.alert("Please enter User Group Name.");
		return false;
	}	
	return true;
}

/**
 * Function to show the user group list in table view
 * @param userGroupListData
 */
function populateUserGroupTable(userGroupListData) {
	var tableDiv = document.getElementById("existingUserGroupListId");
	if (null != userGroupListData) {
		var headingString = "<table width='94%' id='userGroupTableId' border=1 align='center' bordercolor='#78C0D3' class='tablesorter'><thead class='tab_header'><tr><th width='11%'><b>&nbsp; SL. No.</b></th><th><b>&nbsp;&nbsp; User Group Name</b></th><th><b>&nbsp;&nbsp; User Profile Name</b></th><th align='center' width='16%'><b>&nbsp;&nbsp; Action</b></th></tr></thead><tbody>";
		var splitRec = userGroupListData.split("$$$");
		var counter = 1;
		for(var outerIndex = 0; outerIndex < splitRec.length-1; outerIndex++){
			var trColor = "";
			if(outerIndex % 2 == 0) {
				trColor = "#D2EAF0";
			} else {
				trColor = "#F1F1F1";
			}
			headingString = headingString + "<tr align='center' class='row_width' bgcolor='"+trColor+"'><td>&nbsp; "+counter+".</td>";
			var objSplit = splitRec[outerIndex].split("~");
			//for(var innerIndex = 0; innerIndex < objSplit.length-1; innerIndex++){
			headingString = headingString + "<td>&nbsp;&nbsp; "+objSplit[0]+"</td>";
			headingString = headingString + "<td>&nbsp;&nbsp; "+objSplit[1]+"</td>";
			var activeStatus = parseInt(objSplit[2]);
			if(activeStatus == 1) { //Active User
				headingString = headingString + "<td class='table_col_txt_style'><a href='#' onClick='editUserGroup(\""+objSplit[3]+"\", \""+objSplit[0]+"\", \""+objSplit[1]+"\", \""+objSplit[4]+"\")'>Edit</a> &nbsp;&nbsp;&nbsp;&nbsp; <a class='inactivate_style' href='#' onClick='toogleGroupStatus(\""+objSplit[3]+"\", \""+activeStatus+"\", \""+objSplit[0]+"\", \""+objSplit[4]+"\")'>Deactivate</a></td>";
			} else { //Inactive User
				headingString = headingString + "<td class='table_col_txt_style'><a href='#' onClick='editUserGroup(\""+objSplit[3]+"\", \""+objSplit[0]+"\", \""+objSplit[1]+"\",\""+objSplit[4]+"\")'>Edit</a> &nbsp;&nbsp;&nbsp;&nbsp; <a href='#' onClick='toogleGroupStatus(\""+objSplit[3]+"\", \""+activeStatus+"\", \""+objSplit[0]+"\",\""+objSplit[4]+"\")'>Activate</a></td>";
			}
			//}
			headingString = headingString + "</tr>";
			counter = counter + 1;
		}
		headingString = headingString + "</tbody></table>";
		tableDiv.innerHTML = headingString;
		//new SortableTable(document.getElementById('userGroupTableId'), 1);
		$("table").tablesorter();
	} else {
		tableDiv.innerHTML = "<div align='center'>No User Group To Display</div>";
	}
}

/**
 * Function to populate the drop-down for user profile 
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
 * Function to edit user group
 * @param tablePkId
 * @param userGroupName
 * @param dropDownVal
 */
function editUserGroup(tablePkId, userGroupName, dropDownVal, profileId){
	fetchUserGroupList(profileId);
	document.getElementById('userGroupInptId').value = userGroupName;
	document.getElementById('userGroupInptId').setAttribute("readonly", "true");
	document.getElementById("updateId").innerHTML = "Update User Group";
	document.getElementById("createUserGroupBtnId").value = "Update";
	document.getElementById("addUserGrpText").innerHTML = "Update User Group";	
	document.getElementById("createUserGroupBtnId").setAttribute("onclick", "updateUserGroup("+tablePkId+",'"+userGroupName+"')");
	var updateSection = document.getElementById("createUserGroupUpdate");
	updateSection.className = updateSection.className + " otherclass";
	//populateUserProfile
	var usrGrp = Spine.Model.sub();
	usrGrp.configure("/admin/manageuser/populateUserProfile", "none");
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
 * Function to update user group
 * @param tablePkId
 * @param prevVal
 */
function updateUserGroup(tablePkId, prevVal) {
	var userGroup = document.getElementById("userGroupInptId").value;
	var userProfile = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value;
	var userProfileText = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].text;
	if(validateUserGroup(userGroup)) {
		var newUserProf = Spine.Model.sub();
		newUserProf.configure("/admin/manageuser/updateUserGroup", "userGroupVal", "userProfileVal", "userProfileId", "tablePkId", "customerId", "roleId");
		newUserProf.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new newUserProf({  
			//userGroupVal: userGroup,
			userGroupVal: prevVal,
			userProfileVal: userProfileText,
			userProfileId: userProfile,
			tablePkId: tablePkId,
			customerId: sessionStorage.getItem("customerId"),
			roleId: 2
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
				alertify.alert("User Group has been edited successfully.");
				document.getElementById("userGroupInptId").value = "";
				//document.getElementById("inputUserProfileInpt").value = "";	
				document.getElementById("createUserGroupBtnId").value = "Create";
				document.getElementById("updateId").innerHTML = "Create User Group";
				document.getElementById("addUserGrpText").innerHTML = "Add User Group";
				document.getElementById("createUserGroupBtnId").setAttribute("onclick", "saveUserGroup()");				
				/*** remove the class **/
				var updateSection = document.getElementById("createUserGroupUpdate");
				updateSection.className = updateSection.className - " otherclass";
				
				//populateUserGroupTable(obj.existingUserGroupList);
				fetchUserGroupList(userProfile);
			} else if(statusCode == 600) {
				alertify.alert("User Group already exists.");
			} else if(statusCode == 500) {
				// Error Message
			}
		});
	}
}

/**
 * Function to change the user group status 
 * as activate and deactivate
 * @param tablePkId
 * @param status
 * @param userGroup
 */
function toogleGroupStatus(tablePkId, status, userGroup, profileId){
	var confirmMsg = "";
	var stt = "";
	if(status == 1) {
		status = 2;
		stt = "Deactivated.";
		confirmMsg = "Once you deactivate the group, all the users of the group will be deactived. \n\n Are you sure you want to continue?";
	} else if(status == 2) {
		status = 1;
		confirmMsg = "Once you activate the group, all the users of the group will be active. \n\n Are you sure you want to continue?";
		stt = "Activated.";
	}
	alertify.set({ buttonReverse: true });
	alertify.confirm(confirmMsg, function (e) {
	    if (e) {		    	
	    	var newUserProf = Spine.Model.sub();
	    	newUserProf.configure("/admin/manageuser/toogleGroupStatus", "tablePkId", "status", "userGroup", "profileId", "customerId", "roleId");
	    	newUserProf.extend( Spine.Model.Ajax );
	    	//Populate the model with data to transfer
	    	var modelPopulator = new newUserProf({  
	    		tablePkId: tablePkId,
	    		status: status,
	    		userGroup: userGroup,
	    		profileId: profileId,
	    		customerId: sessionStorage.getItem("customerId"),
				roleId: 2
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
					alertify.alert("User Group has been successfully "+stt+"");
	    			populateUserGroupTable(obj.existingUserGroupList);
	    		} else if(statusCode == 500) {
	    			// Error Message
	    		}
	    	});
	    } 
	});
}

/**
 * Function add to disable all the fields while reset
 * @param null
 */
function disableAllfield() {
	document.getElementById("userGroupInptId").value = "";	
	document.getElementById('userGroupInptId').setAttribute("disabled", "true");
}

/**
 * Function call against the profile changed
 * @param obj
 */
function changeUserProfile(obj) {
	if(null == obj.value || obj.value == "") {
		document.getElementById('userGroupInptId').setAttribute("disabled", "true");
		fetchUserGroupList(0);
	} else {	
		document.getElementById('userGroupInptId').removeAttribute("disabled", "true");
		fetchUserGroupList(obj.value);
	}
	
}