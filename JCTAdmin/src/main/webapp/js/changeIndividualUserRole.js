/**
 * Function calls when the page is loaded to populate 
 * 1. The drop-down for User
 * 2. The Total Subscribed User and total Registered User
 */
$(document).ready(function() {
	fetchUserDropdown();
	fetchExistingFacilitator();
});
/**
 * Method populate existing user to populate drop down list
 */
function fetchUserDropdown() {
	var userGrp = Spine.Model.sub();
	userGrp.configure("/admin/manageuserFacilitator/populateUserDropDown", "facilitatorEmail", "customerId");
	userGrp.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userGrp({  
		facilitatorEmail: sessionStorage.getItem("jctEmail"),
		customerId : sessionStorage.getItem("customerId")
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
			populateDropDown(obj.userNameMap);
		}  else {
			//Show error message
			alertify.alert("Some thing went wrong.");
		}
	});
}


/**
 * Function to populate the drop down for user
 */
function populateDropDown(userNameMap) {
	document.getElementById("existingUserInputId").innerHTML = "";	
	var userNameSelect = document.getElementById("existingUserInputId");
	var defaultOpn = document.createElement("option");
	defaultOpn.text = "Select User";
	defaultOpn.value = "";
	defaultOpn.className = "form-control-general";
	userNameSelect.add(defaultOpn, null);	
	for (var key in userNameMap) {				
		var option = document.createElement("option");
		option.text = userNameMap[key];
	    option.value = key+"!"+userNameMap[key];
	    option.className = "form-control-general";
	    try {
	    	userNameSelect.add(option, null); //Standard 
	    }catch(error) {
	    	//regionSelect.add(option); // IE only
	    }
	}
}


/**
 * Function to populate the existing user fields
 * while user click on existing user radio button
 * @param selection
 */
function toogleSelectionDiv(selection) {
	if (selection == "E") {
		document.getElementById("newUserDiv").style.display = "none";
		document.getElementById("newUserAddDiv").style.display = "none";	
		document.getElementById("existingUserDiv").style.display = "block";	
		document.getElementById("existingUserAddDiv").style.display = "block";	
		document.getElementById("existingUserInputId").value = "";
	} else {
		document.getElementById("existingUserDiv").style.display = "none";	
		document.getElementById("existingUserAddDiv").style.display = "none";	
		document.getElementById("newUserDiv").style.display = "block";		
		document.getElementById("newUserAddDiv").style.display = "block";	
	}
}

/**
 * Function to populate the existing activate and deactivate facilitator
 */
function fetchExistingFacilitator() {
	var userGrp = Spine.Model.sub();
	userGrp.configure("/admin/manageuserFacilitator/populateExistingFacilitator", "customerId");
	userGrp.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userGrp({  
		customerId : sessionStorage.getItem("customerId")
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
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
			populateTable(obj.existingFacilitators);
		}  else {
			//Show error message
			alertify.alert("Some thing went wrong.");
		}
	});
}

/**
 * Function to show the existing facilitator in table view
 * @param existingUsers
 */
function populateTable(existingUsers) {
	if(existingUsers.length > 0) {
		document.getElementById("existingChangeRoleList").style.display="block";
		var userList = document.getElementById("existingUsersTableId");
		var tableStr = "<table width='94%' border='1' bordercolor='#78C0D3' id='myTable' align='center'><thead class='tab_header'><tr><th>SL. No.</th><th>Facilitator Email</th><th>User Group Name</th><th width='16%'>Status</th></tr></thead><tbody>";
		var counter = 1;
		for (var index=0; index<existingUsers.length; index++) {
			var trColor = "";
			if(index % 2 == 0) {
				trColor = "#D2EAF0";
			} else {
				trColor = "#F1F1F1";
			}
						
			tableStr = tableStr + "<tr class='user_list_row_width' bgcolor='"+trColor+"'><td align='center'>"+counter+".</td><td align='center'>"+existingUsers[index].email+"</td><td align='center'>"+existingUsers[index].userGroup+"</td>";
			tableStr = tableStr + "<td class='user_list_row_width' >" +
					"<table width='100%' border='0'>" +
					"<tr>";
			if(existingUsers[index].softDelete == 0){
				tableStr = tableStr + "<td align='center'>Active</td></td></tr></table></td></tr>";
			} else {
				tableStr = tableStr + "<td align='center'>Inactive</td></td></tr></table></td></tr>";
			}
						
			counter = counter + 1;
		}
		tableStr = tableStr + "</tbody></table>";
		userList.innerHTML = tableStr;
		new SortableTable(document.getElementById('myTable'), 1);
	} else {
		document.getElementById("existingChangeRoleList").style.display="none";
		document.getElementById("existingUsersTableId").innerHTML = "<div align='center'><br /><br /><br /><img src='../img/no-record.png'><br /><div class='textStyleNoExist'>No Existing Facilitators</div></div>";
	}
}

/**
 * Function to update any the existing user to facilitator
 * @param null
 */
function updateUserToFacilitator() {
	var userVal = document.getElementById("existingUserInputId").options[document.getElementById("existingUserInputId").selectedIndex].value;
	if(validateUser(userVal)) {
	alertify.set({ buttonReverse: true });
	alertify.confirm("Are you sure you would like to replace your account with this new facilitator account?", function (e) {
		if (e) {		
			$(".loader_bg").fadeIn();
			$(".loader").fadeIn();
			var addFcltr = Spine.Model.sub();
			addFcltr.configure("/admin/manageuserFacilitator/updateUserToFacilitator", "userVal", "facilitatorEmail" ,"customerId");
			addFcltr.extend( Spine.Model.Ajax );
			//Populate the model with data to transfer
			var modelPopulator = new addFcltr({  
				userVal: userVal,
				facilitatorEmail: sessionStorage.getItem("jctEmail"),
				customerId : sessionStorage.getItem("customerId")
			});
			modelPopulator.save(); //POST
			addFcltr.bind("ajaxError", function(record, xhr, settings, error) {
				alertify.alert("Unable to connect to the server.");
			});
			addFcltr.bind("ajaxSuccess", function(record, xhr, settings, error) {
				var jsonStr = JSON.stringify(xhr);
				var obj = jQuery.parseJSON(jsonStr);
				var statusCode = obj.statusCode;
				if(statusCode == 200) {
					alertify.alert(obj.statusDesc);
					fetchUserDropdown();
					fetchExistingFacilitator();
					$(".loader_bg").fadeOut();
				    $(".loader").hide();			    
				    $('#alertify-ok').click(function() {	
				    	sessionStorage.clear();
				    	window.location = "../signup-page.jsp";
				    });
				} else if(statusCode == 999) {
					alertify.alert(obj.statusDesc);
					$(".loader_bg").fadeOut();
				    $(".loader").hide();			    
				}  else {
					alertify.alert("Some thing went wrong.");
				}
			});
			}		
		});	
	}
}


/**
 * Function to update any the existing user to facilitator
 * @param null
 */
function addNewFacilitator() {
	var newFacilitatorVal = document.getElementById("newUserInputId").value;
	if(validateEmailId(newFacilitatorVal)) {
	alertify.set({ buttonReverse: true });
	alertify.confirm("Are you sure you would like to replace your account with this new facilitator account?", function (e) {
		if (e) {	
			$(".loader_bg").fadeIn();
			$(".loader").fadeIn();
			var addFcltr = Spine.Model.sub();
			addFcltr.configure("/admin/manageuserFacilitator/addNewFacilitator", "newFacilitatorVal", "facilitatorEmail", "customerId");
			addFcltr.extend( Spine.Model.Ajax );
			//Populate the model with data to transfer
			var modelPopulator = new addFcltr({  
				newFacilitatorVal: newFacilitatorVal,
				facilitatorEmail: sessionStorage.getItem("jctEmail"),
				customerId : sessionStorage.getItem("customerId")
			});
			modelPopulator.save(); //POST
			addFcltr.bind("ajaxError", function(record, xhr, settings, error) {
				alertify.alert("Unable to connect to the server.");
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			});
			addFcltr.bind("ajaxSuccess", function(record, xhr, settings, error) {
				var jsonStr = JSON.stringify(xhr);
				var obj = jQuery.parseJSON(jsonStr);
				alertify.alert(obj.message);
				fetchUserDropdown();
				fetchExistingFacilitator();
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			    if(obj.statusCode != 999){
			    	$('#alertify-ok').click(function() {	
				    	sessionStorage.clear();
				    	window.location = "../signup-page.jsp";
				    });
			    }    
			});
		}	
	});
	}
}

/**
 * Function validates the user name.
 * @param userVal
 */
function validateUser(userVal) {
	if(userVal == "") {
		alertify.alert("Please select new facilitator.");
		return false;
	}
	return true;
}

/**
 * Function validates the email id.
 * @param emailId
 */
function validateEmailId(emailId) {	
	var filter = /^([a-zA-Z0-9_\.\-])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})+$/;
	if(emailId == "") {
		alertify.alert("Please provide new facilitator email.");
		return false;
	} else if (!filter.test(emailId)) {
		alertify.alert("Please provide a valid email address.");
    	return false;
    }
    return true;
}