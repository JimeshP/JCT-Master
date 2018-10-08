var emailIdCSVString = "";

/**
 * Function calls when the page is loaded to populate 
 * user group for facilitator
 */
$(document).ready(function() {
	fetchUserGroupForFacilitator();
});

/**
 * Function to populate the drop-down for User Group 
 */
function fetchUserGroupForFacilitator() {
	var userGrp = Spine.Model.sub();
	userGrp.configure("/admin/manageuserFacilitator/populateUserGroupDroDwnFacilitator", "customerId");
	userGrp.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userGrp({  
		customerId: sessionStorage.getItem("customerId")
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
			populateDropDown(obj.userGroupMap);
		}  else {
			//Show error message
			alertify.alert("Some thing went wrong.");
		}
	});
}

/**
 * Function to populate the dropdown for user group
 */
function populateDropDown(userGroupMap) {
	var userGroupSelect = document.getElementById("userGroupInptId");
	for (var key in userGroupMap) {				
		var option = document.createElement("option");
		option.text = userGroupMap[key];
	    option.value = key+"!"+userGroupMap[key];
	    option.className = "form-control-general";
	    try {
	    	userGroupSelect.add(option, null); //Standard 
	    }catch(error) {
	    	//regionSelect.add(option); // IE only
	    }
	}
}
/**
 * Function decides which div to open and plot fetched data
 * against the selected radio button
 * @param id
 */
function decidePlotting(id) {
	var displayDiv = document.getElementById("mainPopulationDiv");
	displayDiv.innerHTML = "<div align='center'>"+
	            		   "<img src='../img/dataLoading.GIF' />"+
	            		   "<p> Loading Existing Users</p>"+
	            		   "</div>";
	var userGrpVal = document.getElementById("userGroupInptId").options[document.getElementById("userGroupInptId").selectedIndex].value;
	var userGrp = Spine.Model.sub();
	var modelPopulator = null;
	if (document.getElementById('activeId').checked || document.getElementById('inctiveId').checked) {
		var status = 0;
		if(document.getElementById('activeId').checked) {
			status = 1;
		}
		userGrp.configure("/admin/manageuserFacilitator/fetchActiveInactiveUserListFacilitator", "emailId", "statusToBe","customerId", "userGroup");
		userGrp.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		modelPopulator = new userGrp({  
			emailId: sessionStorage.getItem("jctEmail"),
			statusToBe: status,
			customerId : sessionStorage.getItem("customerId"),
			userGroup: userGrpVal
		});
	} else {
		userGrp.configure("/admin/manageuserFacilitator/fetchUserListForResetPassword", "emailId", "customerId", "userGroup");
		userGrp.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		modelPopulator = new userGrp({  
			emailId: sessionStorage.getItem("jctEmail"),
			customerId : sessionStorage.getItem("customerId"),
			userGroup: userGrpVal
		});
	}
	
	modelPopulator.save(); //POST
	userGrp.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	});
	userGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if(statusCode == 200) {
			populateTable(obj.existingUsers, id);
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		}  else {
			//Show error message
			alertify.alert("Some thing went wrong.");
		}
	});
}
/**
 * Function populates and displays the fetched data in table.
 * @param existingUsers
 * @param id
 */
function populateTable(existingUsers, id) {
	document.getElementById("existing_data_div_id").style.display = "block";
	document.getElementById("mainPopulationDiv").style.display = "block";
	if (existingUsers.length > 0) {
		var buttonText = "";
		if(id == "activeId") {
			buttonText = "Activate";
		} else if(id == "inctiveId") {
			buttonText = "Deactivate";
		} else if(id == "passwordId") {
			buttonText = "Reset";
		}
		var userList = document.getElementById("mainPopulationDiv");
		var tableStr = "<table width='94%' border='1' bordercolor='#78C0D3' id='myTable' align='center'><thead class='tab_header'><tr><th></th><th>User Name</th></tr></thead><tbody>";
		var counter = 1;
		for (var index=0; index<existingUsers.length; index++) {
			var trColor = "";
			if(index % 2 == 0) {
				trColor = "#D2EAF0";
			} else {
				trColor = "#F1F1F1";
			}
			tableStr = tableStr + "<tr class='user_list_row_width' bgcolor='"+trColor+"'><td align='center'> <input type='checkbox' name='email' value='"+existingUsers[index].email+"' checked> </td><td align='center'>"+existingUsers[index].email+"</td>";
			counter = counter + 1;
		}
		tableStr = tableStr + "</tbody></table>";
		userList.innerHTML = tableStr;
		new SortableTable(document.getElementById('myTable'), 1);
		
		document.getElementById("buttonDivId").style.display = "block";
		document.getElementById("buttonDivId").innerHTML = "<table align='center'><tr><td style='padding:10px 0 0 0;'><input class='btn_admin btn_admin-primary btn_admin-sm search_btn' type='button' value='"+buttonText+"' onclick='performGrpAction(\""+id+"\")'></td></tr></table>";
	} else {
		document.getElementById("buttonDivId").style.display = "none";
		if(id == "activeId") {
			document.getElementById("mainPopulationDiv").innerHTML = "<div align='center'> All the users have been activated.</div>";
		} else if(id == "inctiveId") {
			document.getElementById("mainPopulationDiv").innerHTML = "<div align='center'> All the users have been deactivated.</div>";
		} else {
			document.getElementById("mainPopulationDiv").innerHTML = "<div align='center'> No user to reset the password.</div>";
		}
		
	}
}
/**
 * Function prepares the email id string separated by ~ and depending 
 * upon the type further actions will be performed
 * @param type
 */
function performGrpAction(type) {
	var checkedEmailString = "";
	var checkBoxes=document.getElementsByName("email");
	for (var i=0; i<checkBoxes.length; i++) {
		 if (checkBoxes[i].checked) {
			 if (checkBoxes.length-1 == i) {
				 checkedEmailString = checkedEmailString + checkBoxes[i].value;
			 } else {
				 checkedEmailString = checkedEmailString + checkBoxes[i].value + "~";
			 }
		 }
	}	
	if (checkedEmailString == "") {
		alertify.alert("Please select at least one user.");
	} else {
		if(type == "activeId") {
			changeStatusInBatch(checkedEmailString, 0);
		} else if(type == "inctiveId") {
			changeStatusInBatch(checkedEmailString, 1);
		} else if(type == "passwordId") {
			resetPasswordInBatch(checkedEmailString);
		}
	}
}
/**
 * Function changes the status from active to inactive for the email ids in email id string
 * @param checkedEmailString
 * @param status
 */
function changeStatusInBatch(checkedEmailString, status) {
	$(".loader_bg").fadeIn();
	$(".loader").fadeIn();
	var userGrpVal = document.getElementById("userGroupInptId").options[document.getElementById("userGroupInptId").selectedIndex].value;
	var userGrp = Spine.Model.sub();
	userGrp.configure("/admin/manageuserFacilitator/changeUserStatusInBatchFacilitator", "emailIdString", "statusToBe", "facilitatorId", "customerId", "userGroup");
	userGrp.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userGrp({  
		emailIdString: checkedEmailString,
		statusToBe: status,
		facilitatorId: sessionStorage.getItem("jctEmail"),
		customerId : sessionStorage.getItem("customerId"),
		userGroup: userGrpVal
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
			var activeStatus = "";
			if (status == 0) {
				activeStatus = "activeId";
			} else {
				activeStatus = "inctiveId";
			}
			populateTable(obj.existingUsers, activeStatus);
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		}  else {
			//Show error message
			alertify.alert("Some thing went wrong.");
		}
	});
}
/**
 * Function resets the password for the email ids in email id string
 * @param checkedEmailString
 */
function resetPasswordInBatch(checkedEmailString) {
	$(".loader_bg").fadeIn();
	$(".loader").fadeIn();
	var bulkPwdRst = Spine.Model.sub();
	bulkPwdRst.configure("/admin/manageuserFacilitator/resetPasswordInBatchFacilitator", "emailIdString", "updatedBy", "customerId");
	bulkPwdRst.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new bulkPwdRst({  
		emailIdString: checkedEmailString,
		updatedBy: sessionStorage.getItem("jctEmail"),
		customerId : sessionStorage.getItem("customerId")
	});
	modelPopulator.save(); //POST
	bulkPwdRst.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	bulkPwdRst.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if (statusCode == 201) {
			alertify.alert('Password reset was not successful.');
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		}  else if (statusCode == 222) {
			alertify.alert(obj.statusDesc);
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		} else {
			alertify.alert('Password reset was successful.');
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		}
	});
}

/**
 * Function opens the div which contains the radio buttons 
 * after selecting the user group 
 */
function openRadioDiv() {
	var userGrpVal = document.getElementById("userGroupInptId").options[document.getElementById("userGroupInptId").selectedIndex].value;
	if (userGrpVal != "") {
		document.getElementById('radioDrpDwnList').style.display = 'block';
	} else {
		document.getElementById('radioDrpDwnList').style.display = 'none';
		document.getElementById("buttonDivId").style.display = "none";
		document.getElementById('mainPopulationDiv').style.display = 'none';
		document.getElementById('existing_data_div_id').style.display = 'none';
					
		//remove the checked items...
		 if(document.getElementById('activeId').checked) {
			document.getElementById('activeId').checked = false;
		} else if(document.getElementById('inctiveId').checked) {
			document.getElementById('inctiveId').checked = false;
		} else if(document.getElementById('passwordId').checked) {
			document.getElementById('passwordId').checked = false;
		}
	}
}