var emailIdCSVString = "";
/**
 * Function calls when the page is loaded 
 * to populate the existing user
 */
$(document).ready(function() {
	fetchUserGroup();
	fetchExistingUsers();
});


/**
 * Function to populate the drop-down for User Group 
 */
function fetchUserGroup() {
	var userGrp = Spine.Model.sub();
	userGrp.configure("/admin/manageuser/populateUserGroup", "none");
	userGrp.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userGrp({  
		none: "" 
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
 * Function to populate the existing users
 */
function fetchExistingUsers() {
	var userGrp = Spine.Model.sub();
	userGrp.configure("/admin/manageuser/populateExistingUsers", "none");
	userGrp.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userGrp({  
		none: "" 
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
			populateTable(obj.existingUsers);
		}  else {
			//Show error message
			alertify.alert("Some thing went wrong.");
		}
	});
}


/**
 * Function to show the existing users in table view
 */
function populateTable(existingUsers) {
	if(existingUsers.length > 0) {
		var userList = document.getElementById("existingUsersTableId");
		var tableStr = "<table width='94%' border='1' bordercolor='#78C0D3' id='myTable' align='center'><thead class='tab_header'><tr><th>SL. No.</th><th>User Name</th><th>User Group Name</th><th width='22%'>Action</th></tr></thead><tbody>";
		var counter = 1;
		for (var index=0; index<existingUsers.length; index++) {
			var trColor = "";
			if(index % 2 == 0) {
				trColor = "#D2EAF0";
			} else {
				trColor = "#F1F1F1";
			}
			
			tableStr = tableStr + "<tr class='user_list_row_width' bgcolor='"+trColor+"'><td align='center'>"+counter+".</td><td align='center'>"+existingUsers[index].email+"</td><td align='center'>"+existingUsers[index].userGroup+"</td>";
			tableStr = tableStr + "<td class='table_col_txt_style' >" +
					"<table width='100%' border='0'>" +
					//"<tr><td align='center' style='width: 33%'><a href='#' onClick='deleteUser(\""+existingUsers[index].email+"\")'><font color='red'>Delete</font></a></td>";
					"<tr>";
			if(parseInt(existingUsers[index].softDelete) == 0) {
				tableStr = tableStr + "<td align='center'><a href='#' onclick='updateStatus(\""+existingUsers[index].email+"\", \""+existingUsers[index].softDelete+"\")'> Deactivate</a></td>";
			} else {
				tableStr = tableStr + "<td align='center'><a class='inactivate_style' href='#' onclick='updateStatus(\""+existingUsers[index].email+"\", \""+existingUsers[index].softDelete+"\")'> Activate</a></td>";
			}
			if (parseInt(existingUsers[index].activeInactiveStatus) == 1) {
				tableStr = tableStr + "<td align='center'><a href='#' onclick='resetPassword(\""+existingUsers[index].email+"\")'>Reset Password</a></td></td></tr></table></td></tr>";
			} else {
				tableStr = tableStr + "<td align='center'><span class='non_clickable'>Reset Password</span></td></td></tr></table></td></tr>";
			}
			
			
			counter = counter + 1;
		}
		tableStr = tableStr + "</tbody></table>";
		userList.innerHTML = tableStr;
		new SortableTable(document.getElementById('myTable'), 1);
	} else {
		document.getElementById("existingUsersTableId").innerHTML = "<div align='center'>No Existing Users.</div>";
	}
	
}


/**
 * Function to delete the users
 */
function deleteUser(emailId) {
	alertify.set({ buttonReverse: true });
	alertify.confirm("Are you sure you want to delete "+emailId+"?", function (e) {
		if (e) {
			var userDelete = Spine.Model.sub();
			userDelete.configure("/admin/manageuser/deleteExistingUsers", "emailId");
			userDelete.extend( Spine.Model.Ajax );
			//Populate the model with data to transfer
			var modelPopulator = new userDelete({  
				emailId: emailId
			});
			modelPopulator.save(); //POST
			userDelete.bind("ajaxError", function(record, xhr, settings, error) {
				alertify.alert("Unable to connect to the server.");
			});
			userDelete.bind("ajaxSuccess", function(record, xhr, settings, error) {
				var jsonStr = JSON.stringify(xhr);
				var obj = jQuery.parseJSON(jsonStr);
				var statusCode = obj.statusCode;
				if (statusCode == 401) {
					alertify.alert("Unable to delete "+emailId);
				} else {
					//alertify.set({ delay: 2700 });
					alertify.alert("User has been deleted successfully.");
				}
				populateTable(obj.existingUsers);
			});
		}
	});
}


/**
 * Function to change the status of a particular user
 * status can be activated or deactivated
 */
function updateStatus(emailId, softDelete) {
	var status = "";
	var changedStatus = 0;
	if(softDelete == 0) {
		status = "Deactivate";
		changedStatus = 1;
	} else {
		status = "Activate";
		changedStatus = 0;
	}

	alertify.set({ buttonReverse: true });
	alertify.confirm("Are you sure you want to "+status+" "+emailId+"?", function (e) {
		if (e) {
			var userDelete = Spine.Model.sub();
			userDelete.configure("/admin/manageuser/updateStatus", "emailId", "updatedStt");
			userDelete.extend( Spine.Model.Ajax );
			//Populate the model with data to transfer
			var modelPopulator = new userDelete({  
				emailId: emailId,
				updatedStt: changedStatus
				
			});
			modelPopulator.save(); //POST
			userDelete.bind("ajaxError", function(record, xhr, settings, error) {
				alertify.alert("Unable to connect to the server.");
			});
			userDelete.bind("ajaxSuccess", function(record, xhr, settings, error) {
				var jsonStr = JSON.stringify(xhr);
				var obj = jQuery.parseJSON(jsonStr);
				var statusCode = obj.statusCode;
				if (statusCode == 401) {
					alertify.alert("Unable to "+status+" "+emailId);
				} else {
					//alertify.set({ delay: 2700 });
					alertify.alert("User has been "+status+"d successfully.");
				}
				populateTable(obj.existingUsers);
			});
		}
	});

}


/**
 * Function to reset the password of a prticular user
 */
function resetPassword(emailId) {
	alertify.set({ buttonReverse: true });
	alertify.confirm("Are you sure you want to reset password of "+emailId+"?", function (e) {
		if (e) {
			var userDelete = Spine.Model.sub();
			userDelete.configure("/admin/manageuser/resetPassword", "emailId");
			userDelete.extend( Spine.Model.Ajax );
			//Populate the model with data to transfer
			var modelPopulator = new userDelete({  
				emailId: emailId
			});
			modelPopulator.save(); //POST
			userDelete.bind("ajaxError", function(record, xhr, settings, error) {
				alertify.alert("Unable to connect to the server.");
			});
			userDelete.bind("ajaxSuccess", function(record, xhr, settings, error) {
				var jsonStr = JSON.stringify(xhr);
				var obj = jQuery.parseJSON(jsonStr);
				var statusCode = obj.statusCode;
				
				if (statusCode == 200) {
					///]alertify.set({ delay: 2700 });
					alertify.alert("Password has been reset successful.");
				} else if (statusCode == 201) {
					//alertify.set({ delay: 2700 });
					alertify.alert("Password reset failed.");
				} else if (statusCode == 211) {
					//alertify.set({ delay: 2700 });
					alertify.alert("Sorry, did not get the mail id containing new password.");
				} else if (statusCode == 222) {
					alertify.alert(obj.statusDesc);
				}
				fetchExistingUsers();
				//populateTable(obj.existingUsers);
			});
		}
	});
}


/**
 * Function to populate the dropdown for user group
 */
function populateDropDown(userGroupMap) {
	var userGroupSelect = document.getElementById("inputUserGroupInpt");
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
 * Function to populate the manual entry fields
 * while user click on mnaual entry radio button
 * and to papulate csv file upload field 
 * once user click on csv file upload radio button
 */
function toogleSelectionDiv(selection) {
	if (selection == "M") {
		document.getElementById("fileUploadDiv").style.display = "none";
		document.getElementById("noOFuserDiv").style.display = "block";
		document.getElementById("manualEntry").innerHTML = "block";
		emailIdCSVString = "";
	} else {
		document.getElementById("noOFuserDiv").style.display = "none";
		document.getElementById("manualEntry").style.display = "none";
		document.getElementById("fileUploadDiv").style.display = "block";
		
	}
}


/**
 * Function to populate the manual entry fields
 * while user click on mnaual entry
 */
function populateManualEntryFields() {
	var manualEntryNumber = document.getElementById("inputNoOFUser").options[document.getElementById("inputNoOFUser").selectedIndex].value;
	var plotPlement = document.getElementById("manualEntry");
	document.getElementById("manualEntry").innerHTML = "";
	document.getElementById("manualEntry").style.display = "block";
	var string = "";
	var idCountIndex = 1;
	if (manualEntryNumber == "") {
		alertity.alert("Please select numbers.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	} else {
		//alert('coming in else');
		for (var index = 0; index < manualEntryNumber; index++) {
			var textFieldId = "manualEmailId_"+idCountIndex;
			string = string + "<div class='single_form_item'><div class='col-sm-4 '><label for='inputQtnNo' class='col-sm-12 control-label text_label'>User Name "+idCountIndex+":</label></div>" +
					"<div class='col-sm-6'><input type='text' class='form-control-general-form-redefined' maxlength='50' value='' placeholder='Enter Email Address' id='"+textFieldId+"'></div>" +
					"<div class='col-md-2'>&nbsp;</div><div class='clearfix'></div></div>";							
			idCountIndex = idCountIndex + 1;
		}
		plotPlement.innerHTML = string;
	}
}


$("#filename").change(function(e) {
	emailIdCSVString = "";
	var ext = $("input#filename").val().split(".").pop().toLowerCase();
	var emailIds = "";	
	if($.inArray(ext, ["csv"]) == -1) {
		alertify.alert('Upload only CSV file.');
		$("input#filename").value = null;
		document.getElementById('filename').value="";
		return false;
	}
    /*if (e.target.files != undefined) {
		var reader = new FileReader();
		reader.onload = function(e) {
		var csvval=e.target.result.split("\n");
		//alert(csvval.length);
		for(var index = 1; index<csvval.length-1; index++) {
			var csvvalue = csvval[index].split(",");
			if(csvvalue != "\r") {
				emailIds = emailIds + csvvalue[0].replace("\r", "~");
			}
		}
	emailIdCSVString = JSON.stringify(emailIds);
	//alert(emailIdCSVString);
	};
	
	reader.readAsText(e.target.files.item(0));
	document.getElementById("filename").value = "";*/
	
	/**
	 * New impl
	 */
	if (e.target.files != undefined) {
		var reader = new FileReader();
		reader.onload = function(e) {
		var csvval=e.target.result.split("\r");
		//alert(csvval.length);
		for(var index = 1; index<csvval.length; index++) {
			var csvvalue = csvval[index].split(",");
			emailIds = emailIds + csvvalue[0]+"~";
			/*if(csvvalue != "\r") {
				emailIds = emailIds + csvvalue[0].replace("\r", "~");
			}*/
		}
	//emailIdCSVString = JSON.stringify(emailIds);
	emailIdCSVString = JSON.stringify(emailIds.replace(/\n/g, "").replace(/\r/g, ""));
	//alert(emailIdCSVString);
	};
	
	reader.readAsText(e.target.files.item(0));
	document.getElementById("filename").value = "";
}

return false;
});

/**
 * Function to save the users
 */
function saveUser() {
	$(".loader_bg").fadeIn();
	$(".loader").fadeIn();
	var userGrpVal = document.getElementById("inputUserGroupInpt").options[document.getElementById("inputUserGroupInpt").selectedIndex].value;
	if(userGrpVal == "") {
		alertify.alert("Please select User Group.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
		return false;
	} else {
		if (document.getElementById("manual").checked) {
			var numberOfUsers = document.getElementById("inputNoOFUser").options[document.getElementById("inputNoOFUser").selectedIndex].value;
			if(numberOfUsers == "") {
				alertify.alert("Please select Number of users.");
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
				return false;
			} else {
				saveManualEntry(prepareEmailIdString());
			}
		} else if (document.getElementById("csv").checked){
			saveCSVEntry();
		} else {
			alertify.alert('Please select an option first.');
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		}
	}
}

/**
 * Function to save the users
 * by manual entry
 */
function saveManualEntry(emailIdString) {
	if (emailIdString != "") {
		var userGrpVal = document.getElementById("inputUserGroupInpt").options[document.getElementById("inputUserGroupInpt").selectedIndex].value;
		var model = Spine.Model.sub();
		model.configure("/admin/manageuser/saveManualUser", "userGroup", "emailIdString", "createdBy");
		model.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new model({  
			userGroup: userGrpVal,
			emailIdString: emailIdString,
			createdBy: sessionStorage.getItem("jctEmail")
		});
		modelPopulator.save(); //POST
		model.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		});
		model.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			var statusCode = obj.statusCode;
			alertify.alert(obj.statusDesc);
			fetchExistingUsers();
		});
	} else {
		alertify.alert('Please provide a valid email address.');
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	}
}
/**
 * Function prepares the email id string seperated by ~
 * @returns {String}
 */
function prepareEmailIdString() {
	emailIdCSVString = "";
	var numberOfUsers = document.getElementById("inputNoOFUser").options[document.getElementById("inputNoOFUser").selectedIndex].value;
	var id = 1;
	for (var index=0; index<numberOfUsers; index++) {
		var emailString = document.getElementById("manualEmailId_"+id).value;
		if (validateEmailId(emailString)) {
			if (emailString.length > 5) {
				if(id != numberOfUsers) {
					emailIdCSVString = emailIdCSVString + emailString + "~";
				} else {
					emailIdCSVString = emailIdCSVString + emailString;
				}
				id = id + 1;
			}
		} else {
			//alert('Please provide a valid email address');
		    return false;
		}
		
	}
	return emailIdCSVString;
}
/**
 * Function validates the email id.
 * @param emailId
 */
function validateEmailId(emailId) {	
	var filter = /^([a-zA-Z0-9_\.\-])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})+$/;
    if (!filter.test(emailId)) {
    	//alert('Please provide a valid email address');
    	return false;
    }
    return true;
}
/**
 * Function to save the users
 * by csv file upload
 */
function saveCSVEntry() {
	//FOR CSV
	if (emailIdCSVString == "") {
		alertify.alert("Please select the CSV file.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	} else if (emailIdCSVString.trim().length < 5){
		alertify.alert("CSV file is empty.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	} else {
		var userGrpVal = document.getElementById("inputUserGroupInpt").options[document.getElementById("inputUserGroupInpt").selectedIndex].value;
		var model = Spine.Model.sub();
		model.configure("/admin/manageuser/saveUser", "userGroup", "emailIdString", "createdBy");
		model.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new model({  
			userGroup: userGrpVal,
			emailIdString: emailIdCSVString.substring(1, emailIdCSVString.length-1),
			createdBy: sessionStorage.getItem("jctEmail")
		});
		modelPopulator.save(); //POST
		model.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		});
		model.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			var statusCode = obj.statusCode;
			alertify.alert(obj.statusDesc);
			fetchExistingUsers();
		});
	}
	
	/*if (emailIdCSVString != "") {
		var userGrpVal = document.getElementById("inputUserGroupInpt").options[document.getElementById("inputUserGroupInpt").selectedIndex].value;
		var model = Spine.Model.sub();
		model.configure("/admin/manageuser/saveUser", "userGroup", "emailIdString", "createdBy");
		model.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new model({  
			userGroup: userGrpVal,
			emailIdString: emailIdCSVString.substring(1, emailIdCSVString.length-1),
			createdBy: sessionStorage.getItem("jctEmail")
		});
		modelPopulator.save(); //POST
		model.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		});
		model.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			var statusCode = obj.statusCode;
			alertify.alert(obj.statusDesc);
			fetchExistingUsers();
		});
	} else {
		alertify.alert("Please select the CSV file.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	}*/
}


/**
 * Function to download the template of csv file
 */
function downloadCSVFmt() {
	var link = document.getElementById('downloadCSVFmt');
	link.setAttribute("href", "../../admin/manageuser/downloadCSVFile/Sample.csv");
}