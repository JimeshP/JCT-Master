var emailIdCSVString = "";
var totalUsers = 0;
var usrType = sessionStorage.getItem("userType");
/**
 * Function calls when the page is loaded 
 * to populate the existing user
 */
$(document).ready(function() {
	sessionStorage.removeItem("previous_user_email"); 
	fetchUserGroup();
	$("#manual").prop("checked", true);
	var newOption;
	//	selectUserType
	if(usrType == 1){
		newOption = "<option value='"+usrType+"'>Individual</option>";
		$("#selectUserType").append(newOption);
		$("#selectUserType").change();
	} else {
		newOption = "<option value='"+usrType+"'>Facilitator</option>";
		$("#selectUserType").append(newOption);
		$("#selectUserType").change();		
		$("#inputUserGroupInpt").get(0).selectedIndex = 1;
		$("#inputUserGroupInpt").change();
	}
});
/**
 * Function to allow only numeric value
 */
$("#noOfUser2, #chequeno").keypress(function (e) {   
	 if (e.which != 8 && e.which != 0 && (e.which < 48 || e.which > 57)) {       
	       return false;
	   }
});
/**
 * in case of Amount field number and . should be allowed
 **/
$("#totalAmount, #totalAmountIndi, #chkAmountFaci, #cashAmountFaci").keypress(function (e) {	
	 if (e.which != 8 && e.which != 0 && ((e.which < 48 || e.which > 57) && e.which != 46)) {       
	       return false;
	 }
});
/**
 * function to reset all numeric fields if any non-numeric or special char entered 
 **/
$("#chkAmountFaci, #cashAmountFaci, #totalAmount, #chequeNoFaci, #noOfUser2, #chequeno, #totalAmountIndi").blur(function(){	
	var fieldId = $(this).attr('id');
	var val = $("#"+fieldId).val().trim();
	if( isNaN(val) || val == 0){
		$("#"+fieldId).val("");			
	}
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
	//sessionStorage.getItem("jctEmail")
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
			
			/*document.getElementById("inputUserGroupInpt").disabled = false;*/
			//document.getElementById("selectUserType").selectedIndex = "0";
			document.getElementById("inputUserGroupInpt").selectedIndex = "0";
			document.getElementById("paymentModeInd").selectedIndex = "0";
			document.getElementById("paymentMode1").selectedIndex = "0";
			document.getElementById("chkAmountFaci").value = "";
			document.getElementById("chequeNoFaci").value = "";			
			$(".inputRightSide").hide();
			$("#paymentModeIdIndividual, #expiryDurationId").show();			
			document.getElementById("datepickerCashIndiinput").value = "";			
			document.getElementById("manualEntry").innerHTML = "";
			document.getElementById("manualEntry").style.display = "none";
			
			if (document.getElementById("bankDetailsVal")) {
				document.getElementById("bankDetailsVal").value = "";
			}
			if (document.getElementById("datepicker")) {
				document.getElementById("datepicker").value = "";
			}
			if (document.getElementById("chequeno")) {
				document.getElementById("chequeno").value = "";
			}
			if (document.getElementById("totalAmount")) {
				document.getElementById("totalAmount").value = "";
			}
			if (document.getElementById("inputNoOFUser")) {
				document.getElementById("inputNoOFUser").selectedIndex = "0";
			}
			/*if (document.getElementById("noOfUser1")) {
				document.getElementById("noOfUser1").value = "";
			}*/
			if (document.getElementById("noOfUser2")) {
				document.getElementById("noOfUser2").value = "";
			}
			if (document.getElementById("totalAmount")) {
				document.getElementById("totalAmount").value = "";
			}
			if (document.getElementById("expiryDuration")) {
				document.getElementById("expiryDuration").selectedIndex = "0";
			}
			if (document.getElementById("inputNoOFUser")) {
				document.getElementById("inputNoOFUser").selectedIndex = "0";
			}
			if (document.getElementById("facEmail")) {
				document.getElementById("facEmail").value = "";
			}
			
			  if( $("#facCheck").is(":checked")) {
				  $('#facCheck').attr('checked', false);
			  }
			
		}  else {
			//Show error message
			alertify.alert("Some thing went wrong.");
		}
	});
}


/**
 * Function to populate the existing users
 */
function fetchExistingUsersByUserTypeAndGroup() {
	var usrType = document.getElementById("selectUserType").options[document.getElementById("selectUserType").selectedIndex].value;
	var userGrpVal = "";
	if(usrType == 1){			// individual	
		userGrpVal = document.getElementById("inputUserGroupInpt").options[document.getElementById("inputUserGroupInpt").selectedIndex].value;
	} else if (usrType == 2) {	//	facilitator
		userGrpVal = "1!Default User Group";		
	}
	if (userGrpVal.trim().length == 0) {
		userGrpVal = "UGNS"; //User Group Not Selected
	}
	if (usrType == "0") {
		usrType = "UTNS";    //User Type Not Selected
	}
	var userGrp = Spine.Model.sub();
	userGrp.configure("/admin/manageuser/populateExistingUsersByUserTypeAndGroup", "userGrpVal", "usrType");
	userGrp.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userGrp({  
		userGrpVal: userGrpVal,
		usrType: usrType
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
 * @param existingUsers
 */
function populateTable(existingUsers) {
	
	if(existingUsers.length > 0) {
		document.getElementById("existingDataListDiv").style.display="block";
		var userList = document.getElementById("existingUsersTableId");
		var tableStr = "<table width='94%' border='1' bordercolor='#78C0D3' id='myTable' align='center' class='tablesorter'><thead width='94%' class='tab_header'><tr><th width='6%'>#</th><th>Name (Last, First)</th><th>Email (User ID)</th><th>User Group Name</th><th>Expiration Date</th><th>Creator</th><th width='22%'>Action</th></tr></thead><tbody width='94%'>";
		var counter = 1;
		var fName= "";
		var lName = "";
		var creator = "";
		for (var index=0; index<existingUsers.length; index++) {
			var trColor = "";
			if(index % 2 == 0) {
				trColor = "#D2EAF0";
			} else {
				trColor = "#F1F1F1";
			}
			var expiryDate = "";
			var userExpryDate = "";
			
			if(null != existingUsers[index].lastName) {
				lName = existingUsers[index].lastName;
			} else {
				lName = "N/A";
			}
			if(null != existingUsers[index].firstName) {
				fName = existingUsers[index].firstName;
			} else {
				fName = "N/A";
			}
			var custId = existingUsers[index].customerId.substring(0, 2);
			var adminId = existingUsers[index].jctUserDetailAdminId;
			if (existingUsers[index].roleId == 3) {
				if (adminId == 2) {
					creator = "Self";
				} else {
					creator = "Admin";
				}
			} else {
				if(custId == 99) {
					creator = "Facilitator ("+existingUsers[index].createdBy+")";
				} else if (custId == 98 && adminId == 2){
					creator = "Self";
				} else {
					creator = "Admin";
				}
			}
			
				
			if (existingUsers[index].jctProfileId == 1) {
				expiryDate = new Date(existingUsers[index].jctAccountExpirationDate).toDateString();
				userExpryDate = dateformat(new Date (expiryDate));
			}
			expiryDate = ( expiryDate == ('' || null || 0) ) ? 'N/A' : userExpryDate ;		//	for facilitator -> N/A
			
			tableStr = tableStr + "<tr class='user_list_row_width' bgcolor='"+trColor+"'><td align='center'>"+counter+".</td><td align='center'>"+lName+", "+fName+"</td><td align='center'>"+existingUsers[index].email+"</td><td align='center'>"+existingUsers[index].userGroup+"</td><td align='center'>"+expiryDate+"</td><td align='center'>"+creator+"</td>";
			tableStr = tableStr + "<td class='table_col_txt_style' >" +
					"<table width='100%' border='0'>" +
					"<tr>";
			
			//If profile id is 3 then
			if (existingUsers[index].jctProfileId == 3) {
				tableStr = tableStr + "<td align='center' width='50%'><a href='#' onclick='resetPassword(\""+existingUsers[index].email+"\", \""+existingUsers[index].jctProfileId+"\")'>Reset Password</a></td>" +
								"<td align='center' width='50%'><a href='#' onclick='deleteUser(\""+existingUsers[index].email+"\", \"3\")'>Delete</a></td></td></tr></table></td></tr>";
				
			} else {
				if(parseInt(existingUsers[index].softDelete) == 0) {
					tableStr = tableStr + "<td align='center' width='50%'><a href='#' onclick='updateStatus(\""+existingUsers[index].email+"\", \""+existingUsers[index].softDelete+"\")'> Deactivate</a></td>";
				} else {
					tableStr = tableStr + "<td align='center' width='50%'><a class='inactivate_style' href='#' onclick='updateStatus(\""+existingUsers[index].email+"\", \""+existingUsers[index].softDelete+"\")'> Activate</a></td>";
				}
				if (parseInt(existingUsers[index].activeInactiveStatus) == 1) {
					tableStr = tableStr + "<td align='center' width='50%'><a href='#' onclick='resetPassword(\""+existingUsers[index].email+"\", \""+existingUsers[index].jctProfileId+"\")'>Reset Password</a></td></td><td align='center' width='50%'><a href='#' onclick='deleteUser(\""+existingUsers[index].email+"\", \"1\")'>Delete</a></td></tr></table></td></tr>";
				} else {
					tableStr = tableStr + "<td align='center' width='50%'><span class='non_clickable'>Reset Password</span></td></td><td align='center' width='50%'><a href='#' onclick='deleteUser(\""+existingUsers[index].email+"\", \"1\")'>Delete</a></td></tr></table></td></tr>";
				}
			}
			counter = counter + 1;
		}
		tableStr = tableStr + "</tbody></table>";
		userList.innerHTML = tableStr;
		//$("table").tablesorter(); 
		$(function(){
			  $("table").tablesorter({
			    headers: {
			      0: { sorter: false }      // disable first column			     
			    }
			  });
			});
	} else {
		document.getElementById("existingUsersTableId").innerHTML = "<div align='center'><br /><br /><br /><img src='../img/no-record.png'><br /><div class='textStyleNoExist'>No Existing Users.</div></div>";
		document.getElementById("existingDataListDiv").style.display="none";
	}
}


/**
 * Function to delete the users
 */
function deleteUser(emailId, userType) {
	alertify.set({ buttonReverse: true });
	var message = "Are you sure you want to delete user "+emailId+"?";
	if (userType == 3) {
		message = "Are you sure you want to delete facilitator "+emailId+"? <br /><br /><b>Note: </b>All users under this facilitator will also be deleted.";
	}
	alertify.confirm(message, function (e) {
		if (e) {
			var userDelete = Spine.Model.sub();
			userDelete.configure("/admin/manageuser/deleteExistingUsers", "emailId", "userType");
			userDelete.extend( Spine.Model.Ajax );
			//Populate the model with data to transfer
			var modelPopulator = new userDelete({  
				emailId: emailId,
				userType: userType
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
				//populateTable(obj.existingUsers);
				/*if(usrType == 3) {
					fetchExistingUsersByUserTypeAndGroup();
				} else {
					fetchExistingUsersByUserTypeAndGroup();
				}*/
				fetchExistingUsersByUserTypeAndGroup();
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
				fetchExistingUsersByUserTypeAndGroup();
			});
		}
	});
}


/**
 * Function to reset the password of a prticular user
 */
function resetPassword(emailId, roleId) {
	alertify.set({ buttonReverse: true });
	alertify.confirm("Are you sure you want to reset password of "+emailId+"?", function (e) {
		if (e) {
			// Blur the screen
			$(".loader_bg").fadeIn();
			$(".loader").fadeIn();
			console.log(roleId);
			var userDelete = Spine.Model.sub();
			userDelete.configure("/admin/manageuser/resetPassword", "emailId", "roleId");
			userDelete.extend( Spine.Model.Ajax );
			//Populate the model with data to transfer
			var modelPopulator = new userDelete({  
				emailId: emailId,
				roleId: roleId
			});
			modelPopulator.save(); //POST
			userDelete.bind("ajaxError", function(record, xhr, settings, error) {
				alertify.alert("Unable to connect to the server.");
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			});
			userDelete.bind("ajaxSuccess", function(record, xhr, settings, error) {
				var jsonStr = JSON.stringify(xhr);
				var obj = jQuery.parseJSON(jsonStr);
				var statusCode = obj.statusCode;
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
				if (statusCode == 200) {
					alertify.alert("Password has been reset successfully.");
				} else if (statusCode == 201) {
					alertify.alert("Password reset failed.");
				} else if (statusCode == 211) {
					alertify.alert("Sorry, did not get the mail id containing new password.");
				} else if (statusCode == 222) {
					alertify.alert(obj.statusDesc);
				}
				fetchExistingUsersByUserTypeAndGroup();
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
	if(usrType == 2){	//	facilitator will be always under default grp
		document.getElementById("inputUserGroupInpt").selectedIndex = "1";	
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
		if (document.getElementById("noOfUser2")) {
			document.getElementById("noOfUser2").value = "";
		}		
		document.getElementById("fileUploadDiv").style.display = "none";
		document.getElementById("manualEntry").innerHTML = "block";
		document.getElementById("inputNoOFUser").style.display = "block";		
		if (document.getElementById("noOfUser2")) {
			document.getElementById("noOfUser2").style.display = "none";
		}
		emailIdCSVString = "";
		document.getElementById("inputNoOFUser").value = "";
		document.getElementById("noOfUserLabelId").innerHTML  = "Number of Users:";
	} else {
		if (document.getElementById("noOfUser2")) {
			document.getElementById("noOfUser2").value = "";
		}
		document.getElementById("manualEntry").style.display = "none";
		document.getElementById("fileUploadDiv").style.display = "block";
		document.getElementById("inputNoOFUser").style.display = "none";		
		if (document.getElementById("noOfUser2")) {
			document.getElementById("noOfUser2").style.display = "block";
		}
		document.getElementById("noOfUserLabelId").innerHTML  = "";
	}
}

/**
 * Function to populate the manual entry fields
 * while user click on mnaual entry
 */
 function populateManualEntryFields() {
	$("#qorkingArea1").show();
	divCount = $('#manualEntry').children('.single_form_item:visible').length/2;
	var jsonObj = [];
	var counter = 0;
	for (var i = 1; i <= divCount; i += 1) {
	var unitJ = {};	
	var value = document.getElementById('manualEmailId_'+i).value;
	var fNamevalue = document.getElementById('firstNamelId_'+i).value;
	var lNamevalue = document.getElementById('lastNamelId_'+i).value;
	unitJ["ElementValue"+i] = value;
	unitJ["FirstNameValue"+i] = fNamevalue;
	unitJ["LastNameValue"+i] = lNamevalue;
	jsonObj[counter++] = unitJ;
	}
	sessionStorage.setItem("previous_user_email",JSON.stringify(jsonObj));		
	var manualEntryNumber = document.getElementById("inputNoOFUser").options[document.getElementById("inputNoOFUser").selectedIndex].value;
		
	var plotPlement = document.getElementById("manualEntry");
	document.getElementById("manualEntry").innerHTML = "";
	document.getElementById("manualEntry").style.display = "block";
	var string = "";
	var idCountIndex = 1;
	if (manualEntryNumber == "") {
		alertify.alert("Please select numbers.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	} else {
		for (var index = 0; index < manualEntryNumber; index++) {
			var textFieldId = "manualEmailId_"+idCountIndex;
			var firstNameFieldId = "firstNamelId_"+idCountIndex;
			var lastNameFieldId = "lastNamelId_"+idCountIndex;
			/*string = string + "<div class='single_form_item'><div class='col-sm-3 '><label for='inputQtnNo' class='col-sm-12 control-label text_label'>User Name "+idCountIndex+":</label></div>" +
					"<div class='col-sm-6'><input type='text' class='form-control-general-form-redefined manual_user_width ipad_text_width' maxlength='50' value='' placeholder='Enter Email Address' id='"+textFieldId+"'></div>" +
					"<div class='clearfix'></div></div>";	*/
			string = string + "<div class='single_form_item'><div class='col-sm-3'><label for='inputQtnNo' class='col-sm-12 control-label text_label'>User Name "+idCountIndex+":</label></div>" +
			"<div class='col-sm-6'><input type='text' class='form-control-general-form-redefined manual_user_width ipad_text_width' maxlength='50' value='' placeholder='Enter Email Address' id='"+textFieldId+"'></div>" +
			"<div class='clearfix'></div></div>" +
			"<div class='single_form_item'><div class='col-sm-3'><label for='inputQtnNo' class='col-sm-12 control-label text_label'>First Name:</label></div>" +
			"<div class='col-sm-3'><input type='text' class='form-control-general' maxlength='50' value='' placeholder='Enter First Name' id='"+firstNameFieldId+"' onkeypress='return validateDescription(event)'></div>" +
			"<div class='col-sm-3'><label for='inputQtnNo' class='col-sm-12 control-label text_label'>Last Name:</label></div>" +
			"<div class='col-sm-3'><input type='text' class='form-control-general' maxlength='50' value='' placeholder='Enter Last Name' id='"+lastNameFieldId+"' onkeypress='return validateDescription(event)'></div>" +
			"<div class='clearfix'></div></div></br>";	
			idCountIndex = idCountIndex + 1;
		}
		plotPlement.innerHTML = string;
	}
	
	if(sessionStorage.getItem("previous_user_email") != null){
		var jsonObj = JSON.parse(sessionStorage.getItem("previous_user_email"));
		var count = jsonObj.length;
			for (var j = 0; j<count; j++){	
			var value = jsonObj[j]["ElementValue"+(j+1)];
			var fNamevalue = jsonObj[j]["FirstNameValue"+(j+1)];
			var lNamevalue = jsonObj[j]["LastNameValue"+(j+1)];	
			var div = document.getElementById('manualEmailId_'+(j+1));
			    if (div) {
				document.getElementById('manualEmailId_'+(j+1)).value = value;
				document.getElementById('firstNamelId_'+(j+1)).value = fNamevalue;
				document.getElementById('lastNamelId_'+(j+1)).value = lNamevalue;
				}			
			}		
		}
}
 

$("#filename").change(function(e) {
	emailIdCSVString = "";
	var ext = $("input#filename").val().split(".").pop().toLowerCase();
	var emailIds = "";	
	if($.inArray(ext, ["csv"]) == -1 && ext.length > 0) {
		alertify.alert('Upload only CSV file.');
		$("input#filename").value = null;
		document.getElementById('filename').value="";
		return false;
	}
	
	if (e.target.files != undefined) {
		var reader = new FileReader();
		reader.onload = function(e) {
		var csvval=e.target.result.split("\r");
		totalUsers = 0;
		for(var index = 1; index<csvval.length; index++) {
			totalUsers = totalUsers + 1;
			var csvvalue = csvval[index].split(",");
			if (csvvalue[0].trim() != "")
				//emailIds = emailIds + csvvalue[0]+"~";
				if(csvvalue[0].toString().trim().length == 0 || csvvalue[1].toString().trim().length == 0 || csvvalue[2].toString().trim().length == 0){
					alertify.alert('Please provide all values in CSV file.');
					$("input#filename").value = null;
					document.getElementById('filename').value="";
					return false;
				} else {
					emailIds = emailIds + csvvalue[0]+ "#`#"+ csvvalue[1] + "#`#"+ csvvalue[2] + "~";		
				}
		}
	emailIdCSVString = JSON.stringify(emailIds.replace(/\n/g, "").replace(/\r/g, ""));
	};
	
	reader.readAsText(e.target.files.item(0));
	document.getElementById("filename").value = "";
}

return false;
});

/**
 * Function to save the users
 * by manual entry
 */
function saveManualEntry(emailIdString) {
	if (emailIdString != "") {
		var userGrpVal = document.getElementById("inputUserGroupInpt").options[document.getElementById("inputUserGroupInpt").selectedIndex].value;
		var model = Spine.Model.sub();
		model.configure("/admin/manageuser/saveManualUser", "userGroup", "emailIdString", "createdBy","paymentType");
		model.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new model({  
			userGroup: userGrpVal,
			emailIdString: emailIdString,
			createdBy: sessionStorage.getItem("jctEmail"),
			paymentType: paymentType
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
			alertify.alert(obj.statusDesc);
			resetFields();
			fetchExistingUsersByUserTypeAndGroup();
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
		var fNamevalue = document.getElementById('firstNamelId_'+id).value;
		var lNamevalue = document.getElementById('lastNamelId_'+id).value;
		if (validateEmailId(emailString)) {
			if (emailString.length > 5) {
				if(id != numberOfUsers) {
					//emailIdCSVString = emailIdCSVString + emailString + "~";
					emailIdCSVString = emailIdCSVString + emailString + "#`#" +fNamevalue + "#`#" +lNamevalue + "~";
				} else {
					//emailIdCSVString = emailIdCSVString + emailString;
					emailIdCSVString = emailIdCSVString + emailString+ "#`#" +fNamevalue + "#`#" +lNamevalue;
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
		model.configure("/admin/manageuser/saveUser", "userGroup", "emailIdString", "createdBy","paymentType");
		model.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new model({  
			userGroup: userGrpVal,
			emailIdString: emailIdCSVString.substring(1, emailIdCSVString.length-1),
			createdBy: sessionStorage.getItem("jctEmail"),
			paymentType : paymentType
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
			alertify.alert(obj.statusDesc);
			resetFields();
			fetchExistingUsersByUserTypeAndGroup();
		});
	}
}


/**
 * Function to download the template of csv file
 */
function downloadCSVFmt() {
	var link = document.getElementById('downloadCSVFmt');
	link.setAttribute("href", "https://s3-us-west-2.amazonaws.com/jobcrafting/csv/Sample.csv");
}
function fieldsEnabled() {
	var usrType = document.getElementById("selectUserType").options[document.getElementById("selectUserType").selectedIndex].value;
	$(".inputRightSide").hide();
	if(usrType == "0"){
		disableFieldsOnLoad();
		$("#paymentModeInd")[0].selectedIndex = 0;
		$("#manual").prop("checked", true);
		document.getElementById("fileUploadDiv").style.display = 'none';
		//document.getElementById("noOfUser1").style.display = 'none';
		document.getElementById("noOfUserLabelId").innerHTML  = "Number of Users:";
		document.getElementById("inputNoOFUser").style.display = 'block';
		resetFields();
		fieldsForIndividual();
		
	}
	else if (usrType == "2") { //FACILITATOR
			enableAllFields();
			resetFields();
			fieldsForFacilitator();
	} else {				//INDIVIDUAL
			enableAllFields();
			resetFields();
			fieldsForIndividual();
		}

	fetchExistingUsersByUserTypeAndGroup();
}

function tooglePaymentMode() {

	var paymentMode = '';	
	var usrType = document.getElementById("selectUserType").options[document.getElementById("selectUserType").selectedIndex].value;
	if(usrType == '0'){
		alertify.alert("Please select user type first.");
		$("#paymentModeInd")[0].selectedIndex = 0;
		$("#paymentMode1")[0].selectedIndex = 0;
		return false;
	} else if(usrType == "1") {	//	indi
		paymentMode = document.getElementById("paymentModeInd").options[document.getElementById("paymentModeInd").selectedIndex].value;		
	} else if(usrType == "2"){	//	faci
		paymentMode = document.getElementById("paymentMode1").options[document.getElementById("paymentMode1").selectedIndex].value;		
	}
	
	if (usrType == "2") {
		$("#paymentDt2").hide().val('');	//	used for individual
		if (paymentMode == "CHK") {
			document.getElementById("chkAmmoutIndividual").style.display = "none";
			document.getElementById("chkNo2").style.display = "none";
			$("#chkAmmoutIndividual").hide();
			$("#chkAmountFaciDiv").show();
			$("#cashAmountFaciDiv").hide();
			$("#paymentDtFaciCash").hide();
			document.getElementById("chkNo1").style.display = "block";
			document.getElementById("paymentDt1").style.display = "block";
			
			$("#cashAmountFaci").val('');			//faci cash
			$("#paymentDtFaciCashInpt").val('');
			
			$("#totalAmount").val('');			//indi chk
			$("#chequeno").val('');
			$("#datepicker").val('');
			
			$("#totalAmountIndi").val('');			//indi  cash
			$("#datepickerCashIndiinput").val('');
			
		} else if (paymentMode == "CASH") {
			
			$("#chkAmmoutIndividual").hide();
			$("#chkAmountFaciDiv").hide();
			$("#cashAmountFaciDiv").show();
			$("#paymentDtFaciCash").show();
			$("#paymentDt2").hide();
			$("#paymentDt2").val('');
			$("#chkNo1").hide('');
			$("#paymentDt1").hide().val('');
			
			$("#totalAmount").val('');		//indi chk
			$("#chequeno").val('');
			$("#datepicker").val('');
			
			$("#totalAmountIndi").val('');	//indi cash
			
			$("#chkAmountFaci").val('');	//faci chk
			$("#chequeNoFaci").val('');
			$("#datepickerCashIndiinput").val('');
			
		} else {	//	FREE
			$("#chkAmmoutIndividual").hide();
			$("#chkAmountFaciDiv").hide();
			$("#cashAmountFaciDiv").hide();
			$("#paymentDtFaciCash").hide();
			$("#paymentDt2").hide();
			$("#chkNo1").hide();
			$("#paymentDt1").hide();
			
			$("#totalAmount").val('');		//indi chk
			$("#chequeno").val('');
			$("#datepicker").val('');
			
			$("#totalAmountIndi").val('');	//indi cash
			
			$("#chkAmountFaci").val('');	//faci chk
			$("#chequeNoFaci").val('');
			$("#datepickerCashIndiinput").val('');
			
			$("#cashAmountFaci").val('');			//faci cash
			$("#paymentDtFaciCashInpt").val('');
		}
	} else {
		document.getElementById("paymentDt2").style.display = "none";	//	faci payment date
		if (paymentMode == "CHK") {
			document.getElementById("chkAmmoutIndividual").style.display = "block";
			document.getElementById("cashAmmoutIndividual").style.display = "none";
			document.getElementById("chkNo2").style.display = "block";
			document.getElementById("paymentDt2").style.display = "block";
			document.getElementById("paymentDt1").style.display = "none";
			$("#chkAmountFaciDiv").hide();	
			
			$("#totalAmountIndi").val('');				//indi cash
			$("#datepickerCashIndiinput").val('');
			
			$("#chkAmountFaci").val('');		//faci chk
			$("#chequeNoFaci").val('');
			$("#datepickerCashIndiinput").val('');
			
			$("#cashAmountFaci").val('');		//faci cash
			$("#paymentDtFaciCashInpt").val('');
			
		} else if (paymentMode == "CASH") {
			document.getElementById("cashAmmoutIndividual").style.display = "block";
			document.getElementById("chkAmmoutIndividual").style.display = "none";
			document.getElementById("chkNo2").style.display = "none";
			document.getElementById("paymentDt2").style.display = "none";
			document.getElementById("paymentDt1").style.display = "block";
			$("#chkAmountFaciDiv").hide();
			
			$("#totalAmount").val('');	//indi chk
			$("#chequeno").val('');
			$("#datepicker").val('');
			
			$("#chkAmountFaci").val('');	//faci chk
			$("#chequeNoFaci").val('');
			$("#datepickerCashIndiinput").val('');
			
			$("#cashAmountFaci").val('');		//faci cash
			$("#paymentDtFaciCashInpt").val('');
		}
		
		else {
			document.getElementById("chkAmmoutIndividual").style.display = "none";
			document.getElementById("cashAmmoutIndividual").style.display = "none";
			document.getElementById("chkNo2").style.display = "none";
			document.getElementById("paymentDt2").style.display = "none";
			document.getElementById("paymentDt1").style.display = "none";
			
			$("#totalAmount").val('');		//indi chk
			$("#chequeno").val('');
			$("#datepicker").val('');
			
			$("#totalAmountIndi").val('');		//indi cash
			$("#datepickerCashIndiinput").val('');
			
			$("#cashAmountFaci").val('');		//faci cash
			$("#paymentDtFaciCashInpt").val('');
			
			$("#chkAmountFaci").val('');	//faci chk
			$("#chequeNoFaci").val('');
		}
	}
}

function saveUser() {	
	var usrType = document.getElementById("selectUserType").options[document.getElementById("selectUserType").selectedIndex].value;
	if(usrType == 0){	//		when no user type selected
		alertify.alert("Please select an user type");		
	} else {		
		if (usrType == "2") { //Facilitator
		// check the payment mode
		var paymentMode = document.getElementById("paymentMode1").options[document.getElementById("paymentMode1").selectedIndex].value;
		if (paymentMode.trim().length > 0) {
			if (paymentMode == "CHK") {
				var allFieldValidation = validateAllMandatoryFields ("CHK");
				if (allFieldValidation == true) {
					$(".loader_bg").fadeIn();
					$(".loader").fadeIn();
					saveFacilitatorChkPayment(); // Check Payment
				}
			} else if (paymentMode == "CMP") {	//	free
				var allFieldValidation = validateAllMandatoryFields ("CMP");
				if (allFieldValidation == true) {
					$(".loader_bg").fadeIn();
					$(".loader").fadeIn();
					saveFacilitatorCmpPayment(); // Free
				}
			} else if(paymentMode == "CASH") {
				var allFieldValidation = validateAllMandatoryFields ("CASH");
				if (allFieldValidation == true) {
					$(".loader_bg").fadeIn();
					$(".loader").fadeIn();
					saveFacilitatorCashPayment(); // Cash
				}
			}
		} else {
				alertify.alert("Please select a payment mode");
			}
		} else { //general users
			var expiryDuration = document.getElementById("expiryDuration").options[document.getElementById("expiryDuration").selectedIndex].value;
			var userGrpVal = document.getElementById("inputUserGroupInpt").options[document.getElementById("inputUserGroupInpt").selectedIndex].value;
			if ((expiryDuration == "") || (expiryDuration.trim().length == 0)) {						//		expiration date
				alertify.alert("Please select date of expiration.");
				return false;			
			}
			if (userGrpVal.trim().length == 0) {
				alertify.alert("Please select an user group");
				return false;
			}
			var paymentMode = document.getElementById("paymentModeInd").options[document.getElementById("paymentModeInd").selectedIndex].value;
			if (paymentMode.trim().length > 0) {
				if (paymentMode == "CHK") {
					
					var allFieldValidation = validateAllMandatoryFieldsGeneralUser ("CHK");
					if (allFieldValidation == true) {
						$(".loader_bg").fadeIn();
						$(".loader").fadeIn();
						if (document.getElementById("manual").checked) {			
							validateCSVEntryOnAddUser(prepareEmailIdString(), "manual");
						} else {													
							validateCSVEntryOnAddUser(emailIdCSVString, "csv");
						}
					}
				} else if (paymentMode == "CMP") {
					var allFieldValidation = validateAllMandatoryFieldsGeneralUser ("CMP");
					if (allFieldValidation == true) {
						$(".loader_bg").fadeIn();
						$(".loader").fadeIn();
						if (document.getElementById("manual").checked) {			
							validateCSVEntryOnAddUser(prepareEmailIdString(), "manual");
						} else {													
							validateCSVEntryOnAddUser(emailIdCSVString, "csv");
						}
					}
				} else if (paymentMode == "CASH") {
					var allFieldValidation = validateAllMandatoryFieldsGeneralUser ("CASH");
					if (allFieldValidation == true) {
						$(".loader_bg").fadeIn();
						$(".loader").fadeIn();
						if (document.getElementById("manual").checked) {			
							validateCSVEntryOnAddUser(prepareEmailIdString(), "manual");
						} else {													 
							validateCSVEntryOnAddUser(emailIdCSVString, "csv");
						}
						
					}
				}
			} else {
				alertify.alert("Please select a payment mode");
			}
		}
	}
}

function validateAllMandatoryFieldsGeneralUser (paymentMode) {		
	if (paymentMode == "CHK") {
		var paymentDate = document.getElementById("datepicker").value.trim();
		var chequeNos = document.getElementById("chequeno").value.trim();	
		var totalAmount = document.getElementById("totalAmount").value.trim();	
	
		if ((totalAmount.length == 0) || (totalAmount == '') || isNaN(totalAmount)){
			alertify.alert("Please enter check amount.");
			return false;			
		}
		if (chequeNos.length == 0) {					//		check number
			alertify.alert("Please enter check number");
			return false;
		} else {			
			if( chequeNumberValidation(chequeNos) == false )
				return false;			
		}
		if (paymentDate.length == 0) {					//		payment date
			alertify.alert("Please select a payment date");
			return false;
		}
		
		if (document.getElementById("csv").checked) {		//		CSV upload			
		} else {
			var nosOfUser = document.getElementById("inputNoOFUser").options[document.getElementById("inputNoOFUser").selectedIndex].value;
			if (nosOfUser.trim().length == 0) {
				alertify.alert("Please select number of users");
				return false;
			}
		}
		return true;
	} else if (paymentMode == "CMP") {	//	Free
		var usrType = document.getElementById("selectUserType").options[document.getElementById("selectUserType").selectedIndex].value;
		var userGrpVal = document.getElementById("inputUserGroupInpt").options[document.getElementById("inputUserGroupInpt").selectedIndex].value;
		if (usrType.trim().length == 0) {
				alertify.alert("Please select an user type");
				return false;
			} else if (userGrpVal.trim().length == 0) {
				alertify.alert("Please select an user group");
				return false;
			}
		if (document.getElementById("csv").checked) {		//		CSV upload
			/*var nosOfUser = document.getElementById("noOfUser1").value.trim();	
			if ((nosOfUser.length == 0) || (nosOfUser == 0)) {
				alertify.alert("Please enter number of users");
				return false;
			}*/
		} else {
			var nosOfUser = document.getElementById("inputNoOFUser").options[document.getElementById("inputNoOFUser").selectedIndex].value;
			if (nosOfUser.trim().length == 0) {
				alertify.alert("Please select number of users");
				return false;
			}
		}
		return true;
	} else if (paymentMode == "CASH") {	//	Free
		var usrType = document.getElementById("selectUserType").options[document.getElementById("selectUserType").selectedIndex].value;
		var userGrpVal = document.getElementById("inputUserGroupInpt").options[document.getElementById("inputUserGroupInpt").selectedIndex].value;
		var totalAmount = document.getElementById("totalAmountIndi").value.trim();
		var paymentDate = document.getElementById("datepickerCashIndiinput").value.trim();
		if (usrType.trim().length == 0) {
				alertify.alert("Please select an user type");
				return false;
			} else if (userGrpVal.trim().length == 0) {
				alertify.alert("Please select an user group");
				return false;
		} else if ((totalAmount.length == 0 )||( totalAmount == 0)){
			alertify.alert("Please enter cash amount");
			return false;
		} else if ((paymentDate.length == 0)||( paymentDate == 0)){
			alertify.alert("Please enter payment date");
			return false;
		}		
		if (document.getElementById("csv").checked) {		//		CSV upload
			/*var nosOfUser = document.getElementById("noOfUser1").value.trim();	
			if ((nosOfUser.length == 0) || (nosOfUser == 0)) {
				alertify.alert("Please enter number of users");
				return false;
			}*/
		} else {
			var nosOfUser = document.getElementById("inputNoOFUser").options[document.getElementById("inputNoOFUser").selectedIndex].value;
			if (nosOfUser.trim().length == 0) {
				alertify.alert("Please select number of users");
				return false;
			}
		}
		return true;
	}
}		
function bankDetailsValidation(bankDetails){
	var pattern = /^[A-Za-z0-9 \.\,]+$/;						//  special char aren't allowed
	if(!pattern.test(bankDetails) || isNaN(bankDetails)){
	  alertify.alert('Amount should contain only numeric value.');
	  return false;
	}
}
function chequeNumberValidation(chequeNos){	
//	var pattern = /^[0-9]+$/;
	if( chequeNos.length < 6){
		alertify.alert('Length of Cheque Number should be atleast 6.');
		return false;
	}

}

function validateAllMandatoryFields (paymentMode) {		//	validation for facilitator
	var usrType = document.getElementById("selectUserType").options[document.getElementById("selectUserType").selectedIndex].value;
	var facFirstName = document.getElementById("facFirstNameId").value.trim();
	var facLastName = document.getElementById("facLastNameId").value.trim();
	if (paymentMode == "CHK") {	
		var chkAmount = document.getElementById("chkAmountFaci").value.trim();
		var chequeNos = document.getElementById("chequeNoFaci").value.trim();		
		var	paymentDate = document.getElementById("datepickerCashIndiinput").value.trim();
		var	facilitatorEmail = document.getElementById("facEmail").value.trim();			
		var	numberOfUsers = document.getElementById("noOfUser2").value;
		if( chkAmount.length == 0 ){
			alertify.alert("Please enter check amount.");
			return false;
		} else {
			if( bankDetailsValidation(chkAmount) == false){
				return false;
			}			
		}	
		if( chequeNos.length == 0 ){
			alertify.alert("Please enter check number.");
			return false;
		} else {
			if( chequeNumberValidation(chequeNos) == false){
				return false;
			}	else if (isNaN(chequeNos)){
				alertify.alert("Check number should contain numeric value only");
				return false;
			}		
		}	
		if (paymentDate.length == 0) {
			alertify.alert("Please select a payment date");
			return false;
		}
		if ((numberOfUsers.trim().length == 0) ||(numberOfUsers == 0)) {
			alertify.alert("Please enter number of users the admin has paid for");
			return false;
		}		
		if(usrType == 2 ) {
			if (!validateEmailId(facilitatorEmail)) {
				alertify.alert("Please enter a valid email id");
				return false;
			}
		}
		//return true;
	} else if (paymentMode == "CMP") {	//	Free
		var facilitatorEmail = document.getElementById("facEmail").value.trim();
		var numberOfUsers = document.getElementById("noOfUser2").value.trim();
		if (!validateEmailId(facilitatorEmail)) {
			alertify.alert("Please enter a valid email id");
			return false;
		} else if ((numberOfUsers.length == 0) || (numberOfUsers == 0)) {
			alertify.alert("Please enter number of users the admin has paid for");
			return false;
		}
		//return true;
	} else if (paymentMode == "CASH") {
		var totalAmount = document.getElementById("cashAmountFaci").value.trim();
		var paymentDate = document.getElementById("paymentDtFaciCashInpt").value.trim();
		var facilitatorEmail = document.getElementById("facEmail").value.trim();
		var numberOfUsers = document.getElementById("noOfUser2").value.trim();		
		if((totalAmount.length == 0)  || totalAmount == 0){
			alertify.alert("Please enter cash amount.");
			return false;
		} else if (paymentDate.length == 0) {
			alertify.alert("Please select a payment date");
			return false;
		} else if (!validateEmailId(facilitatorEmail)) {
			alertify.alert("Please enter a valid email id");
			return false;
		} else if ((numberOfUsers.length == 0) || (numberOfUsers == 0)) {
			alertify.alert("Please enter number of users the admin has paid for");
			return false;
		}
		//return true;
	} 
	if(facFirstName.length == 0) {
		alertify.alert("Please enter facilitator first name");
		return false;
	} else if (facLastName.length == 0) {
		alertify.alert("Please enter facilitator last name");
		return false;
	}
	return true;
}
/**
 * Function saves facilitator through check payment
 */
function saveFacilitatorChkPayment() {
	var userGrpVal = document.getElementById("inputUserGroupInpt").options[document.getElementById("inputUserGroupInpt").selectedIndex].value;
	//var bankDetails = document.getElementById("bankDetailsVal").value;
	var facilitatorEmail = document.getElementById("facEmail").value;
	var paymentDate = document.getElementById("datepickerCashIndiinput").value;
	var numberOfUsers = document.getElementById("noOfUser2").value;
	var chequeNos = document.getElementById("chequeNoFaci").value;
	var totalAmountToBePaid = document.getElementById("chkAmountFaci").value;
	var facFirstName = document.getElementById("facFirstNameId").value.trim();
	var facLastName = document.getElementById("facLastNameId").value.trim();
	var facToHaveAccount = "N";
	if (document.getElementById("facCheck").checked) {
		facToHaveAccount = "Y";
	}
	var model = Spine.Model.sub();
	model.configure("/admin/manageuser/saveFacilitatorPayment", "userGroup", "facilitatorEmail",
			"paymentDate", "numberOfUsers", "chequeNos", "totalAmountToBePaid", "facToHaveAccount",
			"createdBy", "paymentType", "facFirstName", "facLastName");
	model.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new model({  
		userGroup: userGrpVal,
		//bankDetails: bankDetails,
		facilitatorEmail: facilitatorEmail,
		paymentDate: paymentDate,
		numberOfUsers: numberOfUsers,
		chequeNos: chequeNos,
		totalAmountToBePaid: totalAmountToBePaid,
		facToHaveAccount: facToHaveAccount,
		createdBy: sessionStorage.getItem("jctEmail"),
		paymentType: "CHECK",
		facFirstName: facFirstName,
		facLastName: facLastName
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
		alertify.alert(obj.message);
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	    if (obj.statusCode != 123) {
	    	resetFields();
	    	fetchExistingUsersByUserTypeAndGroup();
	    }
	});
}

function saveFacilitatorCmpPayment() {
	var userGrpVal = document.getElementById("inputUserGroupInpt").options[document.getElementById("inputUserGroupInpt").selectedIndex].value;
	var facilitatorEmail = document.getElementById("facEmail").value;
	var numberOfUsers = document.getElementById("noOfUser2").value;
	var facFirstName = document.getElementById("facFirstNameId").value.trim();
	var facLastName = document.getElementById("facLastNameId").value.trim();
	var facToHaveAccount = "N";
	if (document.getElementById("facCheck").checked) {
		facToHaveAccount = "Y";
	}
	var model = Spine.Model.sub();
	model.configure("/admin/manageuser/saveFacilitatorPayment", "userGroup", "facilitatorEmail",
			"paymentDate", "numberOfUsers", "chequeNos", "totalAmountToBePaid", "facToHaveAccount",
			"createdBy", "paymentType", "facFirstName", "facLastName");
	model.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new model({
		userGroup: userGrpVal,
		//bankDetails: "",
		facilitatorEmail: facilitatorEmail,
		paymentDate: "",
		numberOfUsers: numberOfUsers,
		chequeNos: "",
		totalAmountToBePaid: 0,
		facToHaveAccount: facToHaveAccount,
		createdBy: sessionStorage.getItem("jctEmail"),
		paymentType: "FREE",
		facFirstName: facFirstName,
		facLastName: facLastName
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
		alertify.alert(obj.message);
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	    resetFields();
	    fetchExistingUsersByUserTypeAndGroup();
	});
}

function saveFacilitatorCashPayment(){
	var userGrpVal = document.getElementById("inputUserGroupInpt").options[document.getElementById("inputUserGroupInpt").selectedIndex].value;
	var facilitatorEmail = document.getElementById("facEmail").value;
	var numberOfUsers = document.getElementById("noOfUser2").value;
	var totalAmount = document.getElementById("cashAmountFaci").value.trim();
	var paymentDate = document.getElementById("paymentDtFaciCashInpt").value.trim();
	var facFirstName = document.getElementById("facFirstNameId").value.trim();
	var facLastName = document.getElementById("facLastNameId").value.trim();
	var facToHaveAccount = "N";
	if (document.getElementById("facCheck").checked) {
		facToHaveAccount = "Y";
	}
	var model = Spine.Model.sub();
	model.configure("/admin/manageuser/saveFacilitatorPayment", "userGroup", "facilitatorEmail",
			"paymentDate", "numberOfUsers", "chequeNos", "totalAmountToBePaid", "facToHaveAccount",
			"createdBy", "paymentType", "facFirstName", "facLastName");
	model.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new model({  
		userGroup: userGrpVal,
		//bankDetails: "",
		facilitatorEmail: facilitatorEmail,
		paymentDate: paymentDate,
		numberOfUsers: numberOfUsers,
		chequeNos: "",
		totalAmountToBePaid: totalAmount,
		facToHaveAccount: facToHaveAccount,
		createdBy: sessionStorage.getItem("jctEmail"),
		paymentType: "CASH",
		facFirstName: facFirstName,
		facLastName: facLastName
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
		alertify.alert(obj.message);
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	    resetFields();
	    fetchExistingUsersByUserTypeAndGroup();
	});	
}

/**
 * Function saves payment details for general user 
 * creation along with the mail ids
 */
function saveGeneralUserManualChkPayment(emailIdString) {
		var paymentDate = document.getElementById("datepicker").value;
		var numberOfUsers = document.getElementById("inputNoOFUser").options[document.getElementById("inputNoOFUser").selectedIndex].value;
		var chequeNos = document.getElementById("chequeno").value;
		var totalAmountToBePaid = document.getElementById("totalAmount").value;
		var userGrpVal = document.getElementById("inputUserGroupInpt").options[document.getElementById("inputUserGroupInpt").selectedIndex].value;
		var expiryDuration = document.getElementById("expiryDuration").options[document.getElementById("expiryDuration").selectedIndex].value;
		
		var model = Spine.Model.sub();
		model.configure("/admin/manageuser/saveGeneralUserPaymentManual", "userGroup", "emailIdString",
				"paymentDate", "numberOfUsers", "chequeNos", "totalAmountToBePaid", "expiryDuration",
				"createdBy", "paymentType");
		model.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new model({  
			userGroup: userGrpVal,
			//bankDetails: bankDetails,
			emailIdString: emailIdString,
			paymentDate: paymentDate,
			numberOfUsers: numberOfUsers,
			chequeNos: chequeNos,
			totalAmountToBePaid: totalAmountToBePaid,
			expiryDuration: expiryDuration,
			createdBy: sessionStorage.getItem("jctEmail"),
			paymentType: "CHECK"
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
			alertify.alert(obj.message);
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		    resetFields();
		    fetchExistingUsersByUserTypeAndGroup();
		});	
}
/**
 * Function saves csv entry for cheque payment
 */
function saveGeneralUserCSVChkPayment() {					//		FOR CSV
		var paymentDate = document.getElementById("datepicker").value;
			var chequeNos = document.getElementById("chequeno").value;
			var totalAmountToBePaid = document.getElementById("totalAmount").value;
			var userGrpVal = document.getElementById("inputUserGroupInpt").options[document.getElementById("inputUserGroupInpt").selectedIndex].value;
			var expiryDuration = document.getElementById("expiryDuration").options[document.getElementById("expiryDuration").selectedIndex].value;		
			var model = Spine.Model.sub();
			model.configure("/admin/manageuser/saveGeneralUserPaymentCSV", "userGroup", "bankDetails", "emailIdString",
					"paymentDate", "numberOfUsers", "chequeNos", "totalAmountToBePaid", "expiryDuration",
					"createdBy", "paymentType");
			model.extend( Spine.Model.Ajax );
			//Populate the model with data to transfer
			var modelPopulator = new model({  
				userGroup: userGrpVal,
				//bankDetails: bankDetails,
				emailIdString: emailIdCSVString,
				paymentDate: paymentDate,
				numberOfUsers: (totalUsers-1),
				chequeNos: chequeNos,
				totalAmountToBePaid: totalAmountToBePaid,
				expiryDuration: expiryDuration,
				createdBy: sessionStorage.getItem("jctEmail"),
				paymentType: "CHECK"
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
				alertify.alert(obj.message);
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			    resetFields();
			    fetchExistingUsersByUserTypeAndGroup();
			});
}
/**
 * Function saves csv entry for cheque payment
 */
function saveGeneralUserCSVCmpPayment() {
			var totalAmountToBePaid = 0;
			var userGrpVal = document.getElementById("inputUserGroupInpt").options[document.getElementById("inputUserGroupInpt").selectedIndex].value;
			var expiryDuration = document.getElementById("expiryDuration").options[document.getElementById("expiryDuration").selectedIndex].value;
		
			var model = Spine.Model.sub();
			model.configure("/admin/manageuser/saveGeneralUserPaymentCSV", "userGroup", "emailIdString",
					"paymentDate", "numberOfUsers", "chequeNos", "totalAmountToBePaid", "expiryDuration",
					"createdBy", "paymentType");
			model.extend( Spine.Model.Ajax );
			//Populate the model with data to transfer
			var modelPopulator = new model({  
				userGroup: userGrpVal,
				//bankDetails: "",
				//emailIdString: emailIdCSVString,
				emailIdString: emailIdCSVString,
				paymentDate: "",
				numberOfUsers: (totalUsers-1),
				chequeNos: "",
				totalAmountToBePaid: totalAmountToBePaid,
				expiryDuration: expiryDuration,
				createdBy: sessionStorage.getItem("jctEmail"),
				paymentType: "FREE"
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
				alertify.alert(obj.message);
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			    resetFields();
			    fetchExistingUsersByUserTypeAndGroup();
			});
}

/**
 * Function saves payment details for general user 
 * creation along with the mail ids
 */
function saveGeneralUserManualCmpPayment(emailIdString) {
		var numberOfUsers = document.getElementById("inputNoOFUser").options[document.getElementById("inputNoOFUser").selectedIndex].value;
		var userGrpVal = document.getElementById("inputUserGroupInpt").options[document.getElementById("inputUserGroupInpt").selectedIndex].value;
		var expiryDuration = document.getElementById("expiryDuration").options[document.getElementById("expiryDuration").selectedIndex].value;
		var model = Spine.Model.sub();
		model.configure("/admin/manageuser/saveGeneralUserPaymentManual", "userGroup", "emailIdString",
				"paymentDate", "numberOfUsers", "chequeNos", "totalAmountToBePaid", "expiryDuration",
				"createdBy", "paymentType");
		model.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new model({  
			userGroup: userGrpVal,
			//bankDetails: "",
			emailIdString: emailIdString,
			paymentDate: "",
			numberOfUsers: numberOfUsers,
			chequeNos: "",
			totalAmountToBePaid: 0,
			expiryDuration: expiryDuration,
			createdBy: sessionStorage.getItem("jctEmail"),
			paymentType: "FREE"
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
			alertify.alert(obj.message);
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		    resetFields();
		    fetchExistingUsersByUserTypeAndGroup();
		});
}

/**
* Function saves manual entry for cash payment
*/
function saveGeneralUserManualCashPayment(emailIdString){
		var numberOfUsers = document.getElementById("inputNoOFUser").options[document.getElementById("inputNoOFUser").selectedIndex].value;
		var userGrpVal = document.getElementById("inputUserGroupInpt").options[document.getElementById("inputUserGroupInpt").selectedIndex].value;
		var expiryDuration = document.getElementById("expiryDuration").options[document.getElementById("expiryDuration").selectedIndex].value;
		var totalAmount = document.getElementById("totalAmountIndi").value.trim();
		var paymentDate = document.getElementById("datepickerCashIndiinput").value.trim();
		var model = Spine.Model.sub();
		model.configure("/admin/manageuser/saveGeneralUserPaymentManual", "userGroup", "emailIdString",
				"paymentDate", "numberOfUsers", "chequeNos", "totalAmountToBePaid", "expiryDuration",
				"createdBy", "paymentType");
		model.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new model({  
			userGroup: userGrpVal,
			//bankDetails: "",
			emailIdString: emailIdString,
			paymentDate: paymentDate,
			numberOfUsers: numberOfUsers,
			chequeNos: "",
			totalAmountToBePaid: totalAmount,
			expiryDuration: expiryDuration,
			createdBy: sessionStorage.getItem("jctEmail"),
			paymentType: "CASH"
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
			alertify.alert(obj.message);
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		    resetFields();
		    fetchExistingUsersByUserTypeAndGroup();
		});
}
/**
* Function saves csv entry for cash payment
*/
function saveGeneralUserCSVCashPayment() {
		var totalAmountToBePaid = document.getElementById("totalAmountIndi").value;
		var userGrpVal = document.getElementById("inputUserGroupInpt").options[document.getElementById("inputUserGroupInpt").selectedIndex].value;
		var expiryDuration = document.getElementById("expiryDuration").options[document.getElementById("expiryDuration").selectedIndex].value;
		
		var model = Spine.Model.sub();
		model.configure("/admin/manageuser/saveGeneralUserPaymentCSV", "userGroup", "emailIdString",
				"paymentDate", "numberOfUsers", "chequeNos", "totalAmountToBePaid", "expiryDuration",
				"createdBy", "paymentType");
		model.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new model({  
			userGroup: userGrpVal,					
			//emailIdString: emailIdCSVString,
			emailIdString: emailIdCSVString,
			paymentDate: "",
			numberOfUsers: (totalUsers-1),
			chequeNos: "",
			totalAmountToBePaid: totalAmountToBePaid,
			expiryDuration: expiryDuration,
			createdBy: sessionStorage.getItem("jctEmail"),
			paymentType: "CASH"
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
			alertify.alert(obj.message);
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		    resetFields();
		    fetchExistingUsersByUserTypeAndGroup();
		});	
}
/**
 * function to add datePicker functionality to selected input 
 **/
$(document).ready(function () {
	 $.datepicker.setDefaults( 
	    {		    	  	
	     showOn: 'button', buttonImageOnly: true, buttonImage: '../img/datepicker-icon.png',dateFormat: "mm/dd/yy"});	 	
	 	$('#datepicker1, #datepicker, #datepickerCashIndiinput, #paymentDtFaciCashInpt').datepicker();
	 	
});

$(function() {
	$(".datepickerPastDate").datepicker("option", "minDate", 0);
 });
/**
 * Disable all other fields on load except select user type
 **/
function disableFieldsOnLoad(){
	document.getElementById("inputUserGroupInpt").setAttribute("disabled", "true");
	document.getElementById("manual").setAttribute("disabled", "true");
	document.getElementById("csv").setAttribute("disabled", "true");
	//document.getElementById("noOfUser1").setAttribute("disabled", "true");
	document.getElementById("inputNoOFUser").setAttribute("disabled", "true");
	document.getElementById("expiryDuration").setAttribute("disabled", "true");
	document.getElementById("paymentModeInd").setAttribute("disabled", "true");
}

/**
 * Fields only for individual user type
 **/
function fieldsForIndividual(){
	$("#paymentMode1")[0].selectedIndex = 0;
	document.getElementById("expiryDurationId").style.display = "block";
	$("#expiryDuration").selectedIndex = 0;
	document.getElementById("paymentModeId").style.display = "none";
	document.getElementById("inputUserGroupInpt").selectedIndex = "0";
	/*document.getElementById("inputUserGroupInpt").disabled = false;*/
	document.getElementById("paymentModeIdIndividual").style.display = "block";
	document.getElementById("radioButtonDivIndividual").style.display = "block";
	document.getElementById("facilitatorNameDiv").style.display = "none";
	document.getElementById("nosuser1").style.display = "block";
	document.getElementById("nosuser2").style.display = "none";
	document.getElementById("noOfUser2").style.display = "none";
	document.getElementById("qorkingArea1").style.display = "block";
	document.getElementById("qorkingArea2").style.display = "none";
	document.getElementById("inputNoOFUser").selectedIndex = "0";
	document.getElementById("facilitatorFirstNameDiv").style.display = "none";
	document.getElementById("facilitatorLastNameDiv").style.display = "none";
}

/**
 * Fields only for facilitator user type
 **/
function fieldsForFacilitator(){
	document.getElementById("expiryDurationId").style.display = "none";
	document.getElementById("paymentModeId").style.display = "block";
	document.getElementById("inputUserGroupInpt").selectedIndex = "1";
	document.getElementById("inputUserGroupInpt").disabled = true;
	document.getElementById("paymentModeIdIndividual").style.display = "none";
	document.getElementById("radioButtonDivIndividual").style.display = "none";
	document.getElementById("facilitatorNameDiv").style.display = "block";
	document.getElementById("nosuser1").style.display = "none";
	document.getElementById("nosuser2").style.display = "block";
	document.getElementById("noOfUser2").style.display = "block";
	document.getElementById("qorkingArea1").style.display = "none";
	document.getElementById("qorkingArea2").style.display = "none";		
	$("#paymentModeInd")[0].selectedIndex = 0;
	document.getElementById("facilitatorFirstNameDiv").style.display = "block";
	document.getElementById("facilitatorLastNameDiv").style.display = "block";
}

/**
 * Reset all field values
 **/
function resetFields(){
	document.getElementById("totalAmount").value = "";
	document.getElementById("chequeno").value = "";
	document.getElementById("datepicker").value = "";
	document.getElementById("totalAmountIndi").value = "";
	document.getElementById("datepickerCashIndiinput").value = "";
	document.getElementById("chkAmountFaci").value = "";
	document.getElementById("chequeNoFaci").value = "";
	document.getElementById("datepickerCashIndiinput").value = "";
	document.getElementById("cashAmountFaci").value = "";
	document.getElementById("paymentDtFaciCashInpt").value = "";
	document.getElementById("facEmail").value = "";
	document.getElementById("noOfUser2").value = "";
	document.getElementById("manualEntry").style.display = "none";
	$("#expiryDuration")[0].selectedIndex = 0;
	document.getElementById("paymentMode1").selectedIndex = "0";
	document.getElementById("inputNoOFUser").selectedIndex = "0";
	document.getElementById("paymentModeInd").selectedIndex = "0";
	document.getElementById("inputUserGroupInpt").selectedIndex = "0";	
	if(usrType == 2){	//	facilitator will be always under default grp
		document.getElementById("inputUserGroupInpt").selectedIndex = "1";	
	}
	document.getElementById("facFirstNameId").value = "";
	document.getElementById("facLastNameId").value = "";
}

/**
 * Enable all the disabled fields when user type selected
 **/
function enableAllFields(){
	document.getElementById("inputUserGroupInpt").removeAttribute("disabled", "true");
	document.getElementById("manual").removeAttribute("disabled", "true");
	document.getElementById("csv").removeAttribute("disabled", "true");
	document.getElementById("inputNoOFUser").removeAttribute("disabled", "true");
	document.getElementById("expiryDuration").removeAttribute("disabled", "true");
	document.getElementById("paymentModeInd").removeAttribute("disabled", "true");
	//document.getElementById("noOfUser1").removeAttribute("disabled", "true");
}


function validateCSVEntryOnAddUser(emailIdCSVString, selectionType) {
	// First Validate CSV entry
	if (validateUserData(emailIdCSVString, selectionType)) {
		$(".loader_bg").fadeIn();
		$(".loader").fadeIn();
		var expiryDuration = document.getElementById("expiryDuration").options[document.getElementById("expiryDuration").selectedIndex].value;
		
		var model = Spine.Model.sub();
		model.configure("/admin/manageuser/validateCSVEntryOnAddUser", "emailIdString", "selectionType", "expiryDuration");
		model.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new model({  			
			emailIdString: emailIdCSVString,
			selectionType: selectionType,
			expiryDuration: expiryDuration
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
		    if (statusCode == 215) { // All entries are correct
		    	populateSuccessValidation(obj);
		    } else {
		    	populateFailureValidation(obj);
		    }
		});
	}
}


function validateUserData(emailIdCSVString, selectionType) {
	if(selectionType == "csv") {
		if (emailIdCSVString == "") {
			alertify.alert("Please select the CSV file.");
			return false;
		} else if (emailIdCSVString.trim().length < 5){
			alertify.alert("CSV file is empty.");
			return false;
		} else {
			return true;
		}
	} else {
		var numberOfUsers = document.getElementById("inputNoOFUser").options[document.getElementById("inputNoOFUser").selectedIndex].value;	
		for (var index=1; index<=numberOfUsers; index++) {
			var emailString = document.getElementById("manualEmailId_"+index).value.trim();
			var fNameString = document.getElementById("firstNamelId_"+index).value.trim();
			var lNameString = document.getElementById("lastNamelId_"+index).value.trim();
			var filter = /^([a-zA-Z0-9_\.\-])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})+$/;
			if(null == emailString || emailString ==""){
				alertify.alert('Please provide a valid email address.');
				return false;
			} else if (!filter.test(emailString)) {
				alertify.alert('Please provide a valid email address.');
		    	return false;
		    } else if(null == fNameString || fNameString =="") {
		    	alertify.alert('Please provide first name.');
				return false;
		    } else if(null == lNameString || lNameString =="") {
		    	alertify.alert('Please provide last name.');
				return false;
		    }
		}			
	}
	return true;
}

function populateSuccessValidation(obj) {
	// Prepare message
	$('#myModal').modal('show');
	document.getElementById("message").innerHTML = obj.statusDesc;
	
	var msg = "";
	msg = msg + "<div align='center' class='error_msg_model'><table style='border: 1px solid black; width: 90%;'>";
	
	if (obj.validEmailIds.length > 0) {
		msg = msg + "<tr><td><table width='100%'><tr class='set-backgrnd-color'><td><b>User(s):</b></td><td><b>Expiration Date(s):</b></td></tr>";
		for (var index=0; index<obj.validEmailIds.length; index++) {
			var res = obj.validEmailIds[index].split("#`~`#");
			var expDt1 = res[1].toString().substring(0,10);
			var expDt2 = res[1].toString().substring(24,res[1].toString().length);
			var finalDate = expDt1+" "+expDt2;
			var expiryDate = "";
			var userExpryDate = "";
			expiryDate = new Date(finalDate).toDateString();
			userExpryDate = dateformat(new Date (expiryDate));
			msg = msg + "<tr><td>"+res[0]+"</td><td> "+userExpryDate+" </td>";
		}
		msg = msg + "</table></td></tr>";
	}
	msg = msg + "</table></div>";
	
	document.getElementById("tableData").innerHTML = msg;
	$(".loader_bg").fadeOut();
    $(".loader").hide();
}

function populateFailureValidation(obj) {
	// Prepare message
	$('#myModal1').modal('show');
	document.getElementById("message1").innerHTML = obj.statusDesc;
	
	var msg = "";
	msg = msg + "<div align='center' class='error_msg_model'><div style='border: 1px solid black; width:auto;height:auto;padding-bottom: 3%;'><table style='width: 100%;margin-top: -3%;'><tr>&nbsp;</tr>";
	if(obj.validEmailIds.length > 0) {
    	msg = msg + "<tr><td><table width='100%'><tr align='center'><td align='center'><b>Valid User(s):</b></th></tr>";
    	for (var index=0; index<obj.validEmailIds.length; index++) {
			msg = msg + "<tr><td align='center'>"+obj.validEmailIds[index]+"</td></tr>";
		}
    	msg = msg + "</table></td></tr>";
	}
	if (obj.invalidEmailIds.length > 0) {
		msg = msg + "<tr><td>&nbsp;</td></tr>";
    	msg = msg + "<tr><td><table width='100%'><tr align='center'><td align='center' colspan='2'><b>Invalid User(s):</b></td></tr>";
    	for (var index=0; index<obj.invalidEmailIds.length; index++) {
    		var res = obj.invalidEmailIds[index].split("#`~`#");
    		msg = msg + "<tr><td align='right' width='50%'>"+res[0]+"&nbsp;&nbsp;&nbsp;&nbsp;</td><td align='left'>"+res[1]+"</td></tr>";
		}
    	msg = msg + "</table></td></tr><tr>&nbsp;</tr>";
	}
	msg = msg + "</table></div></div>";
	document.getElementById("tableData1").innerHTML = msg;
	$(".loader_bg").fadeOut();
    $(".loader").hide();
}

function closeMyModel() {
	$('#myModal').modal('hide');
}


function closeMyModel1() {
	$('#myModal1').modal('hide');
}

function saveUserAfterValidation() {
	$('#myModal').modal('hide');
	var paymentMode = document.getElementById("paymentModeInd").options[document.getElementById("paymentModeInd").selectedIndex].value;
	if (paymentMode == "CHK") {	
			if (document.getElementById("manual").checked) {			 
				saveGeneralUserManualChkPayment(prepareEmailIdString()); 
			} else {													
				saveGeneralUserCSVChkPayment();
			}
	} else if (paymentMode == "CMP") {
			if (document.getElementById("manual").checked) {			
				saveGeneralUserManualCmpPayment(prepareEmailIdString()); 
			} else {													
				saveGeneralUserCSVCmpPayment();
			}
	} else if (paymentMode == "CASH") {
			if (document.getElementById("manual").checked) {			 
				saveGeneralUserManualCashPayment(prepareEmailIdString()); 
			} else {													 
				saveGeneralUserCSVCashPayment();
			}
	}
}

function validateDescription(event){
	var c = event.which || event.keyCode;
    if(( c == 33 ) || (c > 39 && c < 44) || (c > 46 && c < 58) || 
                  (c > 59 && c < 63) || (c > 63 && c < 65) || (c > 90 && c < 97) ||
                      (c > 122 && c !== 127))
       return false;	
}