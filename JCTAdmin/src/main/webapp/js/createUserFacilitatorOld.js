 
var emailIdCSVString = "";

var registeredUser = "";
/**
 * Function calls when the page is loaded to populate 
 * 1. The existing activate deactivate user list
 * 2. The Total Subscribed User and total Registered User
 */
$(document).ready(function() {
	sessionStorage.removeItem("alreadyRegisteredEmails");
    sessionStorage.removeItem("invalidEmailIds");
	sessionStorage.removeItem("previous_user_email"); 
	fetchSubscribedUser();
	fetchUserGroupForFacilitator();
	fetchExistingUsers();
});


/**
 * Function to populate the Total Subscribed User
 * and total Registered User
 */
function fetchSubscribedUser() {
	var emailId = sessionStorage.getItem("jctEmail");
	var userGrp = Spine.Model.sub();
	userGrp.configure("/admin/manageuserFacilitator/populateSubscribedUser", "emailId", "type", "customerId");
	userGrp.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userGrp({  
		emailId: emailId,
		type: "AD",
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
			setSubscribedUser(obj.subscribedUsersList);
		}  else {
			//Show error message
			//alertify.alert(obj.statusDesc);
		}
	});
}

/**
 * Function to set value to the 
 * Total Subscribed User, Registered User 
 * and User Group fields
 * @param subscribedUsersList
 */
function setSubscribedUser(subscribedUsersList) {
	var splitRec = subscribedUsersList.split("~");
	document.getElementById("subscribedUserId").value = splitRec[2];
	document.getElementById("renewdedUserId").value = splitRec[1];
	//document.getElementById("userGroupInptId").value = splitRec[3];
	document.getElementById("expiryDateInputId").value = splitRec[4];
	registeredUser = splitRec[0] - splitRec[1] ;
	populateUserDropDown(splitRec[2]);
}
/**
 * Function add to populate
 * the no of user drop-down
 * @param noOfUser
 */
function populateUserDropDown(noOfUser){	
	document.getElementById("inputNoOFUser").innerHTML = "";
	var noOfUserSelect = document.getElementById("inputNoOFUser");		
	var defaultOptn = document.createElement("option");
	defaultOptn.text = "Select No of User";
	defaultOptn.value = "";
	defaultOptn.className = "form-control-general";
	noOfUserSelect.add(defaultOptn, null);
	if(noOfUser > 10){
		noOfUser = 10;
	}
	for(var i = 1; i<= noOfUser;i++){		
		var option = document.createElement("option");
		option.text = i;
	    option.value = i;
	    option.className = "form-control-general";
	    try {
	    	noOfUserSelect.add(option, null); //Standard 
	    }catch(error) {
	    	//regionSelect.add(option); // IE only
	    }	
	}		
}

/**
 * Function to populate the manual entry fields
 * while user click on manual entry
 */
function populateManualEntryField() {
	divCount = $('#manualEntry').children('.single_form_item:visible').length;
	var jsonObj = [];
	var counter = 0;
	for (var i = 1; i <= divCount; i += 1) {
	var unitJ = {};	
	var value = document.getElementById('manualEmailId_'+i).value;
	unitJ["ElementValue"+i] = value;
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
		alertify.alert("Please select number of users.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	} else {
		for (var index = 0; index < manualEntryNumber; index++) {
			var textFieldId = "manualEmailId_"+idCountIndex;
			string = string + "<div class='single_form_item'><div class='col-sm-4 '><label for='inputQtnNo' class='col-sm-12 control-label text_label'>User Name "+idCountIndex+":</label></div>" +
					"<div class='col-sm-6'><input type='text' class='form-control-general-form-redefined' maxlength='50' value='' placeholder='Enter Email Address' id='"+textFieldId+"'></div>" +
					"<div class='col-md-2'>&nbsp;</div><div class='clearfix'></div></div>";							
			idCountIndex = idCountIndex + 1;
		}
		plotPlement.innerHTML = string;
	}
	
	if(sessionStorage.getItem("previous_user_email") != null){
		var jsonObj = JSON.parse(sessionStorage.getItem("previous_user_email"));
		var count = jsonObj.length;
			for (var j = 0; j<count; j++){	
			var value = jsonObj[j]["ElementValue"+(j+1)];
			var div = document.getElementById('manualEmailId_'+(j+1));
			    if (div) {
				document.getElementById('manualEmailId_'+(j+1)).value = value;
				}			
			}		
		}
}



/**
 * Function to populate the existing users
 * under the facilitator 
 */
function fetchExistingUsers() {
	var userGrp = Spine.Model.sub();
	userGrp.configure("/admin/manageuserFacilitator/populateUsersByFacilitatorId", "emailId", "customerId");
	userGrp.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userGrp({  
		emailId: sessionStorage.getItem("jctEmail"),
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
		document.getElementById("existingList").style.display="block";
		var userList = document.getElementById("existingUsersTableId");
		var tableStr = "<table width='94%' border='1' bordercolor='#78C0D3' id='myTable' align='center'><thead class='tab_header'><tr><th>SL. No.</th><th>User Name</th><th>User Group Name</th><th>Expiration Date</th><th width='22%'>Action</th></tr></thead><tbody>";
		var counter = 1;
		for (var index=0; index<existingUsers.length; index++) {
			var trColor = "";
			if(index % 2 == 0) {
				trColor = "#D2EAF0";
			} else {
				trColor = "#F1F1F1";
			}
			var expiryDate = new Date(existingUsers[index].jctAccountExpirationDate).toDateString();
			var userExpryDate = dateformat(new Date (expiryDate));			
			
			tableStr = tableStr + "<tr class='user_list_row_width' bgcolor='"+trColor+"'><td align='center'>"+counter+".</td><td align='center'>"+existingUsers[index].email+"</td><td align='center'>"+existingUsers[index].userGroup+"</td><td align='center'>"+userExpryDate+"</td>";
			tableStr = tableStr + "<td class='table_col_txt_style' ><table width='100%' border='0'><tr>";
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
		document.getElementById("existingList").style.display="none";
		document.getElementById("existingUsersTableId").innerHTML = "<div align='center'><br /><br /><br /><img src='../img/no-record.png'><br /><div class='textStyleNoExist'>No Existing Users</div></div>";
	}
	
}


/**
 * Function to delete the users
 * @param emailId
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
 * @param emailId
 * @param softDelete
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
				fetchExistingUsers();
			});
		}
	});

}


/**
 * Function to reset the password of a particular user
 * @param emailId
 */
function resetPassword(emailId) {
	alertify.set({ buttonReverse: true });
	alertify.confirm("Are you sure you want to reset password of "+emailId+"?", function (e) {
		if (e) {
			// Blur the screen
			$(".loader_bg").fadeIn();
			$(".loader").fadeIn();
			
			var userDelete = Spine.Model.sub();
			userDelete.configure("/admin/manageuser/resetPassword", "emailId", "roleId");
			userDelete.extend( Spine.Model.Ajax );
			//Populate the model with data to transfer
			var modelPopulator = new userDelete({  
				emailId: emailId,
				roleId: 1
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
					alertify.alert("Password has been reset successful.");
				} else if (statusCode == 201) {
					alertify.alert("Password reset failed.");
				} else if (statusCode == 211) {
					alertify.alert("Sorry, did not get the mail id containing new password.");
				} else if (statusCode == 222) {
					alertify.alert(obj.statusDesc);
				}
				fetchExistingUsers();
			});
		}
	});
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
		document.getElementById("manualEntry").innerHTML = "";
		document.getElementById("inputNoOFUser").selectedIndex = "";
		emailIdCSVString = "";
	} else {
		document.getElementById("noOFuserDiv").style.display = "none";
		document.getElementById("manualEntry").style.display = "none";
		document.getElementById("fileUploadDiv").style.display = "block";
		
	}
}

/***
 * Function call when the CSV file is uoladed
 * check the extension of the file
 */
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
	/**
	 * New impl
	 */
	if (e.target.files != undefined) {
		var reader = new FileReader();
		reader.onload = function(e) {
		var csvval=e.target.result.split("\r");
		for(var index = 1; index<csvval.length; index++) {
			var csvvalue = csvval[index].split(",");			
			if(csvvalue[0].trim()){
				emailIds = emailIds + csvvalue[0]+"~";
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
 */
function saveUser() {
var userGrpVal = document.getElementById("userGroupInptId").options[document.getElementById("userGroupInptId").selectedIndex].value;
	if (document.getElementById("manual").checked) {
		var numberOfUsers = document.getElementById("inputNoOFUser").options[document.getElementById("inputNoOFUser").selectedIndex].value;
		if(userGrpVal.trim().length == 0){
			alertify.alert("Please select user group");
			return false;
		} else if(numberOfUsers == "") {			
		    alertify.alert("Please select number of users.");
			return false;
		} else {
			validateUserForSave(prepareEmailIdString(), registeredUser, "manual");
		}
	} else if (document.getElementById("csv").checked){
		 if(userGrpVal.trim().length == 0){
			alertify.alert("Please select an user group");
			return false;
		} else{
			validateUserForSave(emailIdCSVString, registeredUser, "csv");
		}		
	} else {		
	    alertify.alert('Please select an option first.');
	}
}


function validateAllUser(registeredUser, selectionType) {
	if (selectionType == "csv") {
		var splitRec = emailIdCSVString.split("~");
		if (emailIdCSVString == "") {
			alertify.alert("Please select the CSV file.");
			return false;
		} else if (emailIdCSVString.trim().length < 5){
			alertify.alert("CSV file is empty.");
			return false;
		} else if (registeredUser == 0){		
			alertify.alert("No more subscription.");
			return false;
		} else if(splitRec.length -1 > registeredUser) {
			alertify.alert("Cannot suscribe for "+(splitRec.length -1)+" users. Please update the user list.");
			return false;
		}
	} else {
		var numberOfUsers = document.getElementById("inputNoOFUser").options[document.getElementById("inputNoOFUser").selectedIndex].value;	
		for (var index=1; index<=numberOfUsers; index++) {
			var emailString = document.getElementById("manualEmailId_"+index).value.trim();
			var filter = /^([a-zA-Z0-9_\.\-])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})+$/;
			if(null == emailString || emailString ==""){
				alertify.alert('Please provide a valid email address.');
				return false;
			} else if (!filter.test(emailString)) {
				alertify.alert('Please provide a valid email address.');
		    	return false;
		    }
		}	
	}
	return true;
}

/**
 * Function to save the users
 * by manual entry
 */
function saveManualEntry(emailIdString) {
	var userGrpVal = document.getElementById("userGroupInptId").options[document.getElementById("userGroupInptId").selectedIndex].value;
	var expiryDate = document.getElementById("expiryDateInputId").value;
	var model = Spine.Model.sub();
	model.configure("/admin/manageuserFacilitator/saveManualUserFacilitator", "userGroup", "emailIdString", "createdBy", "expiryDate", "customerId");
	model.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new model({  
		userGroup: userGrpVal,
		emailIdString: emailIdString,
		createdBy: sessionStorage.getItem("jctEmail"),
		expiryDate: expiryDate,
		customerId : sessionStorage.getItem("customerId")
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
		reset();
		fetchSubscribedUser();
		fetchExistingUsers();
		
		$(".loader_bg").fadeOut();
	    $(".loader").hide();			
	});
}


/**
 * Reset all values
 */
function reset(){
	document.getElementById("manualEntry").innerHTML = "";
	document.getElementById("manualEntry").style.display = "none";
	 $('#userGroupInptId option').prop('selected', function() {
	        return this.defaultSelected;
	    });
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
	var userGrpVal = document.getElementById("userGroupInptId").value;
	var expiryDate = document.getElementById("expiryDateInputId").value;	
	var model = Spine.Model.sub();
	model.configure("/admin/manageuserFacilitator/saveCSVFacilitator", "userGroup", "emailIdString", "createdBy" ,"expiryDate", "customerId");
	model.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new model({  
		userGroup: userGrpVal,
		emailIdString: emailIdCSVString.substring(1, emailIdCSVString.length-1),
		createdBy: sessionStorage.getItem("jctEmail"),
		expiryDate: expiryDate,
		customerId : sessionStorage.getItem("customerId")
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
			fetchSubscribedUser();
			fetchExistingUsers();
			$(".loader_bg").fadeOut();
		    $(".loader").hide(); 	
	});		
}

function validateUserForSave(emailString, registeredUser, selectionType) {
	if(validateAllUser(registeredUser, selectionType)){
		$(".loader_bg").fadeIn();
		$(".loader").fadeIn();
		var emailIdString = "";
		if(selectionType == "csv") {
			emailIdString = emailIdCSVString.substring(1, emailIdCSVString.length-1);
		} else {
			emailIdString = emailString;
		}
		var expiryDate = document.getElementById("expiryDateInputId").value;
		var valModel = Spine.Model.sub();
		valModel.configure("/admin/manageuserFacilitator/validateUserDataFacilitator", "emailIdString", "expiryDate", "selectionType");
		valModel.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var valModelPopulator = new valModel({  
			emailIdString: emailIdString,
			expiryDate: expiryDate,
			selectionType: selectionType
		});
		valModelPopulator.save(); //POST
		valModel.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		});
		valModel.bind("ajaxSuccess", function(record, xhr, settings, error) {
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

function openDescPopupAll() {
	$(".alertify-cover").addClass("alertify-cover-hidden");
    $(".alertify").addClass("alertify-hide alertify-hidden");
    $('#myModal').modal('show');
   document.getElementById("myModalLabel").innerHTML = "Details";
    var regContent  = "";
    var invContent  = "";
   if(sessionStorage.getItem("alreadyRegisteredEmails") != null){
    	document.getElementById("info-div-header").style.display = 'block';
    	document.getElementById("info-div").style.display = 'block';
    	 var regEmails = sessionStorage.getItem("alreadyRegisteredEmails").split(',').join('<br>');
         for(var index=0; index<regEmails.length-1; index++){
        	 regContent = regContent + regEmails[index];
          }
        document.getElementById("info-div").innerHTML = regContent;
    }
    else{
    	document.getElementById("info-div-header").style.display = 'none';
    	document.getElementById("info-div").style.display = 'none';
    }
   
    if(sessionStorage.getItem("invalidEmailIds") != null){
        	if(sessionStorage.getItem("alreadyRegisteredEmails")==null){
    		$("#invalid_div_header").addClass("invalidDivHeaderTop");
    		$("#invalid-div").addClass("invalidDivTop");
    	}
    	else{
    		$("#invalid_div_header").removeClass("invalidDivHeaderTop");
    	}
    	document.getElementById("invalid_div_header").style.display = 'block';
    	document.getElementById("invalid-div").style.display = 'block';
    	var invEmails = sessionStorage.getItem("invalidEmailIds").split(',').join('<br>');
    	for(var index=0; index<invEmails.length-1; index++){
		invContent = invContent + invEmails[index];
    	}
       document.getElementById("invalid-div").innerHTML = invContent; 	
    }
    else{
    	document.getElementById("invalid_div_header").style.display = 'none';
    	document.getElementById("invalid-div").style.display = 'none';
    }
    $( ".custom_close_btn" ).click(function() {
           $(".alertify-cover").removeClass("alertify-cover-hidden");
           $(".alertify").removeClass("alertify-hide alertify-hidden");
    });
	
}
/**
 * Function to download the template of csv file
 */
function downloadCSVFmt() {
	var link = document.getElementById('downloadCSVFmt');
	link.setAttribute("href", "../../admin/manageuser/downloadCSVFile/Sample.csv");
}

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

function saveUserFacilitatorAfterValidation() {
	$('#myModal').modal('hide');
	if (document.getElementById("manual").checked) {
		saveManualEntry(prepareEmailIdString());
	} else {
		saveCSVEntry();
	}
}