/**
 * Function calls when the page is loaded to populate 
 * 1. The existing activate deactivate user list
 * 2. The Total Subscribed User and total Registered User
 */
$(document).ready(function() {
	fetchSubscribedUser();
	fetchUsersForRenew();
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
		type: "RW",
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
		}  else if(statusCode == 702) {
			setSubscribedUserToZero();
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
	document.getElementById("totalRenewedUserId").value = splitRec[0];
	registeredUser = splitRec[0] - splitRec[1] ;
	document.getElementById("availRenewedUserId").value = registeredUser;
	
}

/**
 * Function to set value to the 
 * Total Subscribed User, Registered User 
 * to zero
 * @param null
 */
function setSubscribedUserToZero() {
	document.getElementById("totalRenewedUserId").value = 0;
	document.getElementById("availRenewedUserId").value = 0;
	
}

/**
 * Function to populate the existing users
 * under the logged in facilitator 
 */
function fetchUsersForRenew() {
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
		document.getElementById("existingRenewList").style.display="block";
		var userList = document.getElementById("existingUsersTableId");
		var tableStr = "<table width='94%' border='1' bordercolor='#78C0D3' id='myTable' align='center' class='tablesorter'><thead class='tab_header'><tr><th>SL. No.</th><th>User Name</th><th>User Group Name</th><th>Expiration Date</th><th width='16%'>Action</th></tr></thead><tbody>";
		var counter = 1;
		for (var index=0; index<existingUsers.length; index++) {
			var trColor = "";
			if(index % 2 == 0) {
				trColor = "#D2EAF0";
			} else {
				trColor = "#F1F1F1";
			}
			var expiryDate = new Date(existingUsers[index].jctAccountExpirationDate).toDateString();	
			var d = new Date();
			//var currentDate = new Date(d).toLocaleDateString();		
			var currentDate = new Date(d).toDateString();		
			var date1 = new Date(expiryDate);
			var date2 = new Date(currentDate);
			var timeDiff = Math.abs(date2.getTime() - date1.getTime());
			var diffDays = Math.ceil(timeDiff / (1000 * 3600 * 24)); 
			var userExpryDate = dateformat(new Date (expiryDate));
					
			tableStr = tableStr + "<tr class='user_list_row_width' bgcolor='"+trColor+"'><td align='center'>"+counter+".</td><td align='center'>"+existingUsers[index].email+"</td><td align='center'>"+existingUsers[index].userGroup+"</td><td align='center'>"+userExpryDate+"</td>";
			tableStr = tableStr + "<td class='table_col_txt_style' >" +
					"<table width='100%' border='0'>" +
					"<tr>";
			if(date1 < date2) {
				tableStr = tableStr + "<td align='center'><a href='#' onclick='renewUser(\""+existingUsers[index].jctProfileId+"\", \""+existingUsers[index].email+"\")'>Renew</a></td></td></tr></table></td></tr>";
			} else if(diffDays >= 0 && diffDays <= 7){
				tableStr = tableStr + "<td align='center'><a href='#' onclick='renewUser(\""+existingUsers[index].jctProfileId+"\", \""+existingUsers[index].email+"\")'>Renew</a></td></td></tr></table></td></tr>";
			} else {
				/*tableStr = tableStr + "<td align='center'><span class='inactivate_style'>Renew</span></td></td></tr></table></td></tr>";*/
				tableStr = tableStr + "<td align='center'><a href='#' onclick='renewUser(\""+existingUsers[index].jctProfileId+"\", \""+existingUsers[index].email+"\")'>Renew</a></td></td></tr></table></td></tr>";
			}
						
			counter = counter + 1;
		}
		tableStr = tableStr + "</tbody></table>";
		userList.innerHTML = tableStr;
		$("table").tablesorter();
		new SortableTable(document.getElementById('myTable'), 1);
	} else {
		document.getElementById("existingRenewList").style.display="none";
		document.getElementById("existingUsersTableId").innerHTML = "<div align='center'><br /><br /><br /><img src='../img/no-record.png'><br /><div class='textStyleNoExist'>No Existing Users For Renewal</div></div>";
	}
	
}

/**
 * Function to renew the particular user account for next 6 months
 * @param userId
 * @param email
 */
function renewUser(userId, email) {
	var totalRenewUser = document.getElementById("totalRenewedUserId").value;
	var usersToRenew = document.getElementById("availRenewedUserId").value;	
	if (validateRenew(totalRenewUser, usersToRenew)) {
	alertify.set({ buttonReverse: true });
	alertify.confirm("<b> Renewing will add  6 months to your expiration date.</b><br>"+
					"Are you sure you want to renew account for user "+email+"?", function (e) {
		if (e) {
			$(".loader_bg").fadeIn();
			$(".loader").fadeIn();	
			var userRenew = Spine.Model.sub();
			userRenew.configure("/admin/manageuserFacilitator/renewUserByFacilitator", "userId", "userEmail", "facilitatorEmail", "usersToRenew", "customerId");
			userRenew.extend( Spine.Model.Ajax );
			//Populate the model with data to transfer
			var modelPopulator = new userRenew({  
				userId: userId,
				userEmail: email,
				facilitatorEmail: sessionStorage.getItem("jctEmail"),
				usersToRenew: usersToRenew,
				customerId : sessionStorage.getItem("customerId")
			});
			modelPopulator.save(); //POST
			userRenew.bind("ajaxError", function(record, xhr, settings, error) {
				alertify.alert("Unable to connect to the server.");
			});
			userRenew.bind("ajaxSuccess", function(record, xhr, settings, error) {
				var jsonStr = JSON.stringify(xhr);
				var obj = jQuery.parseJSON(jsonStr);
				var statusCode = obj.statusCode;
				if(statusCode == 200) {
					alertify.alert(obj.statusDesc);
					fetchSubscribedUser();
				    fetchUsersForRenew();
					$(".loader_bg").fadeOut();
				    $(".loader").hide();	    
				}  else {
					//Show error message
					alertify.alert("Some thing went wrong.");
				}
			});
		}
	});
	}
}

function validateRenew(totalRenewUser, usersToRenew) {
	if(totalRenewUser == 0) {
		alertify.alert("Please ensure you have sufficient balance in your account to renew the subscription.");
		return false;
	} else if (usersToRenew == 0) {
		alertify.alert("Please ensure you have sufficient balance in your account to renew the subscription.");
		return false;
	}
	return true;
}


function goToNextPage() {
	var userToRenew = document.getElementById("availRenewedUserId").value;
	sessionStorage.setItem("availableUserToRenew", userToRenew);
	var totalRenewUser = document.getElementById("totalRenewedUserId").value;
	var usersToRenew = document.getElementById("availRenewedUserId").value;	
	if(totalRenewUser == 0) {
		alertify.alert("Please ensure you have sufficient balance in your account to renew the subscription.");	
		return false;
	} else if (usersToRenew == 0) {
		alertify.alert("Please ensure you have sufficient balance in your account to renew the subscription.");
		return false;
	} else {
		window.location = "bulkRenewFacilitator.jsp";
	}
	return true;
}