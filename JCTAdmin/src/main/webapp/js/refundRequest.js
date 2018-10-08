$(document).ready(function() {
	//fetchAllDiagramCount();
	//fetchAllDiagramOnLoad();	
});

/**
 * Function searches the diagram which the users have shared 
 * in public database
 */
function searchUserForRefundRequest() {
	var emailId = document.getElementById("userName").value.trim();
	if (validateEmailId(emailId)) {

		$(".loader_bg").fadeIn();
		$(".loader").fadeIn();
		var userDiag = Spine.Model.sub();
		userDiag.configure("/admin/adminRemoveDiagram/searchUserForRefundRequest", "emailId");
		userDiag.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new userDiag({
			emailId: emailId
		});
		modelPopulator.save(); //POST
		userDiag.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("No Record found for this user.");
			$(".loader_bg").fadeOut();
			$(".loader").fadeOut();
		});
		userDiag.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);			
			if(obj.statusCodeInt == 200) {
				$(".loader_bg").fadeOut();
				$(".loader").fadeOut();
				populateTable(obj);
			} else if(obj.statusCodeInt == 500){
				alertify.alert("User has been disabled or not found.");
				$(".loader_bg").fadeOut();
				$(".loader").fadeOut();
				return false;
			}
		});
	}
}

function validateEmailId(emailId) {
	var emailReg = /^\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$/;
	if (!emailReg.test(emailId)){
		alertify.alert("Please enter valid E-mail ID.");
		return false;
	}
	return true;
}

function populateTable(existingUsers) {
	document.getElementById("existingUsersTableDiv").style.display = "block";
	var expiryDate = new Date(existingUsers.expiryDate).toDateString();
	var userExpryDate = dateformat(new Date (expiryDate));
	
	var counter = 1;
	var table = "";
	table = "<table width='94%' border='1' bordercolor='#78C0D3' id='myTable' align='center' class='tablesorter'><thead width='94%' class='tab_header'><tr><th>SL. No.</th><th>User Name</th><th>User Profile</th><th>User Group</th><th>Expiration Date</th><th width='22%'>Action</th></tr></thead>";
	table += "<tbody width='94%'><tr bgcolor='#D2EAF0' align='center'><td>"+counter+"</td><td>"+existingUsers.userName+"</td><td>"+existingUsers.userProfile+"</td><td>"+existingUsers.userGroup+"</td><td>"+userExpryDate+"</td><td><a href='#' onclick='disableUser(\""+existingUsers.jctUserId+"\", \""+existingUsers.userName+"\")'><b>Disable</b></a></td></tr></tbody></table>";
	
	document.getElementById("existingUsersTableId").innerHTML = table;
}

/**
 * function to disable existing user
 * @param userId
 * @param username
 * */	
function disableUser(userId, username) {
	$(".loader_bg").fadeIn();
	$(".loader").fadeIn();
	alertify.set({ buttonReverse: true });
		alertify.set({ labels: {
		    ok     : "Yes",
		    cancel : "No"
		} });
	alertify.confirm("<p>Are you sure you want to disable the user "+username+"</p>", function (e) {
	    if (e) {
	    	// soft delete the row with 2
	    	var userDiag = Spine.Model.sub();
			userDiag.configure("/admin/adminRemoveDiagram/disableUserForRefundRequest", "userId", "emailId");
			userDiag.extend( Spine.Model.Ajax );
			//Populate the model with data to transfer
			var modelPopulator = new userDiag({  
				userId: userId,
				emailId: username
			});
			modelPopulator.save(); //POST
			userDiag.bind("ajaxError", function(record, xhr, settings, error) {
				alertify.alert("Unable to connect to the server.");
			});
			userDiag.bind("ajaxSuccess", function(record, xhr, settings, error) {
				var jsonStr = JSON.stringify(xhr);
				var obj = jQuery.parseJSON(jsonStr);
				if (obj.statusCodeInt == 200) {
					$(".loader_bg").fadeOut();
					$(".loader").fadeOut();
					document.getElementById("existingUsersTableDiv").style.display = "none";
					document.getElementById("userName").value = "";
					alertify.set({ labels: {	//	again change the button caption to OK
					    ok     : "OK"
					} });
					alertify.alert("User has been disabled successfully");
				} 
			});
		} else {
			$(".loader_bg").fadeOut();
			$(".loader").fadeOut();
		}
	});
}