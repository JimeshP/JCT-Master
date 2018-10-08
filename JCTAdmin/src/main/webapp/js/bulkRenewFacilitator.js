var emailIdCSVString = "";
var totalUsers = 0;
var map = new Object();

$(document).ready(function() {
	fetchUsersForRenew();
});


function renewBulkUserByFacilitator() {		
	//var userToRenew = sessionStorage.getItem("availableUserToRenew");
	var userToRenew = document.getElementById("availRenewedUserId").value;
	var emailIdString = "";
	var totalUsers=$('input[name="email"]:checked').length;
	var checkBoxes=document.getElementsByName("email");
	for (var i=0; i<checkBoxes.length; i++) {
		 if (checkBoxes[i].checked) {
			 if (checkBoxes.length-1 == i) {
				 emailIdString = emailIdString + checkBoxes[i].value;
			 } else {
				 emailIdString = emailIdString + checkBoxes[i].value + "~";
			 }
		 }
	}
	if (emailIdString == "") {
		alertify.alert("Please select at least one user.");
	} else if(totalUsers > userToRenew) {
		alertify.alert("Only "+userToRenew+" users can be renewed.");
	} else {
		populateSuccessValidation(emailIdString,totalUsers);
	}
}


function populateSuccessValidation(emailIdString,totalUsers) {
	// Prepare message
	$('#bulkRenewModal').modal('show');
	document.getElementById("renewUser").innerHTML = "You have selected "+totalUsers+" user(s) below. Would you like to <br>renew these users?<br>";
	var msg = "";
	msg = msg + "<div align='center'> <table style='border: 1px solid black; width: 90%;'><tr class='set-backgrnd-color'><td align='center'><b>User(s):</b></td><td align='center'><b>New Expiration Date(s):</b></td></tr>";
	var checkBoxes=document.getElementsByName("email");
	for (var i=0; i<checkBoxes.length; i++) {
		 if (checkBoxes[i].checked) {	
			var expiryDate = new Date(map[checkBoxes[i].value]);
			var expiryDate1= expiryDate.setMonth(expiryDate.getMonth() + 6);
			var newExpiryDate = dateformat(new Date (expiryDate1));								
			msg = msg + "<tr><td align='center'>"+checkBoxes[i].value+"</td><td align='center'>"+newExpiryDate+"</td></tr>";
		 }
	}
	msg = msg + "</table></div>";	
	document.getElementById("tableData").innerHTML = msg;
}


function renewBulkUserAfterSuccess() {
	$('#bulkRenewModal').modal('hide');
	var emailIdString = "";
	var totalUsers=$('input[name="email"]:checked').length;
	var checkBoxes=document.getElementsByName("email");
	for (var i=0; i<checkBoxes.length; i++) {
		 if (checkBoxes[i].checked) {
			 if (checkBoxes.length-1 == i) {
				 emailIdString = emailIdString + checkBoxes[i].value;
			 } else {
				 emailIdString = emailIdString + checkBoxes[i].value + "~";
			 }
		 }
	}
	var model = Spine.Model.sub();
	model.configure("/admin/manageuserFacilitator/renewBulkUserByFacilitator", "emailIdString", "totalUsers", "createdBy", "customerId");
	model.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new model({  			
		emailIdString: emailIdString,		
		totalUsers: totalUsers,
		createdBy: sessionStorage.getItem("jctEmail"),
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
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	    if(obj.statusCode == 200) {
	    	alertify.alert((totalUsers)+" user(s) has been renewed successfully.");
	    	fetchUsersForRenew();
	    	fetchSubscribedUser();
	    } else {
	    	//error
	    }
	});
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
		var userList = document.getElementById("mainPopulationDiv");
		var tableStr = "<table width='94%' border='1' bordercolor='#78C0D3' id='myTable' align='center'><thead class='tab_header'><tr><th width='14%'><input type='checkbox' id='select-all' name='selectAll' value='selectAll'>&nbsp;Select All</th><th>User Name</th><th>User Group Name</th><th>Expiration Date</th></tr></thead><tbody>";
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
			tableStr = tableStr + "<tr class='user_list_row_width' bgcolor='"+trColor+"'><td> <input type='checkbox' name='email' value='"+existingUsers[index].email+"' class='checkBox_align'> </td><td align='center'>"+existingUsers[index].email+"</td><td align='center'>"+existingUsers[index].userGroup+"</td><td align='center'>"+userExpryDate+"</td>";			
			counter = counter + 1;			
			map[existingUsers[index].email] = existingUsers[index].jctAccountExpirationDate;
		}
		tableStr = tableStr + "</tbody></table>";
		userList.innerHTML = tableStr;
		//$("table").tablesorter(); 
		
		document.getElementById("buttonDivId").style.display = "block";
		document.getElementById("buttonDivId").innerHTML = "<table align='center'><tr><td style='padding:5px 0 0 0;'><input class='btn_admin btn_admin-primary btn_admin-sm search_btn' type='button' id='theButton' value='Renew' onclick='renewBulkUserByFacilitator()'></td></tr></table>";
	} else {
		document.getElementById("existingRenewList").style.display="none";
		document.getElementById("existingUsersTableId").innerHTML = "<div align='center'><br /><br /><br /><img src='../img/no-record.png'><br /><div class='textStyleNoExist'>No Existing Users For Renewal</div></div>";
	}
	
	$("#select-all").on("click", function() {
		  var all = $(this);
		  $('input:checkbox').each(function() {
		       $(this).prop("checked", all.prop("checked"));
		  });
	});
}


function populateUserModel() {	
	/*$(".alertify-cover").addClass("alertify-cover-hidden");
	$(".alertify").addClass("alertify-hide alertify-hidden");
	$('#myModal').modal('show');
	document.getElementById("myModalLabel").innerHTML = "Already Registered Users";
	document.getElementById("instruction_panel_text").innerHTML = sessionStorage.getItem("invalidEmailIds");
	document.getElementById("instruction_panel_video").innerHTML = sessionStorage.getItem("notExistEmails");
	
	$( ".custom_close_btn" ).click(function() {
		$(".alertify-cover").removeClass("alertify-cover-hidden");
		$(".alertify").removeClass("alertify-hide alertify-hidden");
	});*/	
	$(".alertify-cover").addClass("alertify-cover-hidden");
    $(".alertify").addClass("alertify-hide alertify-hidden");
    $('#myModal').modal('show');
    document.getElementById("myModalLabel").innerHTML = "Details";
    var regContent  = "";
    var invContent  = "";
    
   if(sessionStorage.getItem("notExistEmails") != null){
    	document.getElementById("info-div-header").style.display = 'block';
    	document.getElementById("info-div").style.display = 'block';
    	 var regEmails = sessionStorage.getItem("notExistEmails").split(',').join('<br>');
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
        	if(sessionStorage.getItem("notExistEmails")==null){
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

function closeMyModel() {
	$('#bulkRenewModal').modal('hide');
}