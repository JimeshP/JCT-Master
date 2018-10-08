/**
 * Function calls when the page is loaded 
 * to populate the user profile list
 */
$(document).ready(function() {
	fetchUserProfileList();
	/***************** To Disable the paste functionality ISSUE# 186***************/
   /* $("#userProfileInptId").bind("paste",function(e) {
        e.preventDefault();
    });*/
});

/**
 * Function to fetch the user profile list
 * @param null
 */
function fetchUserProfileList() {
	var userProf = Spine.Model.sub();
	userProf.configure("/admin/manageuser/populateExistingUserProfile", "none");
	userProf.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userProf({  
		none: "" 
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
			populateUserProfileTable(obj.existingUserProfileList);
		} else if(statusCode == 201) {
			document.getElementById("existingUserProfileListId").innerHTML = "<div align='center'><img src='../img/no-record.png'><br />"+obj.statusDesc+"</div>";
		} else if(statusCode == 500) {
			//Show error message
		}
	});
}

/**
 * Function to save the user profile in database
 * @param null
 */
function saveUserProfile() {
	var userProfile = document.getElementById("userProfileInptId").value;	
	if(validateUserProfile(userProfile)) {
		var newUserProf = Spine.Model.sub();
		newUserProf.configure("/admin/manageuser/saveUserProfile", "userProfileVal", "createdBy");
		newUserProf.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new newUserProf({  
			userProfileVal: userProfile,
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
				alertify.alert("User Profile has been successfully created.");
				document.getElementById("userProfileInptId").value = "";
				populateUserProfileTable(obj.existingUserProfileList);
			} else if(statusCode == 600) {
				alertify.alert("User Profile already exists.");
			} else if (statusCode == 500) {
				//Error message
			}
		});
	}
}

/**
 * Function to validate the user profile text
 * @param userProfile
 */
function validateUserProfile(userProfile){
	var re=/^\s*$/;
	if (re.test(userProfile)){
		alertify.alert("Please enter User Profile.");
		return false;
	}	  
	return true;
}

/**
 * Function to show the user profile list in table view
 * @param userProfileListData
 */
function populateUserProfileTable(userProfileListData) {
	document.getElementById("userProfileInptId").value = "";
	var tableDiv = document.getElementById("existingUserProfileListId");
	var headingString = "<table width='94%' id='userProfileTableId' border=1 align='center' bordercolor='#78C0D3' class='tablesorter'><thead class='tab_header'><tr><th width='11%'><b>&nbsp;&nbsp; SL. No.</b></th><th><b>&nbsp;&nbsp; User Profile Name</b></th></tr></thead><tbody>";
	var splitRec = userProfileListData.split("~");
	var count = 1;
	var content = "";
	for(var index=0; index<splitRec.length-1; index++) {
		var rowColor = "";
		if(index%2 == 0){
			rowColor = "#D2EAF0";
		}
		content = content + "<tr align='center' bgcolor='"+rowColor+"' class='row_width'><td>&nbsp;&nbsp; "+count+".</td><td>&nbsp;&nbsp; "+splitRec[index]+"</td></tr>";
		count = count + 1;
	}
	headingString = headingString + content + "</tbody></table>";
	tableDiv.innerHTML = headingString;
	//new SortableTable(document.getElementById('userProfileTableId'), 1);
	$("table").tablesorter();
}
