var mainHeight = $(".main_contain").height();
/**
 * Function calls when the page is loaded 
 * to populate the drop-down for user profile
 */
$(document).ready(function() {
	fetchDropDownValues();	
});

function fetchTnC(){
	var userProfile = $("#inputUserProfileInpt").val();
	var userType = $("#userType").val();
	if(userProfile.length < 1){
		CKEDITOR.instances.editor1.setData('');
		return false;
		
	}
	
	var userProf = Spine.Model.sub();
	userProf.configure("/admin/contentconfig/fetchTermsAndConditions", "userProfile", "userType");
	userProf.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userProf({  
		userProfile: userProfile,
		userType: userType
	});	
	modelPopulator.save(); //POST
	userProf.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	userProf.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		//var obj = jQuery.parseJSON(jsonStr);
		var obj = JSON.parse(jsonStr);
		var statusCode = obj.statusCode;
		if(statusCode == 200) {
			putTcValue(obj.jctTcDescription);										
		} else if(statusCode == 201) {
			//No data found
			CKEDITOR.instances.editor1.setData('');	
		} else if(statusCode == 500) {
			//Show error message
		}
	});
	$( 'html, body, document' ).stop().animate( { scrollTop: 0 }, 'slow' );	
} 



/**
 * Function to fetch the 
 * user profile drop down
 * @param null
 */
function fetchDropDownValues() {
	var userProf = Spine.Model.sub();
	userProf.configure("/admin/contentconfig/populateUserProfileGeneral", "none");
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
		//var obj = jQuery.parseJSON(jsonStr);
		var obj = JSON.parse(jsonStr);
		var statusCode = obj.statusCode;
		if(statusCode == 200) {
			populateDropDown(obj.userProfileMap);										
		} else if(statusCode == 201) {
			//Show error message
		} else if(statusCode == 500) {
			//Show error message
		}
	});
	$( 'html, body, document' ).stop().animate( { scrollTop: 0 }, 'slow' );
}


/**
 * Function to show the terms & condition
 * for the selected profile and related page
 * @param jctTcDescription
 */
function putTcValue(jctTcDescription) {
	CKEDITOR.instances['editor1'].insertHtml(jctTcDescription);
}

/**
 * Function call if the instruction data is not in db 
 * for the selected profile and related page
 */
function populateNone(){
	CKEDITOR.instances['editor1'].insertHtml('&nbsp;');
	document.getElementById("instructionUpdateDiv").setAttribute("style", "display:none");
	document.getElementById("instructionAddDiv").setAttribute("style", "display:block");
	var userProfileSelect = document.getElementById("inputUserProfileInpt");
	userProfileSelect.removeAttribute("disabled", "true");
	var relatedPage = document.getElementById("inputPageInpt");
	relatedPage.removeAttribute("disabled", "true");
}

/**
 * Function add to populate
 * the user profile drop-down value
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
 * Function to Save/Update the T&C to the database
 */
function saveTC() {
	var tcText = CKEDITOR.instances.editor1.getData();
	var userProfile = $("#inputUserProfileInpt").val();
	var userType = $("#userType").val();
	if(userType == 0){		
		alertify.alert("Please select user type");
		return false;
	}
	if(userProfile.length < 1){
		alertify.alert("Please select User Profile.");
		return false;		
	} else if (null == tcText || tcText == "" || tcText == "&nbsp;"){
		alertify.alert("Please enter Terms & Conditions text data.");
		return false;
	}	
	var newUserProf = Spine.Model.sub();
	newUserProf.configure("/admin/contentconfig/saveAndUpdateTermAndCondition", "tcText", "userProfile","userType", "createdBy");
	newUserProf.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new newUserProf({  
		tcText: tcText.replace(/\n/g,'<br>'),
		userProfile: userProfile,
		userType: userType,
		createdBy: sessionStorage.getItem("jctEmail"),
	});
	modelPopulator.save(); //POST
	newUserProf.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	newUserProf.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = JSON.parse(jsonStr);
		var statusCode = obj.statusCode;
		if(statusCode == 200) {
			alertify.alert("Terms and Conditions has been saved/updated successfully.");
			clearAllData();
		} else if(statusCode == 500) {
			// Error Message
		}
	});
}

/**
 * Function to clear all data from the page
 */
function clearAllData() {	
	$("#inputUserProfileInpt")[0].selectedIndex = 0;
	$("#userType")[0].selectedIndex = 0;
	document.getElementById("inputUserProfileInpt").disabled = true;
	CKEDITOR.instances.editor1.setData('');	
}
function enableUsrProfile(){
	var usrType = $("#userType").val();
	CKEDITOR.instances.editor1.setData('');
	if(usrType == 0){		
		clearAllData();
		return false;
	} else {
		document.getElementById("inputUserProfileInpt").disabled = false;
	}
	
	
}