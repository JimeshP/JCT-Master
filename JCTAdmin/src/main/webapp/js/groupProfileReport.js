/**
 * Function calls when the page is loaded 
 * to populate the drop-down for user profile
 */
$(document).ready(function() {
	populateDropdowns();
});

/**
 * Function to populate the drop-down for user profile
 * @param null
 */
function populateDropdowns(){
	var option = null;
	var obj = jQuery.parseJSON(sessionStorage.getItem("jsonStr"));
	var usrProfileSelect = document.getElementById("inputUserProfileInpt");
	for (var key in obj.userProfile) {
		//Create a option element dynamically
		option = document.createElement("option");
		option.text = obj.userProfile[key];
	    option.value = key;
	    option.className = "form-control-general";
	    if(sessionStorage.getItem("fnLSelected")){
	    	if(option.value == sessionStorage.getItem("fnLSelected")){
		    	option.setAttribute("selected", "selected");
		    }
	    }
	    try {
	    	usrProfileSelect.add(option, null); //Standard 
	    }catch(error) {
	    }
	}
}

/**
 * Function call to generate the excel report 
 * @param null
 */
function generateXLReport(){
	var link = document.getElementById('generateXL');
	var userProfile = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value;	
	if(validateUserProfile(userProfile)) {
		var reportLink = "../../admin/reports/generateExcelGroupProfileReport/";			
		reportLink = reportLink + userProfile + "/report.xls";		
		link.setAttribute("href", reportLink);
	} else {
		link.removeAttribute("href");
	}
}

/**
 * Function call to validate user profile
 * @param userProfile
 */
function validateUserProfile(userProfile){
	if(document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value == ""){
		alertify.alert("Please select User Profile.");
		return false;
	}
	return true;
}