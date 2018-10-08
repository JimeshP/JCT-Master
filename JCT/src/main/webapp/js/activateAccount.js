var inputTextSurveyId = "";
var inputRadioSurveyId = "";
var inputDropSurveyId = "";
var inputCheckSurveyId = "";
var inputTextSurveyAns = "";
var inputRadioSurveyAns = "";
var inputDropSurveyAns = "";
var inputCheckSurveyAns = "";

var plottedText = 0;
var plottedRadio = 0;
var plottedDrop = 0;
var plottedCheck = 0;


/**
 * Function add to disable browser back button
 * while the page is loaded
 * @param null
 */
window.location.hash="";
window.location.hash="Again-No-back-button";//again because google chrome don't insert first hash into history
window.onhashchange=function(){window.location.hash="";};


$(document).ready( function() {
	storeParamValue();
	decideDivToDisplay();	
});

/**
 * Function to allow only char value and .
 */
$("#inputFN, #inputLN").keyup(function () {	
    $(this).val( $(this).val().replace(/([~!@#$%^&*()_\-+=`{}\[\]\|\\:;'<>,.\/?])+/g, function(){ return ''; }));
});
$("#ocptnSrchTxtId").keyup(function () {	
    //$(this).val( $(this).val().replace(/([*()\+\\{}\[\]\|\/?])+/g, function(){ return ''; }));
});
function validateDescription(event){
	var c = event.which || event.keyCode;
    if(( c == 33 ) || (c > 39 && c < 44) || (c > 46 && c < 58) || 
                  (c > 59 && c < 63) || (c > 63 && c < 65) || (c > 90 && c < 97) ||
                      (c > 122 && c !== 127))
       return false;	
}
function validateNumber(event){
	var c = event.which || event.keyCode;
	if(( c < 48 ) || (c > 57))
       return false;	
}
/**
 * Function fetches the email id from the activation link (url) and stores it in the session.
 */
function storeParamValue() {	
	sessionStorage.clear();
	//sessionStorage.setItem(profileId);
	var parameter = window.location.search.replace( "?", "" ); // will return the GET parameter 
	var values = parameter.split("=");
	sessionStorage.setItem("jctEmail", values[1].split("&")[0]);
	//toogleDiv(values[2]);
}
/**
 * Function decides whether to ask the user to reset the password or to fill up the login details.
 */
function decideDivToDisplay() {
	$(".loader_bg").fadeIn();
    $(".loader").fadeIn();
	//Create a model
	var model = Spine.Model.sub();
	model.configure("/user/auth/decidePage", "emailId");
	model.extend( Spine.Model.Ajax );
	
	//Populate the model
	var modelPopulator = new model({  
		emailId: sessionStorage.getItem("jctEmail")
	});
	//POST
	modelPopulator.save();
	
	model.bind("ajaxError", function(record, xhr, settings, error) {
		//var jsonStr = JSON.stringify(xhr);
		alertify.alert("Unable to connect to the server. Please wait for some time and try again.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	});
	
	model.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
		if (obj.statusCode == "776") {
			alertify.alert(obj.statusDesc);
		} else {
			if (obj.accountExpWarning) {
				sessionStorage.setItem("accountExp", obj.accountExpWarning);
			}
			if (obj.accountExpWarning) {
				sessionStorage.setItem("firstName", obj.firstName);
			}
			if (obj.accountExpWarning) {
				sessionStorage.setItem("lastName", obj.lastName);
			}
			sessionStorage.setItem("profileId", obj.profileId);
			toogleDiv(obj.isCompleted);
			
		}
	});
}
/**
 * Function activates the user status [2] and opens another div so that user can fill in manadtory details.
 */
function activateAccount() {
	var selectedUserName = sessionStorage.getItem("jctEmail");
	var mailedPwd = $('#mailedPassword').val();
	var newPassword = $('#inputNPwd').val();
	var confirmPassword = $('#inputCPwd').val();
	if(validatePasswordEntries(mailedPwd, newPassword, confirmPassword) == true) {
		$(".loader_bg").fadeIn();
	    $(".loader").fadeIn();
	    
		//Create a model
		var passwordResetModel = Spine.Model.sub();
		passwordResetModel.configure("/user/auth/resetPassword", "forUser", "initialPassword", "userPassword", 
				"status", "clickStatus", "dist", "mailedPwd");
		passwordResetModel.extend( Spine.Model.Ajax );
		
		//Populate the model
		var modelPopulator = new passwordResetModel({  
			forUser: selectedUserName,
			initialPassword: encode(mailedPwd+mailedPwd),
			userPassword: encode(newPassword+newPassword),
			status: "2",
			clickStatus: "activation",
			dist: 1,
			mailedPwd : mailedPwd			
		});
		//POST
		modelPopulator.save();
		
		passwordResetModel.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			var statusCode = obj.statusCode;
			if(statusCode == 200){ //success
				sessionStorage.setItem("statusMsg", obj.statusDesc);
				document.getElementById('resetPwdId').innerHTML = "Registration Details";
				document.getElementById('passwordResetForm').style.display = "none";
				document.getElementById('loginForm').style.display = "block";
				sessionStorage.setItem("isReset","Y");
				sessionStorage.setItem("profileId", obj.profileId);
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
				//showDescOnLoad();
			    populateSignupMasterData();
			}else if(statusCode == 404){
				sessionStorage.setItem("statusMsg", obj.statusDesc);
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
				//showDescOnLoad();
			    populateSignupMasterData();
			}
			else if(statusCode == 500){ //Jira JCTP-35 Checking of mailed password with input password

				$(".loader_bg").fadeOut();
			    $(".loader").hide();
				alertify.alert(obj.statusDesc);
			} else if(statusCode == 128){
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			    sessionStorage.setItem("isNextReset", "Y");
			    window.location = "resetPassword.jsp";
			}  
			//returns 200: success
			//1: change the name of the title
			
			
		});
		
		passwordResetModel.bind("ajaxError", function(record, xhr, settings, error) {
			//var jsonStr = JSON.stringify(xhr);
			alertify.alert("Unable to connect to the server.");
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		});
		
	}
}



/**
 * function validates the reset password entries
 * which are provided by the user.
 * @param mailedPassword
 * @param newPassword
 * @param confirmPassword
 * @returns {Boolean}
 */

/* Jira JCTP-35*/
function validatePasswordEntries(mailedPassword, newPassword, confirmPassword){
	
	var illegalChars = /^(?=.*\d)(?=.*[a-z])(?=.*[A-Z])(?=.*[^a-zA-Z0-9])(?!.*\s).{8,16}$/;
	if (mailedPassword == "") {
		alertify.alert('Please enter the Password which we mailed you.');
        return false;
    }   else if (newPassword == "") {
    	alertify.alert('Please enter the New Password.');
        return false;
    } else if ((newPassword.length < 8) || (newPassword.length > 16)) {
    	alertify.alert('New Password should be between 8 - 16 characters.');
        return false;
    } else if (!illegalChars.test(newPassword)) {
    	alertify.alert('New Password should contain atleast one lowercase letter, one uppercase letter, one numeric digit, and one special character.');
        return false;
    } else if (confirmPassword == "") {	//JIRA Defect Id: VMJCT-31
    	alertify.alert('Please enter the Confirm Password.');
        return false;
    } else if ((confirmPassword.length < 8) || (confirmPassword.length > 16)) {
    	alertify.alert('Confirm Password should be between 8 - 16 characters.');
        return false;
    } /*else if (!illegalChars.test(confirmPassword)) { //JIRA Defect Id: VMJCT-227
    	alertify.alert('Confirm Password should contain atleast one lowercase letter, one uppercase letter, one numeric digit, and one special character.');
        return false;
    }*/ else if (newPassword != confirmPassword){
    	alertify.alert('New Password is different from Confirm Password. Please correct it.');
        return false;
    } else if (mailedPassword == newPassword) {
    	alertify.alert('Your password should be different from the one we sent you.');
        return false;
    }
	return true;
}
/**
 * function encodes the new password which is 
 * then sent over the network
 * @param text
 * @returns {String}
 */
function encode(text) { 
	var ref = "P!QR#S$T%U&VW(X)Y*Z+[,-].^/_0`1a2b3c4d5e6f7g8h9i:j;k<l=m>n?o@pAqBrCsDtEuFvGwHxIyJzK{L|M}N~O";
	var result="";
	for (var count=0; count<text.length; count++) {
	    var char=text.substring (count, count+1); 
	    var num=ref.indexOf (char);
	    var encodeChar=ref.substring(num+1, num+2);
	    result += encodeChar;
	}
	return result;
}
/**
 * Method populates master dropdown for location, function group and job level
 */
function populateSignupMasterData(){
	//Create a model
	var master = Spine.Model.sub();
	master.configure("/user/auth/populateSignupMasterData", "none");
	master.extend( Spine.Model.Ajax );
	
	//Populate the model
	var populateMaster = new master({  
		none: ""
	});
	populateMaster.save(); //POST
	master.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to fetch the dropdown values.");
	});
	
	master.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);		
		var statusCode = obj.statusCode;
		if(statusCode == "200"){
			
			document.getElementById("inputFN").value=sessionStorage.getItem("firstName");
			document.getElementById("inputLN").value=sessionStorage.getItem("lastName");
			
			var nosOfYrsInput = document.getElementById("nosOfYrsInput");
			for (var i = 1; i<= 65; i++) {
				var option = document.createElement("option");
				if (i == 1) {
					option.text = i+" year";
					option.value = i+" year";
				} else {
					option.text = i+" years";
					option.value = i+" years";
				}
				option.className = "form-control-general";
				try {
					nosOfYrsInput.add(option, null); //Standard 
			    }catch(error) {
			    	//regionSelect.add(option); // IE only
			    }
			}
			
			
			//get select element for region
			var regionSelect = document.getElementById("inputLocation");
			for (var key in obj.region) {				
				//Create a option element dynamically
				var option = document.createElement("option");
				option.text = obj.region[key];
			    option.value = obj.region[key];
			    option.className = "form-control-general";
			    try {
			    	regionSelect.add(option, null); //Standard 
			    }catch(error) {
			    	//regionSelect.add(option); // IE only
			    }
			}
			
			var functionGrpSelect = document.getElementById("inputFunctionGroup");
			for (var key in obj.functionGroup) {
				//Create a option element dynamically
				var option = document.createElement("option");
				option.text = obj.functionGroup[key];
			    option.value = obj.functionGroup[key];
			    option.className = "form-control-general";
			    try {
			    	functionGrpSelect.add(option, null); //Standard 
			    }catch(error) {
			    	//functionGrpSelect.add(option); // IE only
			    }
			}
			
			var jobLevelSelect = document.getElementById("inputJobLevel");
			for (var key in obj.jobLevel) {
				//Create a option element dynamically
				var option = document.createElement("option");
				option.text = obj.jobLevel[key];
			    option.value = obj.jobLevel[key];
			    option.className = "form-control-general";
			    try {
			    	jobLevelSelect.add(option, null); //Standard 
			    }catch(error) {
			    	//jobLevelSelect.add(option); // IE only
			    }
			}
			
			var plottingDiv = document.getElementById("surveyQuestionsDivId");
			var textSplit = obj.surveyTextList.split("<>");
			var text = "";
			for (var index=0; index < textSplit.length; index++) {
				// check the type of question
				var type = textSplit[index].split("`");
				var mandatory = textSplit[index].split("#");
				if (type[0] == 1) {
					plottedText = 1;
					inputTextSurveyId = inputTextSurveyId + "textQtnId_"+index+"`"+ "textAnsId_"+index+"~~";					
					text = text + "<div class='single_form_item'>";
					text = text + "<div class='col-md-4 '>";
					var qtn = mandatory[0].split("`");
					if( mandatory[1] == "Y" ){	//	Mandatory input field
						text = text + "<label for='txt' class='col-md-12 control-label align-right mandatory_field'><span id='textQtnId_"+index+"'>"+ qtn[1] +"</span></label>";						
						text = text + "</div>";
						text = text + "<div class='col-md-5'><input type='text' class='form-control-general mandatoryInputField' id='textAnsId_"+index+"' onkeypress='return checkInput(event, this)' /></div>";
					} else {
						text = text + "<label for='txt' class='col-md-12 control-label align-right'><span id='textQtnId_"+index+"'>"+ qtn[1] +"</span></label>";
						text = text + "</div>";
						text = text + "<div class='col-md-5'><input type='text' class='form-control-general' id='textAnsId_"+index+"' onkeypress='return checkInput(event, this)' /></div>";
					}
					text = text + "<div class='col-md-3'>&nbsp;</div><div class='clearfix'></div>";
				} else if (type[0] == 2) {
					var qtn = mandatory[0].split("`");
					plottedRadio = 1;
					var individualSplit = textSplit[index].split("~");				
					var mainQues = individualSplit[0].split("#");	
					var subQtnSplit = individualSplit[1].split("^");
					inputRadioSurveyId = inputRadioSurveyId + "radioQtnId_"+index+"`";
					text = text + "<div class='single_form_item'>";
					text = text + "<div class='col-md-4 '>";
					if(mainQues[1] == "Y"){
						text = text + "<label for='txt' class='col-md-12 control-label align-right mandatory_field'><span id='radioQtnId_"+index+"'>"+ qtn[1] +"</span></label>";					
					} else {
						text = text + "<label for='txt' class='col-md-12 control-label align-right'><span id='radioQtnId_"+index+"'>"+ qtn[1] +"</span></label>";
					}
					text = text + "</div>";
					text = text + "<div class='col-md-6'>";
					inputRadioSurveyId = inputRadioSurveyId + "radioSubQtnId_"+index+"`" + "radio_"+index+"~~";
					for (var subQtnIndex = 0; subQtnIndex < subQtnSplit.length; subQtnIndex++) {
						if(mainQues[1] == "Y"){
							text = text + "<div class='col-md-3'><input type='radio' class='mandatoryInputField' value='"+subQtnSplit[subQtnIndex]+"' name='radio_"+index+"' id='radioAnsId_"+index+"' />&nbsp;<span id='id='radioSubQtnId_"+index+"''>"+subQtnSplit[subQtnIndex]+"</span></div>";
						} else {
							text = text + "<div class='col-md-3'><input type='radio' value='"+subQtnSplit[subQtnIndex]+"' name='radio_"+index+"' id='radioAnsId_"+index+"' />&nbsp;<span id='id='radioSubQtnId_"+index+"''>"+subQtnSplit[subQtnIndex]+"</span></div>";
						}
					}
					text = text + "</div>";	
					text = text + "<div class='col-md-3'>&nbsp;</div><div class='clearfix'></div>";
				} else if (type[0] == 3) {
					var qtn = mandatory[0].split("`");
					plottedDrop = 1;
					var individualSplit = textSplit[index].split("~");
					var mainQues = individualSplit[0].split("#");	
					var subQtnSplit = individualSplit[1].split("^");				
					inputDropSurveyId = inputDropSurveyId + "dropQtnId_"+index+"`";
					text = text + "<div class='single_form_item'>";
					text = text + "<div class='col-md-4 '>";
					if( mainQues[1] =="Y" ){
						text = text + "<label for='txt' class='col-md-12 control-label align-right mandatory_field'><span id='dropQtnId_"+index+"'>"+ qtn[1] +"</span></label>";					
						text = text + "</div>";
						text = text + "<div class='col-md-5'>";
						text = text + "<select class='form-control-general mandatoryInputField' id='drpDwn_"+index+"' type='dropdown'>";
					} else {
						text = text + "<label for='txt' class='col-md-12 control-label align-right'><span id='dropQtnId_"+index+"'>"+ qtn[1] +"</span></label>";
						text = text + "</div>";
						text = text + "<div class='col-md-5'>";
						text = text + "<select class='form-control-general' id='drpDwn_"+index+"' type='dropdown'>";
					}			
					inputDropSurveyId = inputDropSurveyId + "drpDwn_"+index+"~~";
					text = text + "<option class='form-control-general' value='NOTANSWERED'>Select an option</option>";
	         		for (var subQtnIndex = 0; subQtnIndex < subQtnSplit.length; subQtnIndex++) {
						text = text + "<option class='form-control-general' value='"+subQtnSplit[subQtnIndex]+"'>"+subQtnSplit[subQtnIndex]+"</option>";
					}
	         		text = text + "</select>";
					text = text + "</div>";	
					text = text + "<div class='col-md-3'>&nbsp;</div><div class='clearfix'></div>";
				} else if (type[0] == 4) {
					var qtn = mandatory[0].split("`");
					plottedCheck = 1;
					var individualSplit = textSplit[index].split("~");
					var mainQues = individualSplit[0].split("#");
					var subQtnSplit = individualSplit[1].split("^");
					inputCheckSurveyId = inputCheckSurveyId + "checkBoxQtnId_"+index+"`";
					
					text = text + "<div class='single_form_item'>";
					text = text + "<div class='col-md-4 '>";
					if(mainQues[1] == "Y"){
						text = text + "<label for='txt' class='col-md-12 control-label align-right'><span class='mandatory_field' id='checkBoxQtnId_"+index+"'>"+ qtn[1] +"</span></label>";					
					}else {
						text = text + "<label for='txt' class='col-md-12 control-label align-right'><span id='checkBoxQtnId_"+index+"'>"+ qtn[1] +"</span></label>";
					}
					text = text + "</div>";
					text = text + "<div class='col-md-6'>";
					inputCheckSurveyId = inputCheckSurveyId + "checkBoxAnsId_"+index+"`" + "check_"+index+"~~";
					for (var subQtnIndex = 0; subQtnIndex < subQtnSplit.length; subQtnIndex++) {
						if(mainQues[1] == "Y"){
							text = text + "<div class='col-md-3'><input type='checkbox' class='mandatoryInputField' value='"+subQtnSplit[subQtnIndex]+"' name='check_"+index+"' id='checkBoxAnsId_"+index+"' />&nbsp;"+subQtnSplit[subQtnIndex]+"</div>";
						} else {
							text = text + "<div class='col-md-3'><input type='checkbox' value='"+subQtnSplit[subQtnIndex]+"' name='check_"+index+"' id='checkBoxAnsId_"+index+"' />&nbsp;"+subQtnSplit[subQtnIndex]+"</div>";
						}
					}					
					text = text + "</div>";	
					text = text + "<div class='col-md-3'>&nbsp;</div><div class='clearfix'></div>";
				} 
				
			}
			plottingDiv.innerHTML = text;			
		}
	});
}
/**
 * Function saves the user information which the user fills up after resetting his password.
 */
function saveUserDeatils() {
	
	
//	/if($('#iAgree').is(':checked')) {
	
	
	var firstName = $('#inputFN').val();
	var lastName = $('#inputLN').val();
	var location = "";
	var gender = null;
	var supervisePeople = "";
	var tenure = "";
	var functionGrp = "";
	var jobLevel = "";
	var occupation = document.getElementById("ocptnSrchTxtId").value;
	//var nosOfYrs = document.getElementById("nosOfYrs").value;
	var nosOfYrs = document.getElementById("nosOfYrsInput").options[document.getElementById("nosOfYrsInput").selectedIndex].value;
	if(validateSignupForm(occupation, nosOfYrs) == true)	{
		if( validateMandatoryFields() == true){
			if($('#iAgree').is(':checked')) {
				// get the survey question and answer strings
		    	var surveyQtnTextString = getSurveyQuestionTextString();
		    	var surveyQtnRadioString = getSurveyQuestionRadioString();
		    	var surveyQtnCheckBoxString = getSurveyQuestionCheckboxString();
		    	var surveyQtnDropdownString = getSurveyQuestionDropdownString();
		    	if(document.getElementById("male").checked){
					gender = document.getElementById("male").value;
				} else if(document.getElementById("female").checked){
					gender = document.getElementById("female").value;
				} else {
					gender = "U";
				}
				var occupationVal = document.getElementById("ocptnSrchTxtId").value;
				$(".loader_bg").fadeIn();
			    $(".loader").fadeIn();
			    //Create a model
				var userProfile = Spine.Model.sub();
				userProfile.configure("/user/auth/register", "firstNames", "lastNames", "location",  "genders", 
						"supervisePeoples", "tenures", "functionGrps", "jobLevels", "emailIds", "profileId", 
						"surveyQtnTextString", "surveyQtnRadioString", "surveyQtnCheckBoxString", "surveyQtnDropdownString", "occupationVal", "nosOfYrs");
				userProfile.extend( Spine.Model.Ajax );
					
				//Populate the model with data to transfer
				var modelPopulation = new userProfile({  
					firstNames: firstName,
					lastNames: lastName,
					location: location,
					genders: gender,
					supervisePeoples: supervisePeople,
					tenures: tenure,
					functionGrps: functionGrp,
					jobLevels: jobLevel,
					emailIds: sessionStorage.getItem("jctEmail"),
					profileId: sessionStorage.getItem("profileId"),
					surveyQtnTextString: surveyQtnTextString,
					surveyQtnRadioString: surveyQtnRadioString,
					surveyQtnCheckBoxString: surveyQtnCheckBoxString,
					surveyQtnDropdownString: surveyQtnDropdownString, 
					occupationVal: occupationVal,
					nosOfYrs: nosOfYrs
				});
				
				modelPopulation.save(); //POST the data
				
				userProfile.bind("ajaxError", function(record, xhr, settings, error) {
					alertify.alert("Unable to connect to the server.");
				});
				
				userProfile.bind("ajaxSuccess", function(record, xhr, settings, error) {
					var jsonStr = JSON.stringify(xhr);
					var obj = jQuery.parseJSON(jsonStr);
					var statusCode = obj.statusCode;
					/** THIS IS FOR PUBLIC VERSION **/
					if (obj.jctUserId) {
						sessionStorage.setItem("jctUserId", obj.jctUserId);
					}
					if (obj.accountExpWarning) {
						sessionStorage.setItem("accountExp", obj.accountExpWarning);
					}
					if (obj.identifier) {
						sessionStorage.setItem("rowIdentity", obj.identifier);
					}
					/********************************/
					if((statusCode == 200)){
						sessionStorage.setItem("src", "Tool");
						sessionStorage.setItem("name", obj.firstName);
						sessionStorage.setItem("jctEmail", obj.email);
						sessionStorage.setItem("jobReferenceNo", obj.jobRefNo); //Job reference
						sessionStorage.setItem("url", obj.url);
						sessionStorage.setItem("totalTime",obj.totalTime);
						sessionStorage.setItem("rowIdentity", obj.identifier);
						sessionStorage.setItem("jobTitle", obj.jobTitle);
						sessionStorage.setItem("isLogout", "N");
						sessionStorage.setItem("pageSequence", 0);
						sessionStorage.setItem("profileId", obj.profileId);
						sessionStorage.setItem("myPage", "BS");
						sessionStorage.setItem("userId", obj.userId);
						sessionStorage.setItem("url", "/user/view/beforeSketch.jsp");
						window.location = "landing-page.jsp";
					} 
					else if(statusCode == 888){ // mail sent error
						alertify.alert(obj.statusDesc);
						$(".loader_bg").fadeOut();
					    $(".loader").hide();
					} else if (statusCode == "865") { // Onet data not found
						alertify.alert("<img src='../img/server-na-icon.png'><br /><p>"+obj.statusDesc+"</p>");
						$(".loader_bg").fadeOut();
					    $(".loader").hide();
					}
				});
			} else {
				alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please agree to the terms & conditions to proceed.</p>");
				return false;
			}
			return;
		} else {
			return fakse;
		}
	}
	
}


function saveSurveyDeatils() {
	$(".loader_bg").fadeIn();
    $(".loader").fadeIn();
   
    if( validateMandatoryFields() == true){
    	
    	// get the survey question and answer strings
    	var surveyQtnTextString = getSurveyQuestionTextString();
    	var surveyQtnRadioString = getSurveyQuestionRadioString();
    	var surveyQtnCheckBoxString = getSurveyQuestionCheckboxString();
    	var surveyQtnDropdownString = getSurveyQuestionDropdownString();
    	 //Create a model
    	var stub = Spine.Model.sub();
    	stub.configure("/user/auth/saveSurveyQuestions", "userId",  "emailIds",
    			"surveyQtnTextString", "surveyQtnRadioString", "surveyQtnCheckBoxString", "surveyQtnDropdownString");
    	stub.extend( Spine.Model.Ajax );
    		
    	//Populate the model with data to transfer
    	var modelPopulation = new stub({ 
    		userId: sessionStorage.getItem("jctUserId"),
    		emailIds: sessionStorage.getItem("jctEmail"),
    		surveyQtnTextString: surveyQtnTextString,
    		surveyQtnRadioString: surveyQtnRadioString,
    		surveyQtnCheckBoxString: surveyQtnCheckBoxString,
    		surveyQtnDropdownString: surveyQtnDropdownString
    	});
    	modelPopulation.save(); //POST the data
    	
    	stub.bind("ajaxError", function(record, xhr, settings, error) {
    		$(".loader_bg").fadeOut();
    	    $(".loader").hide();
    		alertify.alert("Unable to connect to the server.");
    	});
    	
    	stub.bind("ajaxSuccess", function(record, xhr, settings, error) {
    		sessionStorage.setItem("url", "/user/view/beforeSketch.jsp");
    		window.location = "landing-page.jsp";
    	});
    } else {
    	$(".loader_bg").fadeOut();
	    $(".loader").hide();
	    alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please provide data for all mandatory field.");
    	return false;
    }
}

/**
 * Function checks if mandatory selections [function group and job level] have been selected
 * @param functionGrp
 * @param jobLevel
 * @returns {Boolean}
 */
function validateSignupForm(occupation, nosOfYrs){
	if (occupation.trim().length == 0) {
		document.getElementById("ocptnSrchTxtId").focus();
		alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please select Occupation.");
        return false;
    }
	if (nosOfYrs.trim().length == 0) {
		//document.getElementById("nosOfYrs").focus();
		alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please select how long have you worked in this occupation?");
        return false;
    }
	return true;
}
/**
 * Function toogles between divs.
 * @param showFlag
 */
function toogleDiv(showFlag){
	if (parseInt(showFlag) == 2) {
		document.getElementById('resetPwdId').innerHTML = "Registration Details";
		document.getElementById('passwordResetForm').style.display = "none";
		document.getElementById('loginForm').style.display = "block";
		document.getElementById('userValidatedSection').style.display = "none";
		populateSignupMasterData();
	} else if (parseInt(showFlag) == 1) {
		document.getElementById('resetPwdId').innerHTML = "Reset Password";
		document.getElementById('passwordResetForm').style.display = "block";
		document.getElementById('loginForm').style.display = "none";
		document.getElementById('userValidatedSection').style.display = "none";
	} else if (parseInt(showFlag) == 99) { // no user exists
		document.getElementById('resetPwdId').innerHTML = "";
		document.getElementById('passwordResetForm').style.display = "none";
		document.getElementById('loginForm').style.display = "none";
		document.getElementById('userValidatedSection').style.display = "block";
		document.getElementById('noSuchUserId').innerHTML = "No such user found."
	} else {
		document.getElementById('resetPwdId').innerHTML = "";
		document.getElementById('passwordResetForm').style.display = "none";
		document.getElementById('loginForm').style.display = "none";
		document.getElementById('userValidatedSection').style.display = "block";
	}
	
}
/**
 * function to call forgot password against Enter key.
 */
function searchKeyEvent(e)
{
    // look for window.event in case event isn't passed in
    if (typeof e == 'undefined' && window.event) { e = window.event; }
    if (e.keyCode == 13)
    {
        document.getElementById('resetMyPassword').click();
    }
}

function getSurveyQuestionTextString() {
	inputTextSurveyAns = "";
	// Break the id and corresponding mapping
	var individualMappingString = inputTextSurveyId.split("~~");
	for (var outerIndex = 0; outerIndex < individualMappingString.length-1; outerIndex++) {
		var qtnAnsMapping = individualMappingString[outerIndex].split("`");
		inputTextSurveyAns = inputTextSurveyAns + document.getElementById(qtnAnsMapping[0]).innerHTML;
		inputTextSurveyAns = inputTextSurveyAns + "`";
		var ans = document.getElementById(qtnAnsMapping[1]).value;
		if (ans.trim().length > 0) {
			inputTextSurveyAns = inputTextSurveyAns + "" + ans + "~~";
		} else {
			inputTextSurveyAns = inputTextSurveyAns + "NOTANSWERED~~";
		}
	}
	return inputTextSurveyAns;
}

function getSurveyQuestionRadioString() {
	inputRadioSurveyAns = "";
	// Break the id and corresponding mapping
	var individualMappingString = inputRadioSurveyId.split("~~");
	for (var outerIndex = 0; outerIndex < individualMappingString.length-1; outerIndex++) {
		var qtnAnsMapping = individualMappingString[outerIndex].split("`");
		inputRadioSurveyAns = inputRadioSurveyAns + document.getElementById(qtnAnsMapping[0]).innerHTML;
		inputRadioSurveyAns = inputRadioSurveyAns + "`";
		var checkedItems = document.getElementsByName(qtnAnsMapping[2]);
		for (var checkedIndex = 0; checkedIndex < checkedItems.length; checkedIndex++) {
			if (checkedItems[checkedIndex].checked) {
				inputRadioSurveyAns = inputRadioSurveyAns + "" + checkedItems[checkedIndex].value + "~~";
			}
		}
	}
	return inputRadioSurveyAns;
}

function getSurveyQuestionCheckboxString() {
	inputCheckSurveyAns = "";
	// Break the id and corresponding mapping
	var individualMappingString = inputCheckSurveyId.split("~~");
	for (var outerIndex = 0; outerIndex < individualMappingString.length-1; outerIndex++) {
		var qtnAnsMapping = individualMappingString[outerIndex].split("`");
		inputCheckSurveyAns = inputCheckSurveyAns + document.getElementById(qtnAnsMapping[0]).innerHTML;
		inputCheckSurveyAns = inputCheckSurveyAns + "`";
		var checkedItems = document.getElementsByName(qtnAnsMapping[2]);
		for (var checkedIndex = 0; checkedIndex < checkedItems.length; checkedIndex++) {
			if (checkedItems[checkedIndex].checked) {
				inputCheckSurveyAns = inputCheckSurveyAns + "" + checkedItems[checkedIndex].value + "#";
			}
		}
		inputCheckSurveyAns = inputCheckSurveyAns + "~~"; 
	}
	return inputCheckSurveyAns;
}

function getSurveyQuestionDropdownString() {
	inputDropSurveyAns = "";
	// Break the id and corresponding mapping
	var individualMappingString = inputDropSurveyId.split("~~");
	for (var outerIndex = 0; outerIndex < individualMappingString.length-1; outerIndex++) {
		var qtnAnsMapping = individualMappingString[outerIndex].split("`");
		inputDropSurveyAns = inputDropSurveyAns + document.getElementById(qtnAnsMapping[0]).innerHTML;
		inputDropSurveyAns = inputDropSurveyAns + "`";
		var drp = document.getElementById(qtnAnsMapping[1]).options[document.getElementById(qtnAnsMapping[1]).selectedIndex].value;
		inputDropSurveyAns = inputDropSurveyAns + "" + drp + "~~";
	}
	return inputDropSurveyAns;
}

function checkInput(event, obj) {
	if(event.keyCode == 96 || event.keyCode == 126 || event.keyCode == 35) {
		return false;
	}  else {
		return true;
	}
}

function searchOccupationSearch() {
	var occupationSearchVal = document.getElementById("ocptnSrchTxtId").value.trim();
	occupationSearchVal = occupationSearchVal.replace(/,+/g, ' ');
	var val = occupationSearchVal.replace(/ +/g, ',');
	if (val.indexOf(",") == 0) {
		val = val.substring(1, val.length);
	}
	val = val.replace(/([*()\+\\{}\[\]\|\/?])+/g, '');
	if (val == "") {
		//alertify.alert("Only special character is not valid.");
		var plottingDiv = document.getElementById("dataPlottingDiv");
		plottingDiv.innerHTML = "<div align='center'><br /><br /><img src='../img/not-found.png'><br /><p>The occupation you have entered is not in the system.<br />Please enter search words and select an occupation from the list provided.</p></div>";
		return false;
	}
	
	
	if (document.getElementById("dataPlottingDiv")) {
		document.getElementById("dataPlottingDiv").innerHTML = "";
	}
	var occupationSearchVal = document.getElementById("ocptnSrchTxtId").value.trim();
	if (occupationSearchVal.trim().length == 0) {		
		alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please provide search words that describe your occupation.");
		//$('#searchOccupation').off('click');
		$("#searchOccupation").click(function() {$.fancybox.close({});});
		return false;
	} else {
		document.getElementById("goDiv").style.display = "none";
		document.getElementById("dataPlottingDiv").innerHTML = "";
		document.getElementById("occupationPlottingDiv").innerHTML = "<br /><br /><img src='../img/Processing.gif' />";			
		$("#searchOccupation").click(function() {
			$.fancybox.open({
				href : '#occupationDiv',
				closeClick : false,
				openEffect : 'elastic',
				openSpeed  : 150,
				closeEffect: 'elastic',
				closeSpeed : 150,
				padding    : 5,
				closeClick : true,
				modal      : true,
				showCloseButton : true,
				afterShow : function() {
					$('.fancybox-skin').append('<a title="Close" class="fancybox-item fancybox-close" href="javascript:jQuery.fancybox.close();"></a>');
				}
			});
		});
	}
	// Search the o* net list
	var stub = Spine.Model.sub();
	stub.configure("/user/auth/searchOccupationList", "searchString");
	stub.extend( Spine.Model.Ajax );
	/*occupationSearchVal = occupationSearchVal.replace(/,+/g, ' ');
	var val = occupationSearchVal.replace(/ +/g, ',');
	if (val.indexOf(",") == 0) {
		val = val.substring(1, val.length);
	}*/
	//Populate the model with data to transfer
	var modelPopulation = new stub({ 
		searchString: val
	});
	modelPopulation.save(); //POST the data
	
	stub.bind("ajaxError", function(record, xhr, settings, error) {
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
		alertify.alert("Unable to connect to the server.");
	});
	
	stub.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if (statusCode == "200") {
			plotData(obj.dtoList);
		} else {
			document.getElementById("occupationPlottingDiv").innerHTML = "";
			var plottingDiv = document.getElementById("dataPlottingDiv");
			plottingDiv.innerHTML = "<div align='center'><br /><br /><img src='../img/not-found.png'><br /><p>"+obj.statusMsg+"</p></div>";
		}
	});	
}

function plotData(dtoList) {
	document.getElementById("occupationPlottingDiv").innerHTML = "";
	var plottingDiv = document.getElementById("dataPlottingDiv");
	plottingDiv.innerHTML = "";
	var str = "<table width='65%' class='custom_table_onet'><thead><tr class='tab_header'></tr></thead><tbody>";
	
	for (var index = 0; index < dtoList.length; index++) {
		var trColor = "";
		 var titleId = "title_" + index;
		 var descId = "desc_" + index;
		 var showHideId = "show_hide_" + index;
		str = str + "<tr id='"+titleId+"' bgcolor="+trColor+"><td class='custom_td_onet' ><input type='radio' class='custom_radio_onet' name='occupation' value='"+dtoList[index].title+"'>"+dtoList[index].title+"</td><td id='"+showHideId+"' class='hide_onet' onclick='showDescription(\""+descId+"\", \""+dtoList.length+"\", \""+showHideId+"\")'></td></tr><tr id= '"+descId+"' style='display:none;'><td class='onet_desc_custom_cls' >"+dtoList[index].desc+"</td></tr>";
	}
	str = str + "</tbody></table>";
	plottingDiv.innerHTML = str;
	document.getElementById("goDiv").style.display = "block";
}

function closeDiv() {
	if (undefined != $("input[name=occupation]:checked").val()) {
		document.getElementById("ocptnSrchTxtId").value = $("input[name=occupation]:checked").val();
		jQuery.fancybox.close();
	} else {
		alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please select an occupation from the list</p>");
		return false;
	}
}

/**
 * function to validate all mandatory input field
 * @return {boolean}
 * 
 * */
function validateMandatoryFields(){
	
	var flag = 0;
	$(":input.mandatoryInputField").each(function(){		
		var fieldId = $(this).attr('id');
		var fieldType = $(this).attr('type');
		if( fieldType == 'text' && $("#"+fieldId).val().trim() == ""){
			flag = 1;
			document.getElementById(fieldId).focus();
			alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please answer mandatory text question</p>");
			return false;		    				
		}		  
		if( fieldType == 'radio' && ($("input[id='"+fieldId+"']:checked").length == 0) ){
			flag = 1;
			alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please answer mandatory radio question</p>");
			return false;		    				
		}		
		if( fieldType == 'checkbox' && ($("input[id='"+fieldId+"']:checked").length == 0) ){
			flag = 1;
			alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please answer mandatory checkbox question</p>");
			return false;		    				
		}		 	        
		if( fieldType == 'dropdown') {
			var ddvalue = document.getElementById(fieldId).options[document.getElementById(fieldId).selectedIndex].value;		
			if(ddvalue == '' || ddvalue == 'NOTANSWERED') {				
				flag = 1;
				alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please answer mandatory dropdown question</p>");
				return false;
			}
		 }
	});	
	return (flag ==1 ? false : true);	
}

function validateOnetData(){
    /*var textInput = document.getElementById("ocptnSrchTxtId").value;
    //textInput = textInput.replace(/[^A-Za-z]/g, "");
    textInput = textInput.replace(/%/g, "");
    document.getElementById("ocptnSrchTxtId").value = textInput;*/
}

function showDescription(descId, length, showHideId) {
	if($('#'+showHideId).attr('class') == "hide_onet") {
		$("#" + showHideId).removeClass("hide_onet");
		$("#" + showHideId).addClass("show_onet");
		document.getElementById(descId).style.display = "block";
	} else {
		$("#" + showHideId).removeClass("show_onet");
		$("#" + showHideId).addClass("hide_onet");
		document.getElementById(descId).style.display = "none";
	}
	
	
}

function fetchTnstructions(){
	var userProfile = parseInt(sessionStorage.getItem("profileId"));
	var tnC = Spine.Model.sub();
	tnC.configure("/user/auth/fetchTnC", "userProfile", "userType");
	tnC.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new tnC({  
		userProfile: userProfile,
		userType: 1
	});
	modelPopulator.save(); //POST
	tnC.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	tnC.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		console.log(obj.jctTcDescription);
		if(statusCode == 200) {
			$('#myModal').modal('show');
			document.getElementById("terms_condtn_div_id_survey").innerHTML = obj.jctTcDescription;			
		} else if (statusCode = 100){
			alertify.alert("No terms & conditions found for your profile.");
			return false;
		} else {
			//Show error message
			alertify.alert("Some thing went wrong while fetching Terms & Conditions.");
		}
	});
	
	
	
}
