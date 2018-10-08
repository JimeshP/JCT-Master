var mainQtnMemory = "";
var radioSameValidationMemory = "";
var dropSameValidationMemory = "";
var checkSameValidationMemory = "";
var qtnTypeStore = "";

$(document).ready(function() {
	 document.getElementById("userType").removeAttribute("disabled", "true");
	 document.getElementById("userType").selectedIndex=""; 
	 document.getElementById("chgOrderSurvQtnId").disabled = true;
	sessionStorage.removeItem("previous_text_value"); 
	sessionStorage.removeItem("previous_option_value"); 
	populateProfileDropDown();
	fetchExistingSurveyQuestions();
});

/**
 * Function fetches user profile dropdown data
 * @param null
 */
function populateProfileDropDown(){
	var userGrp = Spine.Model.sub();
	userGrp.configure("/admin/manageuser/populateUserProfile", "none");
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
			populateDropDown(obj.userProfileMap);
		}  else {
			//Show error message
			alertify.alert("Some thing went wrong.");
		}
	});
}

/**
 * Function to populate the dropdown for user group
 */
function populateDropDown(userProfileMap) {
	var userProfilSelect = document.getElementById("inputUserProfileInpt");
	for (var key in userProfileMap) {				
		var option = document.createElement("option");
		option.text = userProfileMap[key];
	    option.value = key+"!"+userProfileMap[key];
	    option.className = "form-control-general";
	    try {
	    	userProfilSelect.add(option, null); //Standard 
	    }catch(error) {
	    	//regionSelect.add(option); // IE only
	    }
	}
}

function toogleSelectionDiv(selection) {
  if (selection == "R" || selection == "D" || selection == "C") {
	  document.getElementById("radiopt").style.display = "block";
	  document.getElementById("textopt").style.display = "none";
	  document.getElementById("quesEntry").style.display = "none";
	  document.getElementById("radioEntry").style.display = "block";
	  document.getElementById("radioEntry").innerHTML = "";
	  document.getElementById("radioqtn").style.display = "block";
	  document.getElementById("questions").value = "";
	  document.getElementById("inputNoOFUser").selectedIndex=""; 
	  document.getElementById("inputNoOFQues").selectedIndex="";
	  document.getElementById("quesEntry").innerHTML = "";
	  if( $("#mandatoryChkbox").is(":checked")) {
		  $('#mandatoryChkbox').attr('checked', false);
	  }
	  //$('input:checkbox[id=mandatoryChkbox]').attr('checked',false);
  } else {
	  document.getElementById("radiopt").style.display = "none";
	   document.getElementById("textopt").style.display = "block";
	   document.getElementById("quesEntry").style.display = "block";
	   document.getElementById("radioEntry").style.display = "none";
	   document.getElementById("radioqtn").style.display = "none";
	   document.getElementById('inputNoOFQues').disabled=false;
  }
  var tableDiv = document.getElementById("existingSurveyQtnsTableId");
  tableDiv.innerHTML = "<div align='center'><img src='../img/dataLoading.GIF' /><p> Loading Existing Survey Questions</p></div>";
  if (selection == "T") {
	  fetchAllSurveyQuestionsByUserTypeAndQtnType(1);
  } else if (selection == "R") {
	  fetchAllSurveyQuestionsByUserTypeAndQtnType(2);
  } else if (selection == "D") {
	  fetchAllSurveyQuestionsByUserTypeAndQtnType(3);
  } else {
	  fetchAllSurveyQuestionsByUserTypeAndQtnType(4);
  }
 }

function populateFields(){		
	/**** JCTP-9 DATE 04/12/2014****/
	divCount = $('#quesEntry').children('.single_form_item:visible').length;
	var jsonObj = [];
	var counter = 0;
	for (var i = 1; i <= divCount; i += 1) {
	var unitJ = {};	
	var value = document.getElementById('manualTextQues_'+i).value;
	unitJ["ElementValue"+i] = value;
	jsonObj[counter++] = unitJ;
	}
	sessionStorage.setItem("previous_text_value",JSON.stringify(jsonObj)); 
	/**** END ****/
	
	var quesNumber = document.getElementById("inputNoOFQues").options[document.getElementById("inputNoOFQues").selectedIndex].value;
	var plotPlement = document.getElementById("quesEntry");
	document.getElementById("quesEntry").innerHTML = "";
	document.getElementById("quesEntry").style.display = "block";
	var string = "";
	var idCountIndex = 1;
	
	if (quesNumber == "") {
		/* alertity.alert("Please select numbers.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide(); */
	} else {
		if (document.getElementById("saveConfig")) {
			document.getElementById("saveConfig").style.display = "block";
		}
		for (var index = 0; index < quesNumber; index++) {
			var textQuesId = "manualTextQues_"+idCountIndex;
			var mandatoryChkBoxID =	"mandatoryChkBox_"+idCountIndex;
			
			string = string + "<div class='single_form_item'><div class='col-sm-4 '><label for='inputQtnNo' class='col-sm-12 control-label text_label' id='qtn_div_id"+idCountIndex+"'>Question "+idCountIndex+":</label></div>" +
					"<div class='col-sm-4'><input type='text' class='form-control-general-form-redefined surveyText' value='' id='"+textQuesId+"' maxLength='100' onkeypress='return validateDescription(event)'></div>" +
					"<div class='col-md-2 mandatoryCheckBox'><label for='inputFN' class='text_label' style='margin-left: -13%;'>Mandatory:</label>" +
					"<input type='checkbox'  id='"+mandatoryChkBoxID+"' class='mandatoryChkBox' name='mandatory' value='Y'></div><div class='clearfix'></div></div>";							
			idCountIndex = idCountIndex + 1;
		}	
		plotPlement.innerHTML = string;
	}
	
	/**** JCTP-9 DATE 04/12/2014****/
	if(sessionStorage.getItem("previous_text_value") != null){
		var jsonObj = JSON.parse(sessionStorage.getItem("previous_text_value"));
		var count = jsonObj.length;
			for (var j = 0; j<count; j++){	
			var value = jsonObj[j]["ElementValue"+(j+1)];
			var div = document.getElementById('manualTextQues_'+(j+1));
			    if (div) {
				document.getElementById('manualTextQues_'+(j+1)).value = value;
				}			
			}		
		}
	 /**** END ****/
}

function populateFieldsRadio(){
	
	/**** JCTP-9 DATE 04/12/2014****/
	divCount = $('#radioEntry').children('.single_form_item:visible').length;
	var jsonObj = [];
	var counter = 0;
	for (var i = 1; i <= divCount; i += 1) {
	var unitJ = {};	
	var value = document.getElementById('manualRadioQues_'+i).value;
	unitJ["ElementValue"+i] = value;
	jsonObj[counter++] = unitJ;
	}
	sessionStorage.setItem("previous_option_value",JSON.stringify(jsonObj)); 
	/**** END ****/
	
	var manualEntryNumber = document.getElementById("inputNoOFUser").selectedIndex;
	var plotPlement = document.getElementById("radioEntry");
	document.getElementById("radioEntry").innerHTML = "";
	document.getElementById("dropdwnEntry").innerHTML= "";
	document.getElementById("checkboxEntry").innerHTML= "";
	var string = "";
	var idCountIndex = 1;
	if (manualEntryNumber == "") {
		/* alertity.alert("Please select numbers.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide(); */
	} else {
		if (document.getElementById("saveConfig")) {
			document.getElementById("saveConfig").style.display = "block";
		}
		if (document.getElementById("quesEntry")) {
			document.getElementById("quesEntry").style.display = "block";
		}
		if (document.getElementById("radioEntry")) {
			document.getElementById("radioEntry").style.display = "block";
		}
		if (document.getElementById("dropdwnEntry")) {
			document.getElementById("dropdwnEntry").style.display = "block";
		}
		if (document.getElementById("checkboxEntry")) {
			document.getElementById("checkboxEntry").style.display = "block";
		}
		for (var index = 0; index < manualEntryNumber; index++) {
			var radioQuesId = "manualRadioQues_"+idCountIndex;
			string = string + "<div class='single_form_item'><div class='col-sm-4 '><label for='inputQtnNo' class='col-sm-12 control-label text_label' id='qtn_div_id"+idCountIndex+"'>Option "+idCountIndex+":</label></div>" +
					"<div class='col-sm-4'><input type='text' class='form-control-general-form-redefined survey_radio_width' style='' value='' id='"+radioQuesId+"' maxLength='100' onkeypress='return validateDescription(event)'></div>" +
					"<div class='col-md-2'>&nbsp;</div><div class='clearfix'></div></div>";							
			idCountIndex = idCountIndex + 1;
		}
		plotPlement.innerHTML = string;
	}	
	/**** JCTP-9 DATE 04/12/2014****/
	if(sessionStorage.getItem("previous_option_value") != null){
		var jsonObj = JSON.parse(sessionStorage.getItem("previous_option_value"));
		var count = jsonObj.length;
			for (var j = 0; j<count; j++){	
			var value = jsonObj[j]["ElementValue"+(j+1)];
			var div = document.getElementById('manualRadioQues_'+(j+1));
			    if (div) {
				document.getElementById('manualRadioQues_'+(j+1)).value = value;
				}			
			}		
		}
	 /**** END ****/
} 

function fieldsDisabled() {
	var radiobut = document.getElementsByName('radios');
	var e = document.getElementById("userType");
	var value = e.options[e.selectedIndex].value;
	if(value == '1' || value == '3'){
		document.getElementById('inputUserProfileInpt').disabled=false;
		document.getElementById('inputNoOFQues').disabled=false;
		document.getElementById('inputNoOFUser').disabled=false;
		document.getElementById('questions').readOnly=false;
		document.getElementById("quesEntry").innerHTML = "";
		document.getElementById("radioEntry").innerHTML = "";
		document.getElementById("inputNoOFUser").selectedIndex=""; 
		document.getElementById("inputNoOFQues").selectedIndex="";
		document.getElementById("questions").value= "";
		/*for (var i = 0; i < radiobut.length; i++) {
		    radiobut[i].disabled = false;
		}*/
	}
	else {
		document.getElementById('inputUserProfileInpt').disabled=true;
		document.getElementById('inputNoOFQues').disabled=true;
		document.getElementById('inputNoOFUser').disabled=true;
		document.getElementById('questions').readOnly=true;
		document.getElementById('radioEntry').style.display='none';
		document.getElementById("inputUserProfileInpt").selectedIndex=""; 
		document.getElementById("inputNoOFUser").selectedIndex=""; 
		document.getElementById("inputNoOFQues").selectedIndex="";  
		document.getElementById("quesEntry").innerHTML = "";
		/*for (var i = 0; i < radiobut.length; i++) {
		    radiobut[i].disabled = true;
		}*/
	}
	var tableDiv = document.getElementById("existingSurveyQtnsTableId");
	tableDiv.innerHTML = "<div align='center'><img src='../img/dataLoading.GIF' /><p> Loading Existing Survey Questions</p></div>";
	//Check if the radio buttons are also selected
	if (document.getElementById("text").checked) {
		fetchAllSurveyQuestionsByUserTypeAndQtnType(1);
	} else if (document.getElementById("rad").checked) {
		fetchAllSurveyQuestionsByUserTypeAndQtnType(2);
	} else if (document.getElementById("drop").checked) {
		fetchAllSurveyQuestionsByUserTypeAndQtnType(3);
	} else if (document.getElementById("check").checked) {
		fetchAllSurveyQuestionsByUserTypeAndQtnType(4);
	} else {		
		fetchAllSurveyQuestionsByUserType(parseInt(value));
	}
} 


function loadProfileData() {
	var radiobut = document.getElementsByName('radios');
	for (var i = 0; i < radiobut.length; i++) {
	    radiobut[i].disabled = false;
	}
	qtnTypeStore = "";

	var e = document.getElementById("userType");
	var value = e.options[e.selectedIndex].value;
	var userType = parseInt(value);
	var userProfile = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value;
	
	if(null == userProfile || userProfile == 0) {
		document.getElementById("chgOrderSurvQtnId").disabled = true;
	} else {
		document.getElementById("chgOrderSurvQtnId").disabled = false;
	}
	
	if (document.getElementById("text").checked) {
		qtnTypeStore = 1;
	} else if (document.getElementById("rad").checked) {
		 qtnTypeStore = 2;
	} else if (document.getElementById("drop").checked) {
		 qtnTypeStore = 3;
	} else if (document.getElementById("check").checked) {
		qtnTypeStore = 4;
	} else {
		qtnTypeStore = 0;
	}
	
	
	// if user type and user profile are available
	if ((userType > 0) && (userProfile != "0") && (qtnTypeStore > 0)) {
		var userGrp = Spine.Model.sub();
		var profile = userProfile.split("!");
		var profileId = parseInt(profile[0]);
		userGrp.configure("/admin/surveyQtns/fetchAllExistingSurveyQtnsAllSelected", "userType", "qtnType", "profile");
		userGrp.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new userGrp({  
			userType: userType,
			qtnType: qtnTypeStore,
			profile: profileId
		});
		modelPopulator.save(); //POST
		userGrp.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
		});
		userGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			populateTable (obj.resultList, obj.resultList.length);
		});
	} else if ((userType == 0) && (userProfile != "0") && (qtnTypeStore > 0)) {
		var userGrp = Spine.Model.sub();
		var profile = userProfile.split("!");
		var profileId = parseInt(profile[0]);
		userGrp.configure("/admin/surveyQtns/fetchAllExistingSurveyQtnsOnlyUserProfile", "profile", "qtnType");
		userGrp.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new userGrp({  
			userType: userType,
			qtnType: qtnTypeStore,
			profile: profileId
		});
		modelPopulator.save(); //POST
		userGrp.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
		});
		userGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			populateTable (obj.resultList, obj.resultList.length);
		});
	} else if ((userType > 0) && (userProfile == "0") && (qtnTypeStore > 0)) {
		var userGrp = Spine.Model.sub();
		userGrp.configure("/admin/surveyQtns/fetchAllExistingSurveyQtnsByUserTypeAndQtnType", "userType", "qtnType");
		userGrp.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new userGrp({  
			userType: userType,
			qtnType: qtnTypeStore
		});
		modelPopulator.save(); //POST
		userGrp.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
		});
		userGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			populateTable (obj.resultList, obj.resultList.length);
		});
	} else if ((qtnTypeStore == 0)){
		var profile = userProfile.split("!");
		var profileId = parseInt(profile[0]);
		var userGrp = Spine.Model.sub();
		userGrp.configure("/admin/surveyQtns/fetchSurveyQtnsByProfileAndUserType", "profile", "userType");
		userGrp.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new userGrp({  
			profile: profileId,
			userType: userType
		});
		modelPopulator.save(); //POST
		userGrp.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
		});
		userGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			populateTable (obj.resultList, obj.resultList.length);
		});
	} else {
		var userGrp = Spine.Model.sub();
		userGrp.configure("/admin/surveyQtns/fetchAllExistingSurveyQtnsByQtnType", "qtnType");
		userGrp.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new userGrp({  
			qtnType: qtnTypeStore
		});
		modelPopulator.save(); //POST
		userGrp.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
		});
		userGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			populateTable (obj.resultList, obj.resultList.length);
		});
	
	}
}

/**
 * Function saves the survey questions
 */
function saveSurveyQtns() {
	// Check if user type selected
	var userType = parseInt(document.getElementById("userType").options[document.getElementById("userType").selectedIndex].value);
	var userProfile = parseInt(document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value);
	if(userType == "0") {
		alertify.alert("Please Select User Type.");
		return false;
	} else if(userProfile == "0") {
		alertify.alert("Please Select User Profile.");
		return false;
	} else {
		if (document.getElementById("text").checked) {
			if (validateFormSubmission("T")) {
				saveFreeTextData();
			}
		} else if (document.getElementById("rad").checked) {
			if (validateFormSubmission("R")) {
				saveRadioData();
			}
		} else if (document.getElementById("drop").checked) {
			if (validateFormSubmission("D")) {
				saveDropDownData();
			}
		} else if (document.getElementById("check").checked) {
			if (validateFormSubmission("C")) {
				saveCheckboxData();
			}
		} else {
			alertify.alert("Please Choose any one Answer type");
			return false;
		}
	}
}

function validateFormSubmission (type) {
	if (type == "T") {
		var quesNumber = document.getElementById("inputNoOFQues").options[document.getElementById("inputNoOFQues").selectedIndex].value;
		if (quesNumber == "") {
			alertify.alert("Please select number of questions.");
			return false;
		} else {
			for (var index = 1; index<=quesNumber; index++) {
				var inputTextVal = document.getElementById("manualTextQues_"+index).value;
				if (inputTextVal.trim().length == 0) {
					//alertify.alert("Please enter a question value.");
					alertify.alert("Please enter question "+index);
					return false;
				} else if( !specialCharValidation(inputTextVal) ){
					return false;			
				}
			}
		}
		
		return true;
	} else if (type == "R") {
		var questions = document.getElementById("questions").value;
		if (questions.trim().length == 0) {
			alertify.alert("Please enter Main Question.");
			return false;
		} else if( !specialCharValidation(questions) ){
			return false;			
		}
		var nosOfOptions = document.getElementById("inputNoOFUser").options[document.getElementById("inputNoOFUser").selectedIndex].value;
		if (nosOfOptions == "") {
			alertify.alert("Please select number of options.");
			return false;
		} else {
			for (var index = 1; index<=nosOfOptions; index++) {
				var inputTextVal = document.getElementById("manualRadioQues_"+index).value;
				if (inputTextVal.trim().length == 0) {
					alertify.alert("Please enter option "+index);
					return false;
				} else if( !specialCharValidation(inputTextVal) ){
					return false;			
				}
			}
		}
		return true;
	} else if (type == "D") {
		var questions = document.getElementById("questions").value;
		if (questions.trim().length == 0) {
			alertify.alert("Please enter Main Question.");
			return false;
		} else if( !specialCharValidation(questions) ){
			return false;			
		}
		var nosOfOptions = document.getElementById("inputNoOFUser").options[document.getElementById("inputNoOFUser").selectedIndex].value;
		if (nosOfOptions == "") {
			alertify.alert("Please select number of options.");
			return false;
		} else {
			for (var index = 1; index<=nosOfOptions; index++) {
				var inputTextVal = document.getElementById("manualRadioQues_"+index).value;
				if (inputTextVal.trim().length == 0) {
					alertify.alert("Please enter option "+index);
					return false;
				} else if( !specialCharValidation(inputTextVal) ){
					return false;			
				}
			}
		}
		return true;
	} else if (type == "C") {
		var questions = document.getElementById("questions").value;
		if (questions.trim().length == 0) {
			alertify.alert("Please enter Main Question.");
			return false;
		} else if( !specialCharValidation(questions) ){
			return false;			
		}
		var nosOfOptions = document.getElementById("inputNoOFUser").options[document.getElementById("inputNoOFUser").selectedIndex].value;
		if (nosOfOptions == "") {
			alertify.alert("Please select number of options.");
			return false;
		} else {
			for (var index = 1; index<=nosOfOptions; index++) {
				var inputTextVal = document.getElementById("manualRadioQues_"+index).value;
				if (inputTextVal.trim().length == 0) {
					alertify.alert("Please enter option "+index);
					return false;
				} else if( !specialCharValidation(inputTextVal) ){
					return false;			
				}
			}
		}
		return true;
	}
}

function saveFreeTextData() {
	$(".loader_bg").fadeIn();
 	$(".loader").fadeIn();
	var textQtns = "";
	var userType = parseInt(document.getElementById("userType").options[document.getElementById("userType").selectedIndex].value);
	var userProfile = parseInt(document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value);
	var userProfileText = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].text;
	userProfile = userProfile +"!"+ userProfileText;
	
	var nosOfTextQuestions = parseInt(document.getElementById("inputNoOFQues").options[document.getElementById("inputNoOFQues").selectedIndex].value);
	for (var index = 1; index<=nosOfTextQuestions; index++) {		
		var isMandatory = $("#"+"mandatoryChkBox_"+index).is(":checked") ? "Y" : "N";
		textQtns = textQtns + document.getElementById("manualTextQues_"+index).value+"~"+isMandatory+"##";
	}	
	var userGrp = Spine.Model.sub();
	userGrp.configure("/admin/surveyQtns/saveFreeText", "userType", "textQtns", "createdBy", "userProfile");
	userGrp.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userGrp({  
		userType: userType,
		textQtns: textQtns,
		createdBy: sessionStorage.getItem("jctEmail"),
		userProfile: userProfile
	});
	modelPopulator.save(); //POST
	userGrp.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	});
	userGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		alertify.alert(obj.message);
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	    document.getElementById("saveConfig").style.display = "none";
		document.getElementById("inputNoOFQues").selectedIndex = "0";
		$("#quesEntry").html("");
				
		if (obj.resultList) {
			populateTable (obj.resultList);
		}
	});
}

function saveRadioData() {
	$(".loader_bg").fadeIn();
 	$(".loader").fadeIn();
	var subQtns = "";
	var userType = parseInt(document.getElementById("userType").options[document.getElementById("userType").selectedIndex].value);
	var userProfile = parseInt(document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value);
	var userProfileText = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].text;
	userProfile = userProfile +"!"+ userProfileText;
	var mainQuestion = document.getElementById("questions").value;
	var nosOfSubQuestions = parseInt(document.getElementById("inputNoOFUser").options[document.getElementById("inputNoOFUser").selectedIndex].value);
	var isMandatory = $("#mandatoryChkbox").is(":checked") ? "Y" : "N";
	for (var index = 1; index<=nosOfSubQuestions; index++) {
		subQtns = subQtns + document.getElementById("manualRadioQues_"+index).value+"`~!~`";
	}
	var userGrp = Spine.Model.sub();
	userGrp.configure("/admin/surveyQtns/saveRadio", "userType", "mainQuestion", "subQtns", "createdBy", "mandatoryChkVal", "userProfile");
	userGrp.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userGrp({  
		userType: userType,
		mainQuestion: mainQuestion,
		subQtns:subQtns,
		createdBy: sessionStorage.getItem("jctEmail"),
		mandatoryChkVal: isMandatory,
		userProfile: userProfile
	});
	modelPopulator.save(); //POST
	userGrp.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	});
	userGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		alertify.alert(obj.message);
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
		document.getElementById("questions").value = "";	
		document.getElementById("inputNoOFUser").selectedIndex = "0";
		if( $("#mandatoryChkbox").is(":checked")) {
			$('#mandatoryChkbox').attr('checked', false);
	 	}
		$("#radioEntry").html("");
		if (obj.resultList) {
			populateTable (obj.resultList);
		}
		
	});
}

function saveDropDownData() {
	$(".loader_bg").fadeIn();
 	$(".loader").fadeIn();
	var subQtns = "";
	var userType = parseInt(document.getElementById("userType").options[document.getElementById("userType").selectedIndex].value);
	var userProfile = parseInt(document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value);
	var userProfileText = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].text;
	userProfile = userProfile +"!"+ userProfileText;
	var mainQuestion = document.getElementById("questions").value;
	var nosOfSubQuestions = parseInt(document.getElementById("inputNoOFUser").options[document.getElementById("inputNoOFUser").selectedIndex].value);
	var mandatoryChkVal = $("#mandatoryChkbox").is(":checked") ? "Y" : "N";
	for (var index = 1; index<=nosOfSubQuestions; index++) {
		subQtns = subQtns + document.getElementById("manualRadioQues_"+index).value+"`~!~`";
	}

	var userGrp = Spine.Model.sub();
	userGrp.configure("/admin/surveyQtns/saveDropDown", "userType", "mainQuestion", "subQtns", "createdBy", "mandatoryChkVal", "userProfile");
	userGrp.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userGrp({  
		userType: userType,
		mainQuestion: mainQuestion,
		subQtns:subQtns,
		createdBy: sessionStorage.getItem("jctEmail"),
		mandatoryChkVal: mandatoryChkVal,
		userProfile: userProfile
		
	});
	modelPopulator.save(); //POST
	userGrp.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	});
	userGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		alertify.alert(obj.message);
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
		document.getElementById("questions").value = "";
		$("#radioEntry").html("");
		document.getElementById("inputNoOFUser").selectedIndex = "0";
		if( $("#mandatoryChkbox").is(":checked")) {
			$('#mandatoryChkbox').attr('checked', false);
		}
		if (obj.resultList) {
			populateTable (obj.resultList);
		}
	});
}

function saveCheckboxData() {
	$(".loader_bg").fadeIn();
 	$(".loader").fadeIn();
	var subQtns = "";
	var userType = parseInt(document.getElementById("userType").options[document.getElementById("userType").selectedIndex].value);
	var userProfile = parseInt(document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value);
	var userProfileText = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].text;
	userProfile = userProfile +"!"+ userProfileText;
	var mainQuestion = document.getElementById("questions").value;
	var nosOfSubQuestions = parseInt(document.getElementById("inputNoOFUser").options[document.getElementById("inputNoOFUser").selectedIndex].value);
	var mandatoryChkVal = $("#mandatoryChkbox").is(":checked") ? "Y" : "N";
	for (var index = 1; index<=nosOfSubQuestions; index++) {
		subQtns = subQtns + document.getElementById("manualRadioQues_"+index).value+"`~!~`";
	}

	var userGrp = Spine.Model.sub();
	userGrp.configure("/admin/surveyQtns/saveCheckBox", "userType", "mainQuestion", "subQtns", "createdBy", "mandatoryChkVal", "userProfile");
	userGrp.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userGrp({  
		userType: userType,
		mainQuestion: mainQuestion,
		subQtns:subQtns,
		createdBy: sessionStorage.getItem("jctEmail"),
		mandatoryChkVal: mandatoryChkVal,
		userProfile: userProfile
		
	});
	modelPopulator.save(); //POST
	userGrp.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	});
	userGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		alertify.alert(obj.message);
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
		document.getElementById("questions").value = "";
		document.getElementById("inputNoOFUser").selectedIndex = "0";
		$("#radioEntry").html("");
		if( $("#mandatoryChkbox").is(":checked")) {
			$('#mandatoryChkbox').attr('checked', false);
		} 
		
		if (obj.resultList) {
			populateTable (obj.resultList);
		}
	});
}

function populateTable (resultList, size) {
	
	$(".loader_bg").fadeOut();
    $(".loader").hide();
	if (size == 0) {
		plotNoData();
	} else {
		document.getElementById("existing_list_Id").style.display="block";
		var tableDiv = document.getElementById("existingSurveyQtnsTableId");
		tableDiv.innerHTML = "";
		var tableStr = "<table width='94%' border='1' bordercolor='#78C0D3' id='myTable' align='center' class='tablesorter'><thead class='tab_header'><tr><th width='9%'>SL. No.</th><th width='12%'>User Type</th><th width='12%'>User Profile</th><th width='12%'>Answer Type</th><th width=''>Question Description</th><th width='8%'>Mandatory</th><th width='20%'>Options</th><th width='12%'>Action</th></tr></thead><tbody>";
		var rowSplit = resultList.split("~");
		var counter = 1;
		for (var rowIndex=0; rowIndex<rowSplit.length-1; rowIndex++) {
			var individualRow = rowSplit[rowIndex].split("`");
			var trColor = "";
			if(rowIndex % 2 == 0) {
				trColor = "#D2EAF0";
			} else {
				trColor = "#F1F1F1";
			}
			var subQtns = individualRow[3].split("^");			
			var subQtnsDisp = "";
			var subQtnsLngthCntr = 1;
			var subQtnStr = "";
			for (var subIndex=0; subIndex<subQtns.length; subIndex++) {
				if (subQtns[subIndex] != "NA") {
					//subQtnsLngthCntr = subQtnsLngthCntr + 1;
					//subQtnsDisp = subQtnsDisp + "<b>"+subQtnsLngthCntr+".  </b>"+subQtns[subIndex]+"<br />";
					subQtnsDisp = subQtnsDisp + "<img src='../img/bullet.png' />&nbsp;&nbsp;&nbsp;&nbsp;"+subQtns[subIndex]+"<br />";
					
					subQtnStr = subQtnStr + subQtns[subIndex] + "`";
					subQtnsLngthCntr = subQtnsLngthCntr + 1;
				} else {
					subQtnsDisp = "<b>Not Applicable</b>";
					subQtnStr = subQtnStr + subQtns[subIndex] + "`";
				}
			}
			var isMandatory = (individualRow[7] == "Y") ? "YES" : "NO";
			
			tableStr = tableStr + "<tr class='user_list_row_width' bgcolor='"+trColor+"'><td align='center'>"+counter+".</td><td align='center'>"+individualRow[5]+"</td><td align='center'>"+individualRow[8]+"</td><td align='center'>"+individualRow[6]+"</td><td align='left'>"+individualRow[2]+"</td><td align='center'>"+isMandatory+"</td><td align='left'>"+subQtnsDisp+"</td><td align='center'><a href='#' onClick='editSurveyQtn(\""+individualRow[1]+"\", \""+individualRow[0]+"\", \""+individualRow[2]+"\", \""+individualRow[4]+"\", \""+subQtnStr+"\", \""+isMandatory+"\",\""+individualRow[8]+"\")'><img src = '../img/edit.png' /></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='#' onClick='deleteSurveyQtn(\""+individualRow[1]+"\", \""+individualRow[0]+"\", \""+individualRow[2]+"\")'><img src = '../img/delete.png' /></a></td>";
			counter = counter + 1;
		}
		tableStr = tableStr + "</tbody></table>";
		tableDiv.innerHTML = tableStr;
		//new SortableTable(document.getElementById('myTable'), 1);
		$("table").tablesorter();
	}	
}

function fetchExistingSurveyQuestions() {
	var userGrp = Spine.Model.sub();
	userGrp.configure("/admin/surveyQtns/fetchAllExistingSurveyQtns", "none");
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
		populateTable (obj.resultList, obj.resultList.length);
	});
}

function fetchAllSurveyQuestionsByUserType(userType) {
	if (userType == 0) {
		fetchExistingSurveyQuestions();
	} else {
		var userGrp = Spine.Model.sub();
		userGrp.configure("/admin/surveyQtns/fetchAllExistingSurveyQtnsByUserType", "userType");
		userGrp.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new userGrp({  
			userType: userType
		});
		modelPopulator.save(); //POST
		userGrp.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
		});
		userGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			populateTable (obj.resultList, obj.resultList.length);
		});
	}
}

function fetchAllSurveyQuestionsByUserTypeAndQtnType(qtnType) {
	var e = document.getElementById("userType");
	var value = e.options[e.selectedIndex].value;
	var userType = parseInt(value);
	var userProfile = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value;
	
	// if user type and user profile are available
	if ((userType > 0) && (userProfile != "0")) {
		var userGrp = Spine.Model.sub();
		var profile = userProfile.split("!");
		var profileId = parseInt(profile[0]);
		userGrp.configure("/admin/surveyQtns/fetchAllExistingSurveyQtnsAllSelected", "userType", "qtnType", "profile");
		userGrp.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new userGrp({  
			userType: userType,
			qtnType: qtnType,
			profile: profileId
		});
		modelPopulator.save(); //POST
		userGrp.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
		});
		userGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			populateTable (obj.resultList, obj.resultList.length);
		});
	} else if ((userType == 0) && (userProfile != "0")) {
		var userGrp = Spine.Model.sub();
		var profile = userProfile.split("!");
		var profileId = parseInt(profile[0]);
		userGrp.configure("/admin/surveyQtns/fetchAllExistingSurveyQtnsOnlyUserProfile", "profile", "qtnType");
		userGrp.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new userGrp({  
			userType: userType,
			qtnType: qtnType,
			profile: profileId
		});
		modelPopulator.save(); //POST
		userGrp.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
		});
		userGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			populateTable (obj.resultList, obj.resultList.length);
		});
	} else if ((userType > 0) && (userProfile == "0")) {
		var userGrp = Spine.Model.sub();
		userGrp.configure("/admin/surveyQtns/fetchAllExistingSurveyQtnsByUserTypeAndQtnType", "userType", "qtnType");
		userGrp.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new userGrp({  
			userType: userType,
			qtnType: qtnType
		});
		modelPopulator.save(); //POST
		userGrp.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
		});
		userGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			populateTable (obj.resultList, obj.resultList.length);
		});
	} else {
		var userGrp = Spine.Model.sub();
		userGrp.configure("/admin/surveyQtns/fetchAllExistingSurveyQtnsByQtnType", "qtnType");
		userGrp.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new userGrp({  
			qtnType: qtnType
		});
		modelPopulator.save(); //POST
		userGrp.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
		});
		userGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			populateTable (obj.resultList, obj.resultList.length);
		});
	}
	
	
	/*if (userType > 0) {
		var userGrp = Spine.Model.sub();
		userGrp.configure("/admin/surveyQtns/fetchAllExistingSurveyQtnsByUserTypeAndQtnType", "userType", "qtnType");
		userGrp.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new userGrp({  
			userType: userType,
			qtnType: qtnType
		});
		modelPopulator.save(); //POST
		userGrp.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
		});
		userGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			populateTable (obj.resultList, obj.resultList.length);
		});
	} else {
		var userGrp = Spine.Model.sub();
		userGrp.configure("/admin/surveyQtns/fetchAllExistingSurveyQtnsByQtnType", "qtnType");
		userGrp.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new userGrp({  
			qtnType: qtnType
		});
		modelPopulator.save(); //POST
		userGrp.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
		});
		userGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			populateTable (obj.resultList, obj.resultList.length);
		});
	}*/
}

function plotNoData() {
	var tableDiv = document.getElementById("existingSurveyQtnsTableId");
	tableDiv.innerHTML = "<div align='center'><br /><br /><p><img src='../img/no-record.png'><br /><div class='textStyleNoExist'>No Existing Survey Questions Found</div></p></div>";
	document.getElementById("existing_list_Id").style.display="none";
}

function editSurveyQtn(userType, answerType, qtnDescription, nosOfSubQtns, subQuestions, isMandatory, userProfile) {
	mainQtnMemory = "";
	mainQtnMemory = qtnDescription;
	if (document.getElementById("radiopt")) {
		document.getElementById("radiopt").style.display = "none";
	}
	
	document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].text = userProfile;
	document.getElementById('inputUserProfileInpt').disabled=true;
	/*if( $("#mandatoryChkbox").is(":checked")) {
		$('#mandatoryChkbox').attr('checked', false);
	}*/
	
	
	/**
	 * JCTP-16
	 */	
	if (answerType == 1) {		
		populateDataForText (userType, answerType, qtnDescription, nosOfSubQtns, subQuestions, isMandatory);
		fetchAllSurveyQuestionsByUserTypeAndQtnType(1);
	} else if (answerType == 2) {	//		DONE
		populateDataForRadio (userType, answerType, qtnDescription, nosOfSubQtns, subQuestions, isMandatory);
		fetchAllSurveyQuestionsByUserTypeAndQtnType(2);
	} else if (answerType == 3) {	//		DD
		populateDataForDrop (userType, answerType, qtnDescription, nosOfSubQtns, subQuestions, isMandatory);
		fetchAllSurveyQuestionsByUserTypeAndQtnType(3);
	} else if (answerType == 4) {
		populateDataForCheckBox (userType, answerType, qtnDescription, nosOfSubQtns, subQuestions, isMandatory);
		fetchAllSurveyQuestionsByUserTypeAndQtnType(4);
	}
	var count = $("#radioEntry").children().length;
	for(var i = 0; i<=count-2;i++){
		var textId = "manualRadioQues_"+i;
		$("#"+textId).addClass("survey_radio_width");
	}
	var count = $("#checkboxEntry").children().length;
	for(var i = 0; i<=count-2;i++){
		var textId = "manualCheckQues_"+i;
		$("#"+textId).addClass("survey_check_width");
	}
}

function populateDataForText(userType, answerType, qtnDescription, nosOfSubQtns, subQuestions, isMandatoryOld) {
	
	if (userType == 1) {
		document.getElementById("userType").selectedIndex = "1";
	} else if (userType == 3) {
		document.getElementById("userType").selectedIndex = "2";
	} else {
		document.getElementById("userType").selectedIndex = "0";
	} 
	document.getElementById("userType").disabled = true;
	document.getElementById("text").disabled = true;
	document.getElementById("text").checked = true;
	document.getElementById("rad").disabled = true;
	document.getElementById("drop").disabled = true;
	document.getElementById("check").disabled = true;
	
	
	document.getElementById("textopt").style.display = "block";
	document.getElementById("inputNoOFQues").disabled = true;
	document.getElementById("inputNoOFQues").selectedIndex = "1";
	
	// Create the dynamic element
	var quesNumber = document.getElementById("inputNoOFQues").options[document.getElementById("inputNoOFQues").selectedIndex].value;
	var plotPlement = document.getElementById("quesEntry");
	
	if (document.getElementById("quesEntry")) {
		document.getElementById("quesEntry").innerHTML = "";
		document.getElementById("quesEntry").style.display = "block";
	}
	if (document.getElementById("radioEntry")) {
		document.getElementById("radioEntry").style.display = "none";
	}
	if (document.getElementById("checkboxEntry")) {
		document.getElementById("checkboxEntry").style.display = "none";
	}
	if (document.getElementById("dropdwnEntry")) {
		document.getElementById("dropdwnEntry").style.display = "none";
	}
	var string = "";
	var idCountIndex = 1;
	var mandatoryChkBoxID = "";	
	for (var index = 0; index < quesNumber; index++) {
		var textQuesId = "manualTextQues_"+idCountIndex;
		mandatoryChkBoxID = "mandatoryChkBox_"+idCountIndex;
		
		string = string + "<div class='single_form_item' id='textQtnAnsPltDivId'><div class='col-sm-4 '><label for='inputQtnNo' class='col-sm-12 control-label text_label' id='qtn_div_id"+idCountIndex+"'>Question "+idCountIndex+":</label></div>" +
				"<div class='col-sm-4'><input type='text' class='form-control-general-form-redefined surveyText' value='"+qtnDescription+"' id='"+textQuesId+"' onkeypress='return validateDescription(event)'></div>" +
				"<div class='col-md-2 mandatoryCheckBox'><label for='inputFN' class='text_label' style='margin-left: -13%;'>Mandatory:</label>" +
				"<input type='checkbox'  id='"+mandatoryChkBoxID+"' class='mandatoryChkBox' name='mandatory' value='Y'></div><div class='clearfix'></div></div>"+
				"<div class='clearfix'></div></div>";		
		idCountIndex = idCountIndex + 1;		
	}		
	plotPlement.innerHTML = string + "<input type='button' class='btn_admin btn_admin-primary btn_admin-sm search_btn update_button_ipad search_btn_survey' value='Update' id='updateConfigT'> <input id='resetId' type='button' class='btn_admin btn_admin-primary btn_admin-sm search_btn reset_ipad reset_survey' value='Reset' onclick='resetForm();'>";
	if( isMandatoryOld == "YES" ){								
		$("#"+mandatoryChkBoxID).attr('checked', true);
	}			
	document.getElementById("saveConfig").style.display = "none";
	
	 $("#updateConfigT").click(function () {
	   	 $(".loader_bg").fadeIn();
	     $(".loader").fadeIn();
	     
		     var updatedQtnStr = document.getElementById("manualTextQues_1").value.trim();
				if (updatedQtnStr.length == 0) {
					alertify.alert("Please enter question value");
					$(".loader_bg").fadeOut();
				    $(".loader").hide();
					return false;
				} else if( !specialCharValidation(updatedQtnStr) ){
					return false;			
				} else {
					
					var userGrp = Spine.Model.sub();
			    	userGrp.configure("/admin/surveyQtns/updateTextSurveyQtn", "userType", "answerType", "originalQuestion", "updatedQuestion", "isMandatoryOld", "isMandatory");
			    	userGrp.extend( Spine.Model.Ajax );
			    	//Populate the model with data to transfer
			    	var modelPopulator = new userGrp({  
			    		userType: userType,
			    		answerType: 1,
			    		originalQuestion: mainQtnMemory,
			    		updatedQuestion: document.getElementById("manualTextQues_1").value,
			    		isMandatoryOld: isMandatoryOld == "YES" ? "Y" : "N",
			    		isMandatory: $("#"+mandatoryChkBoxID).is(":checked") ? "Y" : "N"
			    	});
			    	modelPopulator.save(); //POST
			    	userGrp.bind("ajaxError", function(record, xhr, settings, error) {
			    		alertify.alert("Unable to connect to the server.");
			    		$(".loader_bg").fadeOut();
			    	    $(".loader").hide();
			    	});
			    	userGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
			    		var jsonStr = JSON.stringify(xhr);
			    		var obj = jQuery.parseJSON(jsonStr);
			    		if (document.getElementById("loadingId")) {
			    	    	document.getElementById("loadingId").style.display="none";
			    	    }
			    		$(".loader_bg").fadeOut();
			    	    $(".loader").hide();
			    	    alertify.alert(obj.message);
			    		populateTable (obj.resultList, obj.resultList.length);
			    		/**** JIRA ID- JCTP-6 ****/
			    		resetAfterUpdate("text");
			    		return false;
			    	});
				}
		    	
		 });
}


function populateDataForRadio (userType, answerType, qtnDescription, nosOfSubQtns, subQuestions, isMandatory) {
	radioSameValidationMemory = "";
	if (userType == 1) {
		document.getElementById("userType").selectedIndex = "1";
	} else if (userType == 3) {
		document.getElementById("userType").selectedIndex = "2";
	} else {
		document.getElementById("userType").selectedIndex = "0";
	} 
	document.getElementById("userType").disabled = true;
	document.getElementById("rad").disabled = true;
	document.getElementById("rad").checked = true;
	document.getElementById("text").disabled = true;
	document.getElementById("drop").disabled = true;
	document.getElementById("check").disabled = true;
	
	document.getElementById("radioqtn").style.display = "block";
	document.getElementById("questions").value = qtnDescription;
	
	document.getElementById("radiopt").style.display = "block";
	document.getElementById("inputNoOFUser").style.display = "block";
	document.getElementById("inputNoOFUser").disabled = true;
	document.getElementById("inputNoOFUser").selectedIndex = nosOfSubQtns;
	
	if (document.getElementById("textopt")) {
		document.getElementById("textopt").style.display = "none";
	}
	if (document.getElementById("quesEntry")) {
		document.getElementById("quesEntry").style.display = "none";
	}
	if (document.getElementById("radioEntry")) {
		document.getElementById("radioEntry").style.display = "block";
	}
	if (document.getElementById("checkboxEntry")) {
		document.getElementById("checkboxEntry").style.display = "none";
	}
	if (document.getElementById("dropdwnEntry")) {
		document.getElementById("dropdwnEntry").style.display = "none";
	}
	
	radioSameValidationMemory = qtnDescription+"~"+subQuestions;
	
	var ansSplit = subQuestions.split("`");
	var plotPlement = document.getElementById("radioEntry");
	var string = "";
	var idIndex = 1;
	for (var index = 0; index < ansSplit.length-1; index++) {
		var radioQuesId = "manualRadioQues_"+idIndex;
		string = string + "<div class='single_form_item'><div class='col-sm-4 '><label for='inputQtnNo' class='col-sm-12 control-label text_label' id='qtn_div_id"+idIndex+"'>Option "+idIndex+":</label></div>" +
		"<div class='col-sm-4'><input type='text' class='form-control-general-form-redefined' value='"+ansSplit[index]+"' id='"+radioQuesId+"' onkeypress='return validateDescription(event)'></div>" +
		"<div class='col-md-2'>&nbsp;</div><div class='clearfix'></div></div>";		
		idIndex = idIndex + 1;
	}
	plotPlement.innerHTML = string + "<input type='button' class='btn_admin btn_admin-primary btn_admin-sm search_btn update_button_ipad search_btn_survey' value='Update' id='updateConfigR'> <input id='resetId' type='button' class='btn_admin btn_admin-primary btn_admin-sm search_btn reset_ipad reset_survey' value='Reset' onclick='resetForm();'>";
	//plotPlement.innerHTML = string + "<input type='button' class='btn_admin btn_admin-primary btn_admin-sm search_btn search_btn_survey' value='Reset' onclick='resetForm();'>";
	document.getElementById("saveConfig").style.display = "none";
	
	if( isMandatory == "YES" ){			
		document.getElementById("mandatoryChkbox").checked = true;
	}
	
	
	 $("#updateConfigR").click(function () {
		 //document.getElementById("loadingId").style.display="block";
 	     $(".loader_bg").fadeIn();
	    	$(".loader").fadeIn();
	    	var subQtns = "";
	    	var mySubQtns = "";	    	
	    	var userType = parseInt(document.getElementById("userType").options[document.getElementById("userType").selectedIndex].value);
	    	var nosOfSubQuestions = parseInt(document.getElementById("inputNoOFUser").options[document.getElementById("inputNoOFUser").selectedIndex].value);
	    	for (var index = 1; index<=nosOfSubQuestions; index++) {
	    		var temp = document.getElementById("manualRadioQues_"+index).value.trim();
	    		if (temp.length == 0) {
	    			alertify.alert('Please enter all the sub-question values.');
	    			$(".loader_bg").fadeOut();
		    	    $(".loader").hide();
	    			return false;
	    		} else if( !specialCharValidation(temp) ){
					return false;			
				}
	    		subQtns = subQtns + document.getElementById("manualRadioQues_"+index).value+"`~!~`";
	    		mySubQtns = mySubQtns + document.getElementById("manualRadioQues_"+index).value+"`";
	    	}
	    	var mainQues = document.getElementById("questions").value.trim();
	    	if (mainQues.length == 0) {
	    		alertify.alert('Please enter question value.');
    			$(".loader_bg").fadeOut();
	    	    $(".loader").hide();
    			return false;
	    	} else if( !specialCharValidation(mainQues) ){
				return false;			
			}
	    	
	    	var isMandatoryChangeBoolean = ($("#mandatoryChkbox").is(":checked") ? "YES" : "NO") == isMandatory;	 
	    	if ((radioSameValidationMemory == document.getElementById("questions").value+"~"+mySubQtns) && isMandatoryChangeBoolean){
	    		alertify.alert('Please make a change to the existing question / sub-question / Mandatory-option to proceed.');
	    		$(".loader_bg").fadeOut();
	    	    $(".loader").hide();
		    	return false;
	    	}	    	
	    	
	    	var userGrp = Spine.Model.sub();
	    	userGrp.configure("/admin/surveyQtns/updateRadioSurveyQtn", "userType", "answerType", "originalQuestion", "updatedMainQuestion", "subQuestion", "createdBy", "isMandatory");
	    	userGrp.extend( Spine.Model.Ajax );
	    	//Populate the model with data to transfer
	    	var modelPopulator = new userGrp({  
	    		userType: userType,
	    		answerType: 2,
	    		originalQuestion: mainQtnMemory,
	    		updatedMainQuestion: document.getElementById("questions").value,
	    		subQuestion: subQtns,
	    		createdBy: sessionStorage.getItem("jctEmail"),
	    		isMandatory: $("#mandatoryChkbox").is(":checked") ? "Y" : "N"
	    	});
	    	modelPopulator.save(); //POST
	    	userGrp.bind("ajaxError", function(record, xhr, settings, error) {
	    		alertify.alert("Unable to connect to the server.");
	    		$(".loader_bg").fadeOut();
	    	    $(".loader").hide();
	    	});
	    	userGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
	    		var jsonStr = JSON.stringify(xhr);
	    		var obj = jQuery.parseJSON(jsonStr);
	    		if (document.getElementById("loadingId")) {
	    	    	document.getElementById("loadingId").style.display="none";
	    	    }
	    		$(".loader_bg").fadeOut();
	    	    $(".loader").hide();
	    	    alertify.alert(obj.message);
	    		populateTable (obj.resultList, obj.resultList.length);
	    		/**** JIRA ID- JCTP-6 ****/
	    		resetAfterUpdate("radio");
	    	});
	 });
}

function populateDataForDrop (userType, answerType, qtnDescription, nosOfSubQtns, subQuestions, isMandatory) {
	dropSameValidationMemory = "";
	if (userType == 1) {
		document.getElementById("userType").selectedIndex = "1";
	} else if (userType == 3) {
		document.getElementById("userType").selectedIndex = "2";
	} else {
		document.getElementById("userType").selectedIndex = "0";
	} 
	document.getElementById("userType").disabled = true;
	document.getElementById("drop").disabled = true;
	document.getElementById("drop").checked = true;
	document.getElementById("text").disabled = true;
	document.getElementById("rad").disabled = true;
	document.getElementById("check").disabled = true;
	
	document.getElementById("radioqtn").style.display = "block";
	document.getElementById("questions").value = qtnDescription;
	
	document.getElementById("radiopt").style.display = "block";
	document.getElementById("inputNoOFUser").style.display = "block";
	document.getElementById("inputNoOFUser").disabled = true;
	document.getElementById("inputNoOFUser").selectedIndex = nosOfSubQtns;
	if (document.getElementById("textopt")) {
		document.getElementById("textopt").style.display = "none";
	}
	if (document.getElementById("quesEntry")) {
		document.getElementById("quesEntry").style.display = "none";
	}
	if (document.getElementById("radioEntry")) {
		document.getElementById("radioEntry").style.display = "none";
	}
	if (document.getElementById("checkboxEntry")) {
		document.getElementById("checkboxEntry").style.display = "none";
	}
	if (document.getElementById("dropdwnEntry")) {
		document.getElementById("dropdwnEntry").style.display = "block";
	}	
	dropSameValidationMemory = qtnDescription+"~"+subQuestions;
	
	var ansSplit = subQuestions.split("`");
	var plotPlement = document.getElementById("dropdwnEntry");
	var string = "";
	var idIndex = 1;
	for (var index = 0; index < ansSplit.length-1; index++) {
		var radioQuesId = "manualDropQues_"+idIndex;
		string = string + "<div class='single_form_item'><div class='col-sm-4 '><label for='inputQtnNo' class='col-sm-12 control-label text_label' id='qtn_div_id"+idIndex+"'>Option "+idIndex+":</label></div>" +
		"<div class='col-sm-4'><input type='text' class='form-control-general-form-redefined surveyText' value='"+ansSplit[index]+"' id='"+radioQuesId+"' onkeypress='return validateDescription(event)'></div>" +
		"<div class='col-md-2'>&nbsp;</div><div class='clearfix'></div></div>";		
		idIndex = idIndex + 1;
	}
	plotPlement.innerHTML = string + "<input type='button' class='btn_admin btn_admin-primary btn_admin-sm search_btn update_button_ipad search_btn_survey' value='Update' id='updateConfigD'> <input id='resetId' type='button' class='btn_admin btn_admin-primary btn_admin-sm search_btn reset_ipad reset_survey' value='Reset' onclick='resetForm();'>";
	//plotPlement.innerHTML = string + "<input type='button' class='btn_admin btn_admin-primary btn_admin-sm search_btn search_btn_survey' value='Reset' onclick='resetForm();'>";
	document.getElementById("saveConfig").style.display = "none";
	
	if( isMandatory == "YES" ){	
		document.getElementById("mandatoryChkbox").checked = true;
	}
	
	 $("#updateConfigD").click(function () {
		 //document.getElementById("loadingId").style.display="block";
 	     $(".loader_bg").fadeIn();
	    	$(".loader").fadeIn();
	    	var subQtns = "";
	    	var mySubQtns = "";	    		    	
	    	var userType = parseInt(document.getElementById("userType").options[document.getElementById("userType").selectedIndex].value);
	    	var nosOfSubQuestions = parseInt(document.getElementById("inputNoOFUser").options[document.getElementById("inputNoOFUser").selectedIndex].value);
	    	for (var index = 1; index<=nosOfSubQuestions; index++) {
	    		var temp = document.getElementById("manualDropQues_"+index).value.trim();
	    		if (temp.length == 0) {
	    			alertify.alert('Please enter all the sub-question values.');
	    			$(".loader_bg").fadeOut();
		    	    $(".loader").hide();
	    			return false;
	    		} else if( !specialCharValidation(temp) ){
					return false;			
				}
	    		subQtns = subQtns + document.getElementById("manualDropQues_"+index).value+"`~!~`";
	    		mySubQtns = mySubQtns + document.getElementById("manualDropQues_"+index).value+"`";
	    	}
	    	var mainQues = document.getElementById("questions").value.trim();
	    	if (mainQues.length == 0) {
	    		alertify.alert('Please enter question value.');
    			$(".loader_bg").fadeOut();
	    	    $(".loader").hide();
    			return false;
	    	} else if( !specialCharValidation(mainQues) ){
				return false;			
			}
	    	var isMandatoryChangeBoolean = ($("#mandatoryChkbox").is(":checked") ? "YES" : "NO") == isMandatory;
	    	if ((dropSameValidationMemory == document.getElementById("questions").value+"~"+mySubQtns) && isMandatoryChangeBoolean) {
	    		alertify.alert('Please make a change to the existing question / sub-question  / Mandatory-option to proceed.');
	    		$(".loader_bg").fadeOut();
	    	    $(".loader").hide();
		    	return false;
	    	}
	    	var userGrp = Spine.Model.sub();
	    	userGrp.configure("/admin/surveyQtns/updateDropdownSurveyQtn", "userType", "answerType", "originalQuestion", "updatedMainQuestion", "subQuestion", "createdBy", "isMandatory");
	    	userGrp.extend( Spine.Model.Ajax );
	    	//Populate the model with data to transfer
	    	var modelPopulator = new userGrp({  
	    		userType: userType,
	    		answerType: 3,
	    		originalQuestion: mainQtnMemory,
	    		updatedMainQuestion: document.getElementById("questions").value,
	    		subQuestion: subQtns,
	    		createdBy: sessionStorage.getItem("jctEmail"),
	    		isMandatory: $("#mandatoryChkbox").is(":checked") ? "Y" : "N"
	    	});
	    	modelPopulator.save(); //POST
	    	userGrp.bind("ajaxError", function(record, xhr, settings, error) {
	    		alertify.alert("Unable to connect to the server.");
	    		$(".loader_bg").fadeOut();
	    	    $(".loader").hide();
	    	});
	    	userGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
	    		var jsonStr = JSON.stringify(xhr);
	    		var obj = jQuery.parseJSON(jsonStr);
	    		if (document.getElementById("loadingId")) {
	    	    	document.getElementById("loadingId").style.display="none";
	    	    }
	    		$(".loader_bg").fadeOut();
	    	    $(".loader").hide();
	    	    alertify.alert(obj.message);
	    		populateTable (obj.resultList, obj.resultList.length);
	    		/**** JIRA ID- JCTP-6 ****/
	    		resetAfterUpdate("dropdown");
	    	});
	 });
}

function populateDataForCheckBox (userType, answerType, qtnDescription, nosOfSubQtns, subQuestions, isMandatory) {
	checkSameValidationMemory = "";
	if (userType == 1) {
		document.getElementById("userType").selectedIndex = "1";
	} else if (userType == 3) {
		document.getElementById("userType").selectedIndex = "2";
	} else {
		document.getElementById("userType").selectedIndex = "0";
	} 
	document.getElementById("userType").disabled = true;
	document.getElementById("check").disabled = true;
	document.getElementById("check").checked = true;
	document.getElementById("text").disabled = true;
	document.getElementById("rad").disabled = true;
	document.getElementById("drop").disabled = true;
	
	document.getElementById("radioqtn").style.display = "block";
	document.getElementById("questions").value = qtnDescription;
	
	document.getElementById("radiopt").style.display = "block";
	document.getElementById("inputNoOFUser").style.display = "block";
	document.getElementById("inputNoOFUser").disabled = true;
	document.getElementById("inputNoOFUser").selectedIndex = nosOfSubQtns;
	if (document.getElementById("textopt")) {
		document.getElementById("textopt").style.display = "none";
	}
	if (document.getElementById("quesEntry")) {
		document.getElementById("quesEntry").style.display = "none";
	}
	if (document.getElementById("radioEntry")) {
		document.getElementById("radioEntry").style.display = "none";
	}
	if (document.getElementById("dropdwnEntry")) {
		document.getElementById("dropdwnEntry").style.display = "none";
	}
	if (document.getElementById("checkboxEntry")) {
		document.getElementById("checkboxEntry").style.display = "block";
	}
	
	checkSameValidationMemory = qtnDescription+"~"+subQuestions;
	
	var ansSplit = subQuestions.split("`");
	var plotPlement = document.getElementById("checkboxEntry");
	var string = "";
	var idIndex = 1;
	for (var index = 0; index < ansSplit.length-1; index++) {
		var radioQuesId = "manualCheckQues_"+idIndex;
		string = string + "<div class='single_form_item'><div class='col-sm-4 '><label for='inputQtnNo' class='col-sm-12 control-label text_label' id='qtn_div_id"+idIndex+"'>Option "+idIndex+":</label></div>" +
		"<div class='col-sm-6'><input type='text' class='form-control-general-form-redefined' value='"+ansSplit[index]+"' id='"+radioQuesId+"' onkeypress='return validateDescription(event)'></div>" +
		"<div class='col-md-2'>&nbsp;</div><div class='clearfix'></div></div>";		
		idIndex = idIndex + 1;
	}
	plotPlement.innerHTML = string + "<input type='button' class='btn_admin btn_admin-primary btn_admin-sm search_btn update_button_ipad search_btn_survey' value='Update' id='updateConfigC'> <input id='resetId' type='button' class='btn_admin btn_admin-primary btn_admin-sm search_btn reset_ipad reset_survey' value='Reset' onclick='resetForm();'>";
	//plotPlement.innerHTML = string + "<input type='button' class='btn_admin btn_admin-primary btn_admin-sm search_btn search_btn_survey' value='Reset' onclick='resetForm();'>";
	document.getElementById("saveConfig").style.display = "none";
	
	if( isMandatory == "YES" ){	
		document.getElementById("mandatoryChkbox").checked = true;		
	}
	
	 $("#updateConfigC").click(function () {
		//document.getElementById("loadingId").style.display="block";
 	     $(".loader_bg").fadeIn();
	    	$(".loader").fadeIn();
	    	var subQtns = "";
	    	var mySubQtns = "";	    	
	    	var userType = parseInt(document.getElementById("userType").options[document.getElementById("userType").selectedIndex].value);
	    	var nosOfSubQuestions = parseInt(document.getElementById("inputNoOFUser").options[document.getElementById("inputNoOFUser").selectedIndex].value);
	    	for (var index = 1; index<=nosOfSubQuestions; index++) {
	    		var temp = document.getElementById("manualCheckQues_"+index).value.trim();
	    		if (temp.length == 0) {
	    			alertify.alert('Please enter all the sub-question values.');
	    			$(".loader_bg").fadeOut();
		    	    $(".loader").hide();
	    			return false;
	    		} else if( !specialCharValidation(temp) ){
					return false;			
				}
	    		subQtns = subQtns + document.getElementById("manualCheckQues_"+index).value+"`~!~`";
	    		mySubQtns = mySubQtns + document.getElementById("manualCheckQues_"+index).value+"`";
	    	}
	    	var mainQues = document.getElementById("questions").value;
	    	if (mainQues.trim().length == 0) {
	    		alertify.alert('Please enter question value.');
    			$(".loader_bg").fadeOut();
	    	    $(".loader").hide();
    			return false;
	    	} else if( !specialCharValidation(mainQues) ){
				return false;			
			}
	    	var isMandatoryChangeBoolean = ($("#mandatoryChkbox").is(":checked") ? "YES" : "NO") == isMandatory;
	    	if ((checkSameValidationMemory == document.getElementById("questions").value+"~"+mySubQtns) && isMandatoryChangeBoolean) {
	    		alertify.alert('Please make a change to the existing question / sub-question / Mandatory-option to proceed.');
	    		$(".loader_bg").fadeOut();
	    	    $(".loader").hide();
		    	return false;
	    	}
	    	var userGrp = Spine.Model.sub();
	    	userGrp.configure("/admin/surveyQtns/updateCheckboxSurveyQtn", "userType", "answerType", "originalQuestion", "updatedMainQuestion", "subQuestion", "createdBy", "isMandatory");
	    	userGrp.extend( Spine.Model.Ajax );
	    	//Populate the model with data to transfer
	    	var modelPopulator = new userGrp({  
	    		userType: userType,
	    		answerType: 4,
	    		originalQuestion: mainQtnMemory,
	    		updatedMainQuestion: document.getElementById("questions").value,
	    		subQuestion: subQtns,
	    		createdBy: sessionStorage.getItem("jctEmail"),
	    		isMandatory: $("#mandatoryChkbox").is(":checked") ? "Y" : "N"
	    	});
	    	modelPopulator.save(); //POST
	    	userGrp.bind("ajaxError", function(record, xhr, settings, error) {
	    		alertify.alert("Unable to connect to the server.");
	    		$(".loader_bg").fadeOut();
	    	    $(".loader").hide();
	    	});
	    	userGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
	    		var jsonStr = JSON.stringify(xhr);
	    		var obj = jQuery.parseJSON(jsonStr);
	    		$(".loader_bg").fadeOut();
	    	    $(".loader").hide();
	    	    if (document.getElementById("loadingId")) {
	    	    	document.getElementById("loadingId").style.display="none";
	    	    }
	    	    alertify.alert(obj.message);
	    		populateTable (obj.resultList, obj.resultList.length);
	    		/**** JIRA ID- JCTP-6 ****/
	    		resetAfterUpdate("checkbox");
	    	});
	 });
}

function deleteSurveyQtn(userType, answerType, qtnDescription) {
	alertify.set({ buttonReverse: true });
	alertify.confirm("Are you sure you want to delete the survey question?", function (e) {
	    if (e) {		
	    	var deciderParam = "";
	    	var e = document.getElementById("userType");
	    	var value = e.options[e.selectedIndex].value;
	    	if (value == "") {
	    		deciderParam = "0`";
	    	} else if (value == "1") {
	    		deciderParam = "1`";
	    	} else if (value == "3") {
	    		deciderParam = "3`";
	    	} else {
	    		deciderParam = "0`";
	    	}
	    	
	    	if (document.getElementById("text").checked) {
	    		deciderParam = deciderParam + "1";
	    	} else if (document.getElementById("rad").checked) {
	    		deciderParam = deciderParam + "2";
	    	} else if (document.getElementById("drop").checked) {
	    		deciderParam = deciderParam + "3";
	    	} else if (document.getElementById("check").checked) {
	    		deciderParam = deciderParam + "4";
	    	} else {
	    		deciderParam = deciderParam + "5"; //None of the option
	    	}
	    	$(".loader_bg").fadeIn();
	    	$(".loader").fadeIn();
	    	var userGrp = Spine.Model.sub();
	    	userGrp.configure("/admin/surveyQtns/deleteSurveyQtn", "userType", "answerType", "qtnDescription", "deciderParam");
	    	userGrp.extend( Spine.Model.Ajax );
	    	//Populate the model with data to transfer
	    	var modelPopulator = new userGrp({  
	    		userType: userType,
	    		answerType: answerType,
	    		qtnDescription: qtnDescription,
	    		deciderParam: deciderParam
	    	});
	    	modelPopulator.save(); //POST
	    	userGrp.bind("ajaxError", function(record, xhr, settings, error) {
	    		alertify.alert("Unable to connect to the server.");
	    		$(".loader_bg").fadeOut();
	    	    $(".loader").hide();
	    	});
	    	userGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
	    		var jsonStr = JSON.stringify(xhr);
	    		var obj = jQuery.parseJSON(jsonStr);
	    		alertify.alert(obj.message);
	    		$(".loader_bg").fadeOut();
	    	    $(".loader").hide();
	    		populateTable (obj.resultList, obj.resultList.length);
	    	});
	    } 
	});
}

function resetForm() {
	var userType = parseInt(document.getElementById("userType").options[document.getElementById("userType").selectedIndex].value);
	if (userType == 1) {
		document.getElementById("userType").selectedIndex = "1";
	} else {
		document.getElementById("userType").selectedIndex = "2";
	}
	document.getElementById("userType").disabled = false;
	if (document.getElementById("text").checked) {
		document.getElementById("text").checked = true;
		document.getElementById("inputNoOFQues").disabled = false;
		document.getElementById("inputNoOFQues").selectedIndex = "0";
		document.getElementById("textQtnAnsPltDivId").innerHTML = "";
		document.getElementById("updateConfigT").style.display = "none";
		document.getElementById("quesEntry").innerHTML = "";
	} else if (document.getElementById("rad").checked) {
		document.getElementById("rad").checked = true;
		document.getElementById("questions").value = "";
		document.getElementById("inputNoOFUser").disabled = false;
		document.getElementById("inputNoOFUser").selectedIndex = "0";
		document.getElementById("radioEntry").innerHTML = "";
	} else if (document.getElementById("drop").checked) {
		document.getElementById("drop").checked = true;
		document.getElementById("questions").value = "";
		document.getElementById("inputNoOFUser").disabled = false;
		document.getElementById("inputNoOFUser").selectedIndex = "0";
		document.getElementById("dropdwnEntry").innerHTML = "";
	} else if (document.getElementById("check").checked) {
		document.getElementById("check").checked = true;
		document.getElementById("questions").value = "";
		document.getElementById("inputNoOFUser").disabled = false;
		document.getElementById("inputNoOFUser").selectedIndex = "0";
		document.getElementById("checkboxEntry").innerHTML = "";
	}
	
	$('#mandatoryChkbox').attr('checked', false);
	document.getElementById("text").disabled = false;
	document.getElementById("rad").disabled = false;
	document.getElementById("drop").disabled = false;
	document.getElementById("check").disabled = false;
	if (document.getElementById("resetId")) {
		document.getElementById("resetId").style.display = "none";
	}
}

/**
 * Function add to allow only alphabets and numbers
 * and some special character like ?, . 
 * @param event
 */
function validateDescription(event) {	
	 var c = event.which || event.keyCode;
	 if(( c == 33 ) || (c > 34 && c < 39) || (c > 39 && c < 44) || (c > 46 && c < 58) || 
			 (c > 59 && c < 63) || (c > 63 && c < 65) || (c > 90 && c < 97) ||
			     (c > 122 && c !== 127))
	    return false;
	  
}
/**
 * more than one special char aren't allowed
 * @param text 
 **/
function specialCharValidation(text){	
	//var text = $("#target").val();
	if(text.search(/[?:"']{2,}/) > 0) {
		alertify.alert("Consecutive special character is not allowed");
		return false;
	}
	return true;
}

/**** JIRA ID- JCTP-6 ****/
/**
 * Function call while update the value
 */
function resetAfterUpdate(inputOption) {
	var userType = document.getElementById("userType");
	userType.removeAttribute("disabled", "true");
	document.getElementsByName("radios")[0].disabled = false;
	document.getElementsByName("radios")[1].disabled = false;
	document.getElementsByName("radios")[2].disabled = false;
	document.getElementsByName("radios")[3].disabled = false;
	if(inputOption == "text") {
		document.getElementById("inputNoOFQues").selectedIndex="";
		document.getElementById("inputNoOFQues").disabled = false;
		document.getElementById("quesEntry").innerHTML = "";
	} else if (inputOption == "radio") {
		document.getElementById("questions").value = "";
		document.getElementById("questions").disabled = false;
		document.getElementById("inputNoOFUser").selectedIndex="";
		document.getElementById("inputNoOFUser").disabled = false;
		document.getElementById("radioEntry").innerHTML = "";
		
	} else if (inputOption == "dropdown") {
		document.getElementById("questions").value = "";
		document.getElementById("questions").disabled = false;
		document.getElementById("inputNoOFUser").selectedIndex="";
		document.getElementById("inputNoOFUser").disabled = false;
		document.getElementById("dropdwnEntry").innerHTML = "";	
		
	}  else if (inputOption == "checkbox") {
		document.getElementById("questions").value = "";
		document.getElementById("questions").disabled = false;
		document.getElementById("inputNoOFUser").selectedIndex="";
		document.getElementById("inputNoOFUser").disabled = false;
		document.getElementById("checkboxEntry").innerHTML = "";
	}	
	if( $("#mandatoryChkbox").is(":checked")) {
		$('#mandatoryChkbox').attr('checked', false);
	}	
}

function changeOrderSurvQtn() {	
	var e = document.getElementById("userType");
	var value = e.options[e.selectedIndex].value;
	var userType = parseInt(value);
	var userProfile = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value;
	var profile = userProfile.split("!");
	var profileId = parseInt(profile[0]);
	var userGrp = Spine.Model.sub();
	userGrp.configure("/admin/surveyQtns/fetchSurveyQtnsByProfileAndUserType", "profile", "userType");
	userGrp.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userGrp({  
		profile: profileId,
		userType: userType
	});
	modelPopulator.save(); //POST
	userGrp.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	userGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		//populateTable (obj.resultList, obj.resultList.length);
		var statusCode = obj.statusCode;
		if(statusCode == 200) {						
			populateSurveyTable(obj.resultList);	
		} else if(statusCode == 201) {
			document.getElementById("existingUserRoleListId").innerHTML = "<div align='center'>"+obj.statusDesc+"</div>";
		} else if(statusCode == 500) {
			//Show error message
		}
	});
}


function populateSurveyTable(listData) {		
	var tableDiv = document.getElementById("mainAreaListId");
	if (listData.length > 0) {
		var headingString = "<table width='90%' id='table-1' border='1' align='center' class='tablesorter'><thead align='center' class='tab_header'><tr><th><b>Question Description</b></th><th><b>Answer Type</b></th></tr></thead><tbody>";		
		var rowSplit = listData.split("~");
		var counter = 1;
		for (var rowIndex=0; rowIndex<rowSplit.length-1; rowIndex++) {
			var individualRow = rowSplit[rowIndex].split("`");
			headingString = headingString + "<tr align='center' id="+counter+" class='custom_table_row' onmousedown='highlight()'><td align='center'>"+individualRow[2]+"</td><td align='center'>"+individualRow[6]+"</td>";	
			headingString = headingString + "</tr>";
			counter = counter + 1;
		}
		
		headingString = headingString + "</tbody></table>";
		tableDiv.innerHTML = headingString;
		document.getElementById("attrSaveBtnId").style.display = "block";
		
	} else {
		var headingString = "<table width='90%' id='table-1' border='1' align='center' class=''><thead align='center' class='tab_header'><tr><td><b>NO DATA FOUND</b></td></tr></thead><tbody>";
		headingString = headingString + "</tbody></table>";
		headingString = headingString + "</br>";
		tableDiv.innerHTML = headingString;
		document.getElementById("attrSaveBtnId").style.display = "none";
	}
	 $("#table-1").tableDnD();
	 $("table").tablesorter(); 
}
function highlight() {
	$("tr").mousedown(function(){
		$(this).addClass("selected").siblings().removeClass("selected");
	});
}

/**
 * Function to save survey question 
 * by survey question order
 * @param null
 */
function saveSurveyQtnOrder() {
	
	var e = document.getElementById("userType");
	var value = e.options[e.selectedIndex].value;
	var userType = parseInt(value);
	
	var userProfileDesc = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value;
	var profile = userProfileDesc.split("!");
	var userProfile = parseInt(profile[0]);
	
	var builder = "";
	var builderAnsType = "";
	$('#table-1 tbody tr td:first-child').each(function() {
	    var cellText = $(this).html();  
	    builder = builder + cellText +"~~";	
	});
	$('#table-1 tbody tr td:nth-child(2)').each(function() {
	    var cellText = $(this).html();  
	    builderAnsType = builderAnsType + cellText +"~~";	
	});
	var newUserProf = Spine.Model.sub();
	newUserProf.configure("/admin/surveyQtns/saveSurvryQtnOrder", "builder", "userProfile", "userType", "builderAnsType");
	newUserProf.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new newUserProf({  
		builder: builder,
		userProfile: userProfile,
		userType: userType,
		builderAnsType: builderAnsType
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
			alertify.alert("Survey Question Order has been successfully changed.");
			//populateAttributeTable(obj.mainQtnList);
			//fetchMappingList(selectedRadioButton, userProfile , attrCode);
			populateTable (obj.resultList, obj.resultList.length);
		}  else if(statusCode == 500) {
			// Error Message
		}
	});
}