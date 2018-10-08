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


/** JIRA ID - JCTP-10
 * Function add to disable browser back button
 * while the page is loaded
 * @param null
 */
window.location.hash="";
window.location.hash="Again-No-back-button";//again because google chrome don't insert first hash into history
window.onhashchange=function(){window.location.hash="";};


$(document).ready( function() {
	populateSignupMasterData();
});



/**
 * Method populates master dropdown for location, function group and job level
 */
function populateSignupMasterData(){
	//Create a model
	var master = Spine.Model.sub();
	master.configure("/admin/authAdmin/populateSignupMasterData", "none");
	master.extend( Spine.Model.Ajax );
	
	//Populate the model
	var populateMaster = new master({  
		none: ""
	});
	populateMaster.save(); //POST
	master.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to fetch the survey question.");
	});
	
	master.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);		
		var statusCode = obj.statusCode;
		if(statusCode == "200"){						
			if (obj.surveyTextList == "") {
				document.getElementById('surveyQtnNextbtnId').setAttribute("style", "display:block");
				document.getElementById('surveyQtnSavebtnId').setAttribute("style", "display:none");
				
				var plottingDiv = document.getElementById("surveyQuestionsDivId");
				//plottingDiv.innerHTML = "No Survey Qustion is being configured for you. <br>Please click Next to continue.";
				
				var text = "";
				text = text + "<div class='single_form_item'>";
				text = text + "<div class='col-md-3'></div>";
				text = text + "<div class='col-md-6 no_survey_qtn'>No Survey Qustion is being configured for you.</div>";
				text = text + "<div class='col-md-3'></div><div class='clearfix'></div>";
				text = text + "</div>";
				
				text = text + "<div class='single_form_item'>";
				text = text + "<div class='col-md-3'></div>";
				text = text + "<div class='col-md-6 no_survey_qtn'>Please click Next to continue.</div>";
				text = text + "<div class='col-md-3'></div><div class='clearfix'></div>";
				text = text + "</div><br>";
				plottingDiv.innerHTML = text;
			} else {
				document.getElementById('surveyQtnNextbtnId').setAttribute("style", "display:none");
				document.getElementById('surveyQtnSavebtnId').setAttribute("style", "display:block");
				var textSplit = obj.surveyTextList.split("<>");
				var plottingDiv = document.getElementById("surveyQuestionsDivId");
				var text = "";
				for (var index=0; index < textSplit.length; index++) {
					// check the type of question
					var type = textSplit[index].split("`");
					var mandatory = textSplit[index].split("#");
					if (type[0] == 1) {
						plottedText = 1;
						inputTextSurveyId = inputTextSurveyId + "textQtnId_"+index+"`"+ "textAnsId_"+index+"~~";
						var qtn = mandatory[0].split("`");
						text = text + "<div class='single_form_item'>";
						text = text + "<div class='col-md-4 '>";
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
			
		}
	});
}


function saveSurveyDeatils() {
	$(".loader_bg").fadeIn();
    $(".loader").fadeIn(); 

    
    if( validateMandatoryFields() == true){
    	if($('#iAgree').is(':checked')) {
	    	// get the survey question and answer strings
	    	var surveyQtnTextString = getSurveyQuestionTextString();
	    	var surveyQtnRadioString = getSurveyQuestionRadioString();
	    	var surveyQtnCheckBoxString = getSurveyQuestionCheckboxString();
	    	var surveyQtnDropdownString = getSurveyQuestionDropdownString();
	    	 //Create a model
	    	var stub = Spine.Model.sub();
	    	stub.configure("/admin/authAdmin/saveSurveyQuestions", "userId",  "emailIds",
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
	    		window.location = "createUserFacilitator.jsp";
	    	});
    	} else {
    		$(".loader_bg").fadeOut();
    	    $(".loader").hide();
        	alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please agree to the terms & conditions to proceed.</p>");        	
        	return false;
        } 
    } else {
    	$(".loader_bg").fadeOut();
	    $(".loader").hide();
	    alertify.alert("Please provide data for all mandatory field.");
    	return false;
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
			return false;		    				
		}
		  
		if( fieldType == 'radio' && ($("input[id='"+fieldId+"']:checked").length == 0) ){
			flag = 1;
			return false;		    				
		}
				
		/*if( fieldType == 'checkbox' && (!$("#"+fieldId).is(":checked")) ){
			flag = 1;
			return false;		    				
		}*/		
		if( fieldType == 'checkbox' && ($("input[id='"+fieldId+"']:checked").length == 0) ){
			flag = 1;
			return false;		    				
		}
		
		 	        
		if( fieldType == 'dropdown') {			 
			var ddvalue = document.getElementById(fieldId).options[document.getElementById(fieldId).selectedIndex].value;		
			if(ddvalue == '' || ddvalue == 'NOTANSWERED') {				
				 flag = 1;
				 return false;
			}
		 }
	});
	
	return (flag ==1 ? false : true);	
}

function updateActiveStatus() {
	$(".loader_bg").fadeIn();
    $(".loader").fadeIn();
	 //Create a model
	var stub = Spine.Model.sub();
	stub.configure("/admin/authAdmin/updateActiveStatusFacilitator", "userId",  "emailIds");
	stub.extend( Spine.Model.Ajax );
		
	//Populate the model with data to transfer
	var modelPopulation = new stub({ 
		userId: sessionStorage.getItem("jctUserId"),
		emailIds: sessionStorage.getItem("jctEmail")
	});
	modelPopulation.save(); //POST the data
	
	stub.bind("ajaxError", function(record, xhr, settings, error) {
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
		alertify.alert("Unable to connect to the server.");
	});
	
	stub.bind("ajaxSuccess", function(record, xhr, settings, error) {
		window.location = "createUserFacilitator.jsp";
	});
}