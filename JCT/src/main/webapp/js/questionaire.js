var questionIndex = 1;
var subQuestionIndex = 1;
var answerTracker = 1;
var questionAnsTracker = "";
var answerString = "";

window.setInterval(function(){
	updateTime();
}, 60000);

$(document).ready(function() { 
	sessionStorage.setItem("pageSequence", 2);
	var wrapperHeight = parseInt(sessionStorage.getItem("bsDivHeight")) + 100;
	
	if(navigator.userAgent.match(/iPhone|iPad|iPod/i) ){
		$(".form_area_wraper").css({height : wrapperHeight+"px"});
	}
	populatequestionnaire();
	//$('#myModal').modal('show');
	/*var insDisplay = sessionStorage.getItem("qsView");
	if(null == insDisplay || insDisplay == "" || insDisplay == "D"){
		populatePopUp();
	
	} else {
		$('#myModal').modal('hide');
	}*/
});

/**
 * Function add to disable browser back button
 * while the page is loaded
 * @param null
 */
window.location.hash="";
window.location.hash="Again-No-back-button";//again because google chrome don't insert first hash into history
window.onhashchange=function(){window.location.hash="";};


/**
 * Function populates questionnaire on document ready
 * @param null
 */
function populatequestionnaire() {
	$(".loader_bg").fadeIn();
	$(".loader").fadeIn();
	var navigation = Spine.Model.sub();
	navigation.configure("/user/questionnaire/fetchQuestionnaire", "jobReferenceNo", "profileId");
	navigation.extend( Spine.Model.Ajax );
	var replyModel = new navigation({
		jobReferenceNo: sessionStorage.getItem("jobReferenceNo"),
		profileId: sessionStorage.getItem("profileId")
	});
	replyModel.save(); //POST
	
	navigation.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
		return false;
	});
	
	navigation.bind("ajaxSuccess", function(record, xhr, settings, error) {
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var htmlString = "";
		for (var populationIndex=0; populationIndex<obj.questionnaireList.length; populationIndex++) {
			var individualArrItem = obj.questionnaireList[populationIndex];
			var countIndex = 1;
			var innerIndex1 = 0;
			var subQtn = 1;
			var subQtnMarkers = ["a", "b", "c", "d","a", "b", "c", "d","a", "b", "c", "d","a", "b", "c", "d"];
			//alert(jsonStr);
			for (var innerIndex = innerIndex1; innerIndex<individualArrItem.length; innerIndex++) {
				if (innerIndex == 0) {	//question
					htmlString = htmlString + "<div class='question col-xs-12' id='mainquestion_"+questionIndex+"'>"+ questionIndex+". "+individualArrItem[innerIndex] + "</div>";
					questionAnsTracker = questionAnsTracker + questionIndex + "|";
					questionIndex = questionIndex + 1;
					//alert('qth');
				} else {
					if (individualArrItem[innerIndex].search("~") != -1) {
						if (individualArrItem[innerIndex].substring(0, individualArrItem[innerIndex].length-1) == "NA") {
							htmlString = htmlString + "<div class='col-md-12 actions_heading' style='display: none;'><span class='action_plan_text' style='display: none;'>"+subQtn+". </span><span id='subquestion_"+subQuestionIndex+"' class='action_plan_text' style='display: none;'>"+individualArrItem[innerIndex].substring(0, individualArrItem[innerIndex].length-1)+"</span></div>";
						} else {
							var subQtnFirst = parseInt(questionIndex-1);
							htmlString = htmlString + "<div class='col-md-12 actions_heading'><span class='action_plan_text'>"+subQtnFirst+"."+subQtnMarkers[subQtn-1]+". </span><span id='subquestion_"+subQuestionIndex+"' class='action_plan_text'>"+individualArrItem[innerIndex].substring(0, individualArrItem[innerIndex].length-1)+"</span></div>";
							subQtn = subQtn + 1;
						}
						questionAnsTracker = questionAnsTracker + subQuestionIndex + "~";
						subQuestionIndex = subQuestionIndex + 1;
						//alert('sub qth'+questionIndex);
						//subQtn = subQtn + 1;
					} else {
						htmlString = htmlString + "<div class='col-md-12 col-xs-12 single_answer'>";
						//htmlString = htmlString + "<div class='col-md-1 col-xs-1 answer_no'>"+answerTracker+".</div>";
						htmlString = htmlString + "<div class='col-md-11 col-xs-11 answer_input'><textarea class='col-md-12 col-xs-12 well answar_field' id='answer_"+answerTracker+"' maxlength='1500' onkeypress='return disableKey(event);' onpaste='return false' oncut='return false'>"+individualArrItem[innerIndex].trim()+"</textarea></div>";
						htmlString = htmlString + "<div class='clearfix'></div></div>";
						answerTracker = answerTracker + 1;
					}
					
					
				}
			countIndex = countIndex + 1;
			
			}
			questionAnsTracker = questionAnsTracker + "`";
		}
		document.getElementById("questionArea").innerHTML = htmlString;
	});
}
/**
 * function handles the next page 
 * population logic for questionnaire page.
 */
function saveQuestionnaireData(dist) {
	if(sessionStorage.getItem("isLogout")=="N"){
	if(	null == sessionStorage.getItem("isCompleted") || sessionStorage.getItem("isCompleted") == "" || sessionStorage.getItem("isCompleted") == 0){
		alertify.set({ buttonReverse: true });
		alertify.set({ labels: {
		    ok     : "Yes",
		    cancel : "No"
		} });
	alertify.confirm("<img src='../img/confirm-icon.png'><br /><p>Once you move to the next step, you will no longer be able to edit your Before Sketch.<br /> Is your Before Sketch finalized?</p>", function (e) {
	    if (e) {	
	    	$(".loader_bg").fadeIn();
	    	$(".loader").fadeIn();
	    	removePresentSessionItems();
	    	if (dist != "N") {
	    		sessionStorage.setItem("pageSequence", 1);
	    	}
	    	answerString = "";
	    	var textAnswerCounter = 1;
	    	var mainQuestionSplit = questionAnsTracker.split("`");
	    	for (var index = 0; index < mainQuestionSplit.length-1; index ++) {
	    		var mainQuestion = mainQuestionSplit[index].split("|");
	    		var question = document.getElementById("mainquestion_"+mainQuestion[0]).innerHTML;
	    		//fetching only the question
	    		question = question.substr(question.indexOf(".")+2, question.length-1);
	    		
	    		answerString = answerString + question + "@@@";
	    		var subQuestionInd = mainQuestion[1].split("~");
	    		for (var subQtn = 0; subQtn < subQuestionInd.length-1; subQtn ++) {
	    			var subQuestion = document.getElementById("subquestion_"+subQuestionInd[subQtn]).innerHTML;
	    			answerString = answerString + subQuestion + "~";
	    			for ( var answerIndex = 0; answerIndex < 1; answerIndex ++ ) {
	    				var myAnswer = document.getElementById("answer_"+textAnswerCounter).value;
	    				if (myAnswer.trim() == "") {
	    					myAnswer = " ";
	    				}
	    				answerString = answerString + myAnswer.replace(/#/g, "-_-").replace(/\(/g, ":=:").replace(/\)/g, ";=;").replace(/@/g, ":-:").replace(/\?/g, ";_;") + "`";
	    				textAnswerCounter = textAnswerCounter + 1;
	    			}
	    			answerString = answerString + "#";
	    		}
	    		answerString = answerString + ")(";
	    	}
	    	//console.log(answerString);
	    	//return;
	    	if (answerString != "") {
	    		
	    		var isCompleted = "Y";
	    		if(sessionStorage.getItem("isLogout")) {
	    			isCompleted = "Y";
	    		}
	    		//Create a model
	    		var qtrStr = Spine.Model.sub();
	    		qtrStr.configure("/user/questionnaire/saveAnswers", "myAns", "jobRefNo", "user", "profileId", "linkClicked", "isCompleted", "rowId", "isPrevious", "jctUserId");
	    		qtrStr.extend( Spine.Model.Ajax );

	    		//Populate the model with data to transder
	    		var qtrModel = new qtrStr({  
	    			myAns: answerString,
	    			jobRefNo: sessionStorage.getItem("jobReferenceNo"),
	    			user: sessionStorage.getItem("jctEmail"),
	    			profileId: sessionStorage.getItem("profileId"),
	    			linkClicked: sessionStorage.getItem("linkClicked"),
	    			isCompleted: isCompleted,
	    			rowId: sessionStorage.getItem("rowIdentity"),
	    			isPrevious: "N",
	    			jctUserId: sessionStorage.getItem("jctUserId")
	    		});
	    		qtrModel.save(); //POST
	    		
	    		qtrStr.bind("ajaxError", function(record, xhr, settings, error) {
	    			alertify.alert("Unable to connect to the server.");
	    			$(".loader_bg").fadeOut();
	    		    $(".loader").hide();
	    		});
	    		qtrStr.bind("ajaxSuccess", function(record, xhr, settings, error) {
	    			var jsonStr = JSON.stringify(xhr);
	    			var obj = jQuery.parseJSON(jsonStr);
	    			if(sessionStorage.getItem("isLogout") == "N") {
	    				sessionStorage.setItem("strength",JSON.stringify(obj.strengthMap));
	    				sessionStorage.setItem("passion",JSON.stringify(obj.passionMap));
	    				sessionStorage.setItem("value",JSON.stringify(obj.valueMap));
	    				if((null != obj.afterSketchCheckedEle) && (null != obj.afterSkPageOneTotalJson)
	    						&& ('null' != obj.afterSketchCheckedEle) && ('null' != obj.afterSkPageOneTotalJson)){
	    					sessionStorage.setItem("back_to_previous", 1);
	    					sessionStorage.setItem("checked_element", obj.afterSketchCheckedEle);
	    					sessionStorage.setItem("total_json_obj", obj.afterSkPageOneTotalJson);
	    					sessionStorage.setItem("strength_count", obj.strengthCount);
	    					sessionStorage.setItem("passion_count", obj.passionCount);
	    					sessionStorage.setItem("value_count", obj.valueCount);
	    					sessionStorage.setItem("checked_passion", obj.checkedPassion);
	    					sessionStorage.setItem("checked_strength", obj.checkedStrength);
	    					sessionStorage.setItem("checked_value", obj.checkedValue);
	    				}else{
	    					sessionStorage.setItem("strength_count", 1);
	    					sessionStorage.setItem("passion_count", 1);
	    					sessionStorage.setItem("value_count", 1);
	    				}
	    				sessionStorage.setItem("pageSequence", 3);
	    				sessionStorage.setItem("myPage", "AS");
	    				window.location = "afterSketch1.jsp";
	    			} else { //It is logout
	    				//Create a model
	    				var LogoutVO = Spine.Model.sub();
	    				LogoutVO.configure("/user/auth/logout", "jobReferenceString", "rowId", "lastActivePage");
	    				LogoutVO.extend( Spine.Model.Ajax );

	    				//Populate the model with data to transder
	    				var logoutModel = new LogoutVO({  
	    					jobReferenceString: sessionStorage.getItem("jobReferenceNo"),
	    					rowId: sessionStorage.getItem("rowIdentity"),
	    					lastActivePage: "/user/view/Questionaire.jsp"
	    				});
	    				logoutModel.save(); //POST
	    				
	    				LogoutVO.bind("ajaxError", function(record, xhr, settings, error) {
	    					$(".loader_bg").fadeOut();
	    				    $(".loader").hide();
	    					alertify.alert("Unable to connect to the server.");
	    				});
	    				
	    				LogoutVO.bind("ajaxSuccess", function(record, xhr, settings, error) {
	    					sessionStorage.clear();
	    					window.location = "../signup-page.jsp";
	    				});
	    			}
	    		});
	    		}
	    } 
	});	
	} else {		
    	$(".loader_bg").fadeIn();
    	$(".loader").fadeIn();
    	removePresentSessionItems();
    	if (dist != "N") {
    		sessionStorage.setItem("pageSequence", 1);
    	}
    	answerString = "";
    	var textAnswerCounter = 1;
    	var mainQuestionSplit = questionAnsTracker.split("`");
    	for (var index = 0; index < mainQuestionSplit.length-1; index ++) {
    		var mainQuestion = mainQuestionSplit[index].split("|");
    		var question = document.getElementById("mainquestion_"+mainQuestion[0]).innerHTML;
    		//fetching only the question
    		question = question.substr(question.indexOf(".")+2, question.length-1);
    		
    		answerString = answerString + question + "@@@";
    		var subQuestionInd = mainQuestion[1].split("~");
    		for (var subQtn = 0; subQtn < subQuestionInd.length-1; subQtn ++) {
    			var subQuestion = document.getElementById("subquestion_"+subQuestionInd[subQtn]).innerHTML;
    			answerString = answerString + subQuestion + "~";
    			for ( var answerIndex = 0; answerIndex < 1; answerIndex ++ ) {
    				var myAnswer = document.getElementById("answer_"+textAnswerCounter).value;
    				if (myAnswer.trim() == "") {
    					myAnswer = " ";
    				}
    				answerString = answerString + myAnswer.replace(/#/g, "-_-").replace(/\(/g, ":=:").replace(/\)/g, ";=;").replace(/@/g, ":-:").replace(/\?/g, ";_;") + "`";
    				textAnswerCounter = textAnswerCounter + 1;
    			}
    			answerString = answerString + "#";
    		}
    		answerString = answerString + ")(";
    	}
    	//console.log(answerString);
    	//return;
    	if (answerString != "") {
    		
    		var isCompleted = "Y";
    		if(sessionStorage.getItem("isLogout")) {
    			isCompleted = "Y";
    		}
    		//Create a model
    		var qtrStr = Spine.Model.sub();
    		qtrStr.configure("/user/questionnaire/saveAnswers", "myAns", "jobRefNo", "user", "profileId", "linkClicked", "isCompleted", "rowId", "isPrevious", "jctUserId");
    		qtrStr.extend( Spine.Model.Ajax );

    		//Populate the model with data to transder
    		var qtrModel = new qtrStr({  
    			myAns: answerString,
    			jobRefNo: sessionStorage.getItem("jobReferenceNo"),
    			user: sessionStorage.getItem("jctEmail"),
    			profileId: sessionStorage.getItem("profileId"),
    			linkClicked: sessionStorage.getItem("linkClicked"),
    			isCompleted: isCompleted,
    			rowId: sessionStorage.getItem("rowIdentity"),
    			isPrevious: "N",
    			jctUserId: sessionStorage.getItem("jctUserId")
    		});
    		qtrModel.save(); //POST
    		
    		qtrStr.bind("ajaxError", function(record, xhr, settings, error) {
    			alertify.alert("Unable to connect to the server.");
    			$(".loader_bg").fadeOut();
    		    $(".loader").hide();
    		});
    		qtrStr.bind("ajaxSuccess", function(record, xhr, settings, error) {
    			var jsonStr = JSON.stringify(xhr);
    			var obj = jQuery.parseJSON(jsonStr);
    			if(sessionStorage.getItem("isLogout") == "N") {
    				sessionStorage.setItem("strength",JSON.stringify(obj.strengthMap));
    				sessionStorage.setItem("passion",JSON.stringify(obj.passionMap));
    				sessionStorage.setItem("value",JSON.stringify(obj.valueMap));
    				if((null != obj.afterSketchCheckedEle) && (null != obj.afterSkPageOneTotalJson)
    						&& ('null' != obj.afterSketchCheckedEle) && ('null' != obj.afterSkPageOneTotalJson)){
    					sessionStorage.setItem("back_to_previous", 1);
    					sessionStorage.setItem("checked_element", obj.afterSketchCheckedEle);
    					sessionStorage.setItem("total_json_obj", obj.afterSkPageOneTotalJson);
    					sessionStorage.setItem("strength_count", obj.strengthCount);
    					sessionStorage.setItem("passion_count", obj.passionCount);
    					sessionStorage.setItem("value_count", obj.valueCount);
    					sessionStorage.setItem("checked_passion", obj.checkedPassion);
    					sessionStorage.setItem("checked_strength", obj.checkedStrength);
    					sessionStorage.setItem("checked_value", obj.checkedValue);
    				}else{
    					sessionStorage.setItem("strength_count", 1);
    					sessionStorage.setItem("passion_count", 1);
    					sessionStorage.setItem("value_count", 1);
    				}
    				sessionStorage.setItem("pageSequence", 3);
    				sessionStorage.setItem("myPage", "AS");
    				window.location = "afterSketch1.jsp";
    			} else { //It is logout
    				//Create a model
    				var LogoutVO = Spine.Model.sub();
    				LogoutVO.configure("/user/auth/logout", "jobReferenceString", "rowId", "lastActivePage");
    				LogoutVO.extend( Spine.Model.Ajax );

    				//Populate the model with data to transder
    				var logoutModel = new LogoutVO({  
    					jobReferenceString: sessionStorage.getItem("jobReferenceNo"),
    					rowId: sessionStorage.getItem("rowIdentity"),
    					lastActivePage: "/user/view/Questionaire.jsp"
    				});
    				logoutModel.save(); //POST
    				
    				LogoutVO.bind("ajaxError", function(record, xhr, settings, error) {
    					$(".loader_bg").fadeOut();
    				    $(".loader").hide();
    					alertify.alert("Unable to connect to the server.");
    				});
    				
    				LogoutVO.bind("ajaxSuccess", function(record, xhr, settings, error) {
    					sessionStorage.clear();
    					window.location = "../signup-page.jsp";
    				});
    			}
    		});
    		}    	
	}
} else {
	$(".loader_bg").fadeIn();
	$(".loader").fadeIn();
	removePresentSessionItems();
	if (dist != "N") {
		sessionStorage.setItem("pageSequence", 1);
	}
	answerString = "";
	var textAnswerCounter = 1;
	var mainQuestionSplit = questionAnsTracker.split("`");
	for (var index = 0; index < mainQuestionSplit.length-1; index ++) {
		var mainQuestion = mainQuestionSplit[index].split("|");
		var question = document.getElementById("mainquestion_"+mainQuestion[0]).innerHTML;
		//fetching only the question
		question = question.substr(question.indexOf(".")+2, question.length-1);
		
		answerString = answerString + question + "@@@";
		var subQuestionInd = mainQuestion[1].split("~");
		for (var subQtn = 0; subQtn < subQuestionInd.length-1; subQtn ++) {
			var subQuestion = document.getElementById("subquestion_"+subQuestionInd[subQtn]).innerHTML;
			answerString = answerString + subQuestion + "~";
			for ( var answerIndex = 0; answerIndex < 1; answerIndex ++ ) {
				var myAnswer = document.getElementById("answer_"+textAnswerCounter).value;
				if (myAnswer.trim() == "") {
					myAnswer = " ";
				}
				answerString = answerString + myAnswer.replace(/#/g, "-_-").replace(/\(/g, ":=:").replace(/\)/g, ";=;").replace(/@/g, ":-:").replace(/\?/g, ";_;") + "`";
				textAnswerCounter = textAnswerCounter + 1;
			}
			answerString = answerString + "#";
		}
		answerString = answerString + ")(";
	}
	//console.log(answerString);
	//return;
	if (answerString != "") {
		//Create a model
		var qtrStr = Spine.Model.sub();
		qtrStr.configure("/user/questionnaire/saveAnswers", "myAns", "jobRefNo", "user", "profileId", "linkClicked", "isCompleted", "rowId", "isPrevious", "jctUserId");
		qtrStr.extend( Spine.Model.Ajax );

		//Populate the model with data to transder
		var qtrModel = new qtrStr({  
			myAns: answerString,
			jobRefNo: sessionStorage.getItem("jobReferenceNo"),
			user: sessionStorage.getItem("jctEmail"),
			profileId: sessionStorage.getItem("profileId"),
			linkClicked: sessionStorage.getItem("linkClicked"),
			isCompleted: sessionStorage.getItem("isCompleted"),
			rowId: sessionStorage.getItem("rowIdentity"),
			isPrevious: "N",
			jctUserId: sessionStorage.getItem("jctUserId")
		});
		qtrModel.save(); //POST
		
		qtrStr.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		});
		qtrStr.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			if(sessionStorage.getItem("isLogout") == "N") {
				sessionStorage.setItem("strength",JSON.stringify(obj.strengthMap));
				sessionStorage.setItem("passion",JSON.stringify(obj.passionMap));
				sessionStorage.setItem("value",JSON.stringify(obj.valueMap));
				if((null != obj.afterSketchCheckedEle) && (null != obj.afterSkPageOneTotalJson)
						&& ('null' != obj.afterSketchCheckedEle) && ('null' != obj.afterSkPageOneTotalJson)){
					sessionStorage.setItem("back_to_previous", 1);
					sessionStorage.setItem("checked_element", obj.afterSketchCheckedEle);
					sessionStorage.setItem("total_json_obj", obj.afterSkPageOneTotalJson);
					sessionStorage.setItem("strength_count", obj.strengthCount);
					sessionStorage.setItem("passion_count", obj.passionCount);
					sessionStorage.setItem("value_count", obj.valueCount);
					sessionStorage.setItem("checked_passion", obj.checkedPassion);
					sessionStorage.setItem("checked_strength", obj.checkedStrength);
					sessionStorage.setItem("checked_value", obj.checkedValue);
				}else{
					sessionStorage.setItem("strength_count", 1);
					sessionStorage.setItem("passion_count", 1);
					sessionStorage.setItem("value_count", 1);
				}
				sessionStorage.setItem("pageSequence", 3);
				sessionStorage.setItem("myPage", "AS");
				window.location = "afterSketch1.jsp";
			} else { //It is logout
				//Create a model
				var LogoutVO = Spine.Model.sub();
				LogoutVO.configure("/user/auth/logout", "jobReferenceString", "rowId", "lastActivePage");
				LogoutVO.extend( Spine.Model.Ajax );

				//Populate the model with data to transder
				var logoutModel = new LogoutVO({  
					jobReferenceString: sessionStorage.getItem("jobReferenceNo"),
					rowId: sessionStorage.getItem("rowIdentity"),
					lastActivePage: "/user/view/Questionaire.jsp"
				});
				logoutModel.save(); //POST
				
				LogoutVO.bind("ajaxError", function(record, xhr, settings, error) {
					$(".loader_bg").fadeOut();
				    $(".loader").hide();
					alertify.alert("Unable to connect to the server.");
				});
				
				LogoutVO.bind("ajaxSuccess", function(record, xhr, settings, error) {
					sessionStorage.clear();
					window.location = "../signup-page.jsp";
				});
			}
		});
		}
	}
	sessionStorage.setItem("qsView","H");
}

function navigateToPrevious() {
	$(".loader_bg").fadeIn();
	$(".loader").fadeIn();
	saveAppDataOnPrevious();
	removePresentSessionItems();
	var navigation = Spine.Model.sub();
	navigation.configure("/user/navigation/goPrevious", "navigationSeq", "jobReferenceNos" , "profileId", "linkClicked", "rowId");
	navigation.extend( Spine.Model.Ajax );
	
	//Populate the model
	var replyModel = new navigation({
		navigationSeq: "1",
		jobReferenceNos: sessionStorage.getItem("jobReferenceNo"),
		profileId: sessionStorage.getItem("profileId"),
		linkClicked: sessionStorage.getItem("linkClicked"),
		rowId: sessionStorage.getItem("rowIdentity")
	});
	replyModel.save(); //POST
	
	navigation.bind("ajaxError", function(record, xhr, settings, error) {
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
		alertify.alert("Unable to connect to the server.");
		return false;
	});
	
	navigation.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		if(sessionStorage.getItem("isLogout") == "N"){
			sessionStorage.setItem("total_json_add_task", obj.bsJson); //Job reference
			var myJsonObj = JSON.parse(obj.bsJson);
			sessionStorage.setItem("total_count", myJsonObj.length+1);
			sessionStorage.setItem("div_intial_value", obj.initialJson);
			sessionStorage.setItem("isCompleted", obj.isCompleted);
			sessionStorage.setItem("pageSequence", 1);
			sessionStorage.setItem("myPage", "BS");
			window.location = "beforeSketch.jsp";
		}
	});
	sessionStorage.setItem("qsView","H");
}
/**
 * Function saves data when previous link is clicked
 */
function saveAppDataOnPrevious(){
	answerString = "";
	var textAnswerCounter = 1;
	var mainQuestionSplit = questionAnsTracker.split("`");
	for (var index = 0; index < mainQuestionSplit.length-1; index ++) {
		var mainQuestion = mainQuestionSplit[index].split("|");
		var question = document.getElementById("mainquestion_"+mainQuestion[0]).innerHTML;
		
		//fetching only the question
		question = question.substr(question.indexOf(".")+2, question.length-1);
		
		answerString = answerString + question + "@@@";
		var subQuestionInd = mainQuestion[1].split("~");
		for (var subQtn = 0; subQtn < subQuestionInd.length-1; subQtn ++) {
			var subQuestion = document.getElementById("subquestion_"+subQuestionInd[subQtn]).innerHTML;
			answerString = answerString + subQuestion + "~";
			for ( var answerIndex = 0; answerIndex < 1; answerIndex ++ ) {
				var myAnswer = document.getElementById("answer_"+textAnswerCounter).value;
				if (myAnswer.trim() == "") {
					myAnswer = " ";
				}
				answerString = answerString + myAnswer.replace(/#/g, "-_-").replace(/\(/g, ":=:").replace(/\)/g, ";=;").replace(/@/g, ":-:").replace(/\?/g, ";_;") + "`";
				textAnswerCounter = textAnswerCounter + 1;
			}
			answerString = answerString + "#";
		}
		answerString = answerString + ")(";
	}
	
	var qtrStr = Spine.Model.sub();
	qtrStr.configure("/user/questionnaire/saveAnswers", "myAns", "jobRefNo", "user", "profileId", "linkClicked", "isCompleted", "rowId", "isPrevious", "jctUserId");
	qtrStr.extend( Spine.Model.Ajax );

	//Populate the model with data to transder
	var qtrModel = new qtrStr({  
		myAns: answerString,
		jobRefNo: sessionStorage.getItem("jobReferenceNo"),
		user: sessionStorage.getItem("jctEmail"),
		profileId: sessionStorage.getItem("profileId"),
		linkClicked: sessionStorage.getItem("linkClicked"),
		isCompleted: "N",
		rowId: sessionStorage.getItem("rowIdentity"),
		isPrevious: "Y",
		jctUserId: sessionStorage.getItem("jctUserId")
	});
	qtrModel.save(); //POST
}
/**
 * Function add to restrict enter as a input
 * @param evt
 */
function disableKey(evt) {
    var charCode = (evt.which) ? evt.which : event.keyCode;
    if (charCode == 13)
        return false;
    if (charCode == 126) //~
		return false;
    if (charCode == 96) //`
		return false;
    if (charCode > 126) 
    	return false;
    return true;
}
/**
 * Function removes the session items which will not be used
 */
function removePresentSessionItems() {
	sessionStorage.removeItem("questions");
	sessionStorage.removeItem("snapShotURLS");
	if (sessionStorage.getItem("passion")) {
		sessionStorage.removeItem("passion");
	}
	if (sessionStorage.getItem("passion_count")) {
		sessionStorage.removeItem("passion_count");
	}
	if (sessionStorage.getItem("strength")) {
		sessionStorage.removeItem("strength");
	}
	if (sessionStorage.getItem("strength_count")) {
		sessionStorage.removeItem("strength_count");
	}
	if (sessionStorage.getItem("total_json_after_sketch_final")) {
		sessionStorage.removeItem("total_json_after_sketch_final");
	}
	if (sessionStorage.getItem("value")) {
		sessionStorage.removeItem("value");
	}
	if (sessionStorage.getItem("value_count")) {
		sessionStorage.removeItem("value_count");
	}
	
}