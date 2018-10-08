/************************************************************************
 * Description: The js file is specifically for Action Plan page.		*
 * 																		*
 * Functions  : goPrevious() 											*
 * 					--> This function will help to 						*
 * 						navigate to previous page.						*
 * 						navigate to next page.							*
 * 				populateActionPlan()									*
 * 					--> This will help to populate the action plan  	*
 * 						when the page loads up.							*
 * 				saveActionPlan()										*
 * 					--> This function will save data to database and	*
 * 						navigate to the next page. This internally 		*
 * 						calls 2 other methods depending upon the nature	*
 * 																		*
 * Author	  : InterraIT										*
 * 																		*
 * 																		*
 ***********************************************************************/
var questionIndex = 1;
var subQuestionIndex = 1;
var answerTracker = 1;
var questionAnsTracker = "";
var answerString = "";

window.setInterval(function(){
	updateTime();
}, 60000);

/**
 * Function add to disable browser back button
 * while the page is loaded
 * @param null
 */
window.location.hash="";
window.location.hash="Again-No-back-button";//again because google chrome don't insert first hash into history
window.onhashchange=function(){window.location.hash="";};


$(document).ready(function() {
	$(".loader_bg").fadeIn();
	$(".loader").fadeIn();
	sessionStorage.setItem("pageSequence", 5);
	sessionStorage.setItem("pdfRequired", "Y");
	populateActionPlan();
	/***************** To Disable the paste functionality ISSUE# 186***************/
    $(".action_plan_field").bind("paste",function(e) {
        e.preventDefault();
    });
});


/**
 * Function populates action plan on document ready
 * @param null
 */
function populateActionPlan(){
	var navigation = Spine.Model.sub();
	navigation.configure("/user/actionPlan/fetchActionPlan", "jobReferenceNo", "profileId");
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
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var htmlString = "";
		for (var populationIndex=0; populationIndex<obj.actionPlanList.length; populationIndex++) {
			var individualArrItem = obj.actionPlanList[populationIndex];
			var countIndex = 1;
			var innerIndex1 = 0;
			var subQtn = 1;
			var subQtnMarkers = ["a", "b", "c", "d","a", "b", "c", "d","a", "b", "c", "d","a", "b", "c", "d"];
			for (var innerIndex = innerIndex1; innerIndex<individualArrItem.length; innerIndex++) {
				if (innerIndex == 0) {	//question
					htmlString = htmlString + "<div class='question col-xs-12' id='mainquestion_"+questionIndex+"'>"+ questionIndex+". " +individualArrItem[innerIndex] + "</div>";
					questionAnsTracker = questionAnsTracker + questionIndex + "|";
					questionIndex = questionIndex + 1;
				} else {
					if (individualArrItem[innerIndex].search("~") != -1) {
						if (individualArrItem[innerIndex].substring(0, individualArrItem[innerIndex].length-1) == "NA") {
							htmlString = htmlString + "<div class='col-md-12 actions_heading' style='display: none;'><span id='subquestion_"+subQuestionIndex+"' class='action_plan_text'>"+individualArrItem[innerIndex].substring(0, individualArrItem[innerIndex].length-1)+"</span></div>";
						} else {
							var subQtnFirst = parseInt(questionIndex-1);
							//alert(subQtnFirst);
							//htmlString = htmlString + "<div class='col-md-12 actions_heading'><span id='subquestion_"+subQuestionIndex+"' class='action_plan_text'>"+individualArrItem[innerIndex].substring(0, individualArrItem[innerIndex].length-1)+"</span></div>";
							htmlString = htmlString + "<div class='col-md-12 actions_heading'><span class='action_plan_text'>"+subQtnFirst+"."+subQtnMarkers[subQtn-1]+". </span><span id='subquestion_"+subQuestionIndex+"' class='action_plan_text'>"+individualArrItem[innerIndex].substring(0, individualArrItem[innerIndex].length-1)+"</span></div>";
							subQtn = subQtn + 1;
						}
						questionAnsTracker = questionAnsTracker + subQuestionIndex + "~";
						subQuestionIndex = subQuestionIndex + 1;
					} else {
						htmlString = htmlString + "<div class='col-md-12 col-xs-12 single_answer'>";
						//htmlString = htmlString + "<div class='col-md-1 col-xs-1 answer_no'>"+answerTracker+".</div>";
						htmlString = htmlString + "<div class='col-md-11 col-xs-11 answer_input'><textarea class='col-md-12 col-xs-12 well_action action_plan_field' id='answer_"+answerTracker+"' maxlength='1500' onkeypress='return disableKey(event);' onpaste='return false' oncut='return false'>"+individualArrItem[innerIndex].trim()+"</textarea></div>";
						htmlString = htmlString + "<div class='clearfix'></div></div>";
						answerTracker = answerTracker + 1;
					}
				}
			countIndex = countIndex + 1;
			
			}
			questionAnsTracker = questionAnsTracker + "`";
		}
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
		document.getElementById("populationAreaId").innerHTML = htmlString;
	});
}

/**
 * Function save action plan data on next
 * @param dist
 */
function saveActionPlan(dist) {
	$(".loader_bg").fadeIn();
	$(".loader").fadeIn();
	removePresentSessionItems();
	if (dist != "N") {
		sessionStorage.setItem("pageSequence", 4);
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
			for ( var answerIndex = 0; answerIndex < 3; answerIndex ++ ) {
				var myAnswer = document.getElementById("answer_"+textAnswerCounter).value;
				if(sessionStorage.getItem("src") != "Admin") {
					if (dist == "N") {
						if (myAnswer.trim() == "") {
							$(".loader_bg").fadeOut();
						    $(".loader").hide();
							alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please complete all parts of the Action Plan.</p>");
							return false;
						}
					} else {
						if (myAnswer.trim() == "") {
							myAnswer = " ";
						}
					}
				} else {
					if (myAnswer.trim() == "") {
						myAnswer = " ";
					}
				}
				answerString = answerString + myAnswer.replace(/#/g, "-_-").replace(/\(/g, ":=:").replace(/\)/g, ";=;").replace(/@/g, ":-:").replace(/\?/g, ";_;") + "`";
				textAnswerCounter = textAnswerCounter + 1;
			}
			answerString = answerString + "#";
		}
		answerString = answerString + ")(";
	}
	if (answerString != "") {
		
	
	//Create a model
	var actionPlanStr = Spine.Model.sub();
	actionPlanStr.configure("/user/actionPlan/saveActionPlan", "myPlan", "jobRefNo", "user", 
			"linkClicked", "pdfReqs", "rowId", "jctUserId");
	actionPlanStr.extend( Spine.Model.Ajax );

	//Populate the model with data to transder
	var actionPlanModel = new actionPlanStr({  
		myPlan: answerString,
		jobRefNo: sessionStorage.getItem("jobReferenceNo"),
		user: sessionStorage.getItem("jctEmail"),
		linkClicked: sessionStorage.getItem("linkClicked"),
		pdfReqs: sessionStorage.getItem("pdfRequired"),
		rowId: sessionStorage.getItem("rowIdentity"),
		jctUserId: sessionStorage.getItem("jctUserId")
	});
	actionPlanModel.save(); //POST
	
	actionPlanStr.bind("ajaxError", function(record, xhr, settings, error) {
		//alertify.alert("Unable to connect to the server.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	});
	actionPlanStr.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		if(sessionStorage.getItem("isLogout") == "N") {
			/*
			if (sessionStorage.getItem("fromPage") && (dist == "N")) {
				sessionStorage.setItem('open','open');
				sessionStorage.setItem("pageSequence", 6);
				window.location = "resultPageFinal.jsp";
			} else if (!sessionStorage.getItem("fromPage") && (dist == "N")) {
				sessionStorage.setItem('open','open');
				sessionStorage.setItem("pageSequence", 6);
				window.location = "resultPage.jsp";
			} else*/ 
			
			if (!sessionStorage.getItem("fromPage") && (dist != "N")){
				goPrevious();
			} else {
				sessionStorage.setItem('open','open');
				sessionStorage.setItem("pageSequence", 6);
				if (obj.demographicPopupMsg) {
					sessionStorage.setItem("demoInfoMsg", obj.demographicPopupMsg);
				}
				window.location = obj.url;
			}
		} else { //It is logout
			//Create a model
			var LogoutVO = Spine.Model.sub();
			LogoutVO.configure("/user/auth/logout", "jobReferenceString", "rowId", "lastActivePage");
			LogoutVO.extend( Spine.Model.Ajax );
			//Populate the model with data to transder
			var logoutModel = new LogoutVO({  
				jobReferenceString: sessionStorage.getItem("jobReferenceNo"),
				rowId: sessionStorage.getItem("rowIdentity"),
				lastActivePage: "/user/view/actionPlan.jsp"
			});
			logoutModel.save(); //POST			
			LogoutVO.bind("ajaxError", function(record, xhr, settings, error) {
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
				//alertify.alert("Unable to connect to the server.");
			});			
			LogoutVO.bind("ajaxSuccess", function(record, xhr, settings, error) {
				sessionStorage.clear();
				window.location = "../signup-page.jsp";
			});
		}
	});
	}
}

/**
 * Function helps to navigate to previous page
 * @param null
 */
function goPrevious() {
	sessionStorage.setItem("pdfRequired", "N");
	sessionStorage.removeItem("fromPage");
	if(saveActionPlanforPrevious() == true){
	//sessionStorage.removeItem("prevClk");
	$(".loader_bg").fadeIn();
	$(".loader").fadeIn();
	//sessionStorage.setItem("pageSequence", 4);
	var navigation = Spine.Model.sub();
	navigation.configure("/user/navigation/goPrevious", "navigationSeq", "jobReferenceNos", "profileId", "linkClicked", "rowId");
	navigation.extend( Spine.Model.Ajax );
	//Populate the model
	var replyModel = new navigation({
		navigationSeq: "4",
		jobReferenceNos: sessionStorage.getItem("jobReferenceNo"),
		profileId: sessionStorage.getItem("profileId"),
		linkClicked: sessionStorage.getItem("linkClicked"),
		rowId: sessionStorage.getItem("rowIdentity")
	});
	replyModel.save(); //POST
	
	navigation.bind("ajaxError", function(record, xhr, settings, error) {
		//alertify.alert("Unable to connect to the server.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
		return false;
	});
	
	navigation.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		console.log(obj);
		if(sessionStorage.getItem("isLogout") == "N"){
			sessionStorage.setItem("total_json_obj", obj.pageOneJson);
			sessionStorage.setItem("div_intial_value", obj.divInitialValue);
			sessionStorage.setItem("total_json_after_sketch_final", obj.afterSketch2TotalJsonObj);
			sessionStorage.setItem("total_count", JSON.parse(obj.totalJsonAddTask).length+1);
			sessionStorage.setItem("divHeight", obj.divHeight);
			sessionStorage.setItem("divWidth", obj.divWidth);
			sessionStorage.setItem("isCompleted", obj.isCompleted);
			sessionStorage.setItem("total_count", obj.asTotalCount);
			sessionStorage.removeItem("prevClk");
			window.location = "afterSketch2.jsp";
		}
	});
	}
}

/**
 * Function saves Action Plan when going to previous page
 * @param null
 */
function saveActionPlanforPrevious(){
	$(".loader_bg").fadeIn();
	$(".loader").fadeIn();
	if (!sessionStorage.getItem("prevClk")) {
		sessionStorage.setItem("prevClk","Y");
		saveActionPlan("P");
		
	}
	return true;
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
	sessionStorage.removeItem("snapShotURLS");
}