$(document).ready(function() {
	sessionStorage.setItem("myPage", "AS");
	sessionStorage.setItem("pageSequence", 8);
	if (sessionStorage.getItem("bsEdit")) {
		sessionStorage.removeItem("bsEdit");
	}
	if (sessionStorage.getItem("editLink")) {
		sessionStorage.removeItem('editLink');
	}
	sessionStorage.setItem("giveUserOption", "Y");
});

function goToResultPage() {
	sessionStorage.setItem("myPage", "AS");
	sessionStorage.setItem("pageSequence", 8);
	if (sessionStorage.getItem("bsEdit")) {
		sessionStorage.removeItem("bsEdit");
	}
	if (sessionStorage.getItem("editLink")) {
		sessionStorage.removeItem('editLink');
	}
	sessionStorage.setItem("giveUserOption", "Y");
	sessionStorage.setItem("pageSequence", 6);
	sessionStorage.removeItem('open'); // otherwise this will popup the pdf again
	window.location = "resultPage.jsp";
}
/**
 * function populates the landing page and also fetches the 
 * before sketch page data and stores in session storage.
 */
function populateLandingPage() {
	sessionStorage.setItem("myPage", "AS");
	sessionStorage.setItem("pageSequence", 8);
	if (sessionStorage.getItem("bsEdit")) {
		sessionStorage.removeItem("bsEdit");
	}
	if (sessionStorage.getItem("editLink")) {
		sessionStorage.removeItem('editLink');
	}
	if (sessionStorage.getItem("infoDispReq")) {
		sessionStorage.removeItem("infoDispReq");
	}
	sessionStorage.setItem("giveUserOption", "Y");
	$(".loader_bg").fadeIn();
	$(".loader").fadeIn();
	var navigation = Spine.Model.sub();
	navigation.configure("/user/navigation/goPrevious", "navigationSeq", "jobReferenceNos" , "profileId", "linkClicked", "rowId");
	navigation.extend( Spine.Model.Ajax );
	
	//Populate the model
	var replyModel = new navigation({
		navigationSeq: "1",
		jobReferenceNos: sessionStorage.getItem("jobReferenceNo"),
		profileId: sessionStorage.getItem("profileId"),
		linkClicked: "Y",
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
		if(sessionStorage.getItem("isLogout") == "N") {
			sessionStorage.setItem("linkClicked", "Y");
			sessionStorage.setItem("total_json_add_task", obj.bsJson); //Job reference
			var myJsonObj = JSON.parse(obj.bsJson);
			sessionStorage.setItem("total_count", myJsonObj.length+1);
			sessionStorage.setItem("div_intial_value", obj.initialJson);
			sessionStorage.setItem("isCompleted", 1);
			sessionStorage.setItem("pageSequence", 1);
			sessionStorage.setItem("myPage", "BS");
			sessionStorage.setItem("fromPage","final");
			sessionStorage.setItem("url", "/user/view/beforeSketch.jsp");
			saveNewActivePage("/user/view/beforeSketch.jsp");
			window.location = "landing-page.jsp";
		}
	});
}
/**
 * function fetches the before sketch page data and stores in session storage.
 */
function populateBeforeSketch() {
	sessionStorage.setItem("myPage", "AS");
	sessionStorage.setItem("pageSequence", 8);
	if (sessionStorage.getItem("bsEdit")) {
		sessionStorage.removeItem("bsEdit");
	}
	if (sessionStorage.getItem("editLink")) {
		sessionStorage.removeItem('editLink');
	}
	if (sessionStorage.getItem("infoDispReq")) {
		sessionStorage.removeItem("infoDispReq");
	}
	sessionStorage.setItem("giveUserOption", "Y");
	$(".loader_bg").fadeIn();
	$(".loader").fadeIn();
	var navigation = Spine.Model.sub();
	navigation.configure("/user/navigation/goPrevious", "navigationSeq", "jobReferenceNos" , "profileId", "linkClicked", "rowId");
	navigation.extend( Spine.Model.Ajax );
	
	//Populate the model
	var replyModel = new navigation({
		navigationSeq: "1",
		jobReferenceNos: sessionStorage.getItem("jobReferenceNo"),
		profileId: sessionStorage.getItem("profileId"),
		linkClicked: "Y",
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
		if(sessionStorage.getItem("isLogout") == "N") {
			sessionStorage.setItem("linkClicked", "Y");
			sessionStorage.setItem("total_json_add_task", obj.bsJson); //Job reference
			var myJsonObj = JSON.parse(obj.bsJson);
			sessionStorage.setItem("total_count", myJsonObj.length+1);
			sessionStorage.setItem("div_intial_value", obj.initialJson);
			sessionStorage.setItem("isCompleted", 1);
			sessionStorage.setItem("pageSequence", 1);
			sessionStorage.setItem("myPage", "BS");
			sessionStorage.setItem("fromPage","final");
			saveNewActivePage("/user/view/beforeSketch.jsp");
			window.location = "beforeSketch.jsp";
		}
	});
}
/**
 * function fetches the questionnaire page data and stores in session storage.
 */
function populateQuestionnaire() {
	sessionStorage.setItem("myPage", "AS");
	sessionStorage.setItem("pageSequence", 8);
	if (sessionStorage.getItem("bsEdit")) {
		sessionStorage.removeItem("bsEdit");
	}
	if (sessionStorage.getItem("editLink")) {
		sessionStorage.removeItem('editLink');
	}
	if (sessionStorage.getItem("infoDispReq")) {
		sessionStorage.removeItem("infoDispReq");
	}
	sessionStorage.setItem("giveUserOption", "Y");
	$(".loader_bg").fadeIn();
	$(".loader").fadeIn();
	//Fetch the questionaire
	var navigation = Spine.Model.sub();
	navigation.configure("/user/navigation/goPrevious", "navigationSeq", "jobReferenceNos", "profileId", "linkClicked", "rowId");
	navigation.extend( Spine.Model.Ajax );
	//Populate the model
		var replyModel = new navigation({
			navigationSeq: "2",
			jobReferenceNos: sessionStorage.getItem("jobReferenceNo"),
			profileId: sessionStorage.getItem("profileId"),
			linkClicked: "Y",
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
			if(sessionStorage.getItem("isLogout") == "N") {
				sessionStorage.setItem("linkClicked", "Y");
				sessionStorage.setItem("questions", obj.list);
				sessionStorage.setItem("snapShotURLS", obj.jctBaseString);
				sessionStorage.setItem("isCompleted", 1);
				sessionStorage.setItem("pageSequence", 2);
				sessionStorage.setItem("fromPage","final");
				saveNewActivePage("/user/view/Questionaire.jsp");
				window.location = "Questionaire.jsp";
			}
		});
}
/**
 * function fetches the mapping page data and stores in session storage.
 */
function populateMapping() {
	sessionStorage.setItem("myPage", "AS");
	sessionStorage.setItem("pageSequence", 8);
	if (sessionStorage.getItem("bsEdit")) {
		sessionStorage.removeItem("bsEdit");
	}
	if (sessionStorage.getItem("editLink")) {
		sessionStorage.removeItem('editLink');
	}
	if (sessionStorage.getItem("infoDispReq")) {
		sessionStorage.removeItem("infoDispReq");
	}
	sessionStorage.setItem("giveUserOption", "Y");
	$(".loader_bg").fadeIn();
	$(".loader").fadeIn();
	var navigation = Spine.Model.sub();
	navigation.configure("/user/navigation/goPrevious", "navigationSeq", "jobReferenceNos", "profileId", "linkClicked", "rowId");
	navigation.extend( Spine.Model.Ajax );
	//Populate the model
	var replyModel = new navigation({
		navigationSeq: "3",
		jobReferenceNos: sessionStorage.getItem("jobReferenceNo"),
		profileId: sessionStorage.getItem("profileId"),
		linkClicked: "Y",
		rowId: sessionStorage.getItem("rowIdentity")
	});
	replyModel.save(); //POST
	
	navigation.bind("ajaxError", function(record, xhr, settings, error) {
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
		alertify.alert("Unable to connect to the server!");
		return false;
	});
	
	navigation.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		if(sessionStorage.getItem("isLogout") == "N") {
			sessionStorage.setItem("linkClicked", "Y");
			sessionStorage.setItem("back_to_previous", 1);
			sessionStorage.setItem("strength",JSON.stringify(obj.strengthMap));
			sessionStorage.setItem("passion",JSON.stringify(obj.passionMap));
			sessionStorage.setItem("value",JSON.stringify(obj.valueMap));
			sessionStorage.setItem("checked_element", obj.afterSketchCheckedEle);
			sessionStorage.setItem("total_json_obj", obj.afterSkPageOneTotalJson);
			sessionStorage.setItem("strength_count", obj.strengthCount);
			sessionStorage.setItem("passion_count", obj.passionCount);
			sessionStorage.setItem("value_count", obj.valueCount);
			sessionStorage.setItem("checked_passion", obj.checkedPassion);
			sessionStorage.setItem("checked_strength", obj.checkedStrength);
			sessionStorage.setItem("checked_value", obj.checkedValue);
			sessionStorage.setItem("isCompleted", 1);
			sessionStorage.setItem("pageSequence", 3);
			sessionStorage.setItem("fromPage","final");
			saveNewActivePage("/user/view/afterSketch1.jsp");
			window.location = "afterSketch1.jsp";
		}
	});
}
/**
 * function fetches the putting it together page page data and stores in session storage.
 */
function populateAfterDiag() {
	sessionStorage.setItem("myPage", "AS");
	sessionStorage.setItem("pageSequence", 8);
	if (sessionStorage.getItem("bsEdit")) {
		sessionStorage.removeItem("bsEdit");
	}
	if (sessionStorage.getItem("editLink")) {
		sessionStorage.removeItem('editLink');
	}
	if (sessionStorage.getItem("infoDispReq")) {
		sessionStorage.removeItem("infoDispReq");
	}
	sessionStorage.setItem("giveUserOption", "Y");
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
		linkClicked: "Y",
		rowId: sessionStorage.getItem("rowIdentity")
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
		if(sessionStorage.getItem("isLogout") == "N") {
			sessionStorage.setItem("linkClicked", "Y");
			sessionStorage.setItem("total_json_obj", obj.pageOneJson);
			sessionStorage.setItem("div_intial_value", obj.divInitialValue);
			sessionStorage.setItem("total_json_after_sketch_final", obj.afterSketch2TotalJsonObj);
			sessionStorage.setItem("total_count", JSON.parse(obj.totalJsonAddTask).length+1);
			sessionStorage.setItem("divHeight", obj.divHeight);
			sessionStorage.setItem("divWidth", obj.divWidth);
			sessionStorage.setItem("total_count", obj.asTotalCount);
			sessionStorage.setItem("isCompleted", 1);
			sessionStorage.setItem("pageSequence", 4);
			sessionStorage.setItem("fromPage","final");
			saveNewActivePage("/user/view/afterSketch2.jsp");
			window.location = "afterSketch2.jsp";
		}
	});
}
/**
 * function fetches the action plan page data and stores in session storage.
 */
function populateActionPlan() {
	sessionStorage.setItem("myPage", "AS");
	sessionStorage.setItem("pageSequence", 8);
	if (sessionStorage.getItem("bsEdit")) {
		sessionStorage.removeItem("bsEdit");
	}
	if (sessionStorage.getItem("editLink")) {
		sessionStorage.removeItem('editLink');
	}
	if (sessionStorage.getItem("infoDispReq")) {
		sessionStorage.removeItem("infoDispReq");
	}
	sessionStorage.setItem("giveUserOption", "Y");
	$(".loader_bg").fadeIn();
	$(".loader").fadeIn();
	var navigation = Spine.Model.sub();
	navigation.configure("/user/navigation/goPrevious", "navigationSeq", "jobReferenceNos" ,"profileId", "linkClicked", "rowId");
	navigation.extend( Spine.Model.Ajax );
	//Populate the model
	var replyModel = new navigation({
		navigationSeq: "5",
		jobReferenceNos: sessionStorage.getItem("jobReferenceNo"),
		profileId: sessionStorage.getItem("profileId"),
		linkClicked: "Y",
		rowId: sessionStorage.getItem("rowIdentity")
	});
	replyModel.save(); //POST
	
	navigation.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	});
	
	navigation.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		if(sessionStorage.getItem("isLogout") == "N") {
			sessionStorage.setItem("linkClicked", "Y");
			sessionStorage.setItem("pageSequence", 5);
			sessionStorage.setItem("snapShotURLS", obj.snapShot);
			sessionStorage.setItem("isCompleted", 1);
			sessionStorage.setItem("fromPage","final");
			saveNewActivePage("/user/view/actionPlan.jsp");
			window.location = "actionPlan.jsp";
		}
	});
}

/**
 * Function helps in restarting the link
 */
function restartExcersize() {
	sessionStorage.setItem("myPage", "AS");
	sessionStorage.setItem("pageSequence", 8);
	if (sessionStorage.getItem("bsEdit")) {
		sessionStorage.removeItem("bsEdit");
	}
	if (sessionStorage.getItem("editLink")) {
		sessionStorage.removeItem('editLink');
	}
	if (sessionStorage.getItem("infoDispReq")) {
		sessionStorage.removeItem("infoDispReq");
	}
	sessionStorage.setItem("giveUserOption", "Y");
	alertify.set({ buttonReverse: true });
	alertify.set({ labels: {
	    ok     : "Yes",
	    cancel : "No"
	} });
	alertify.confirm("<img src='../img/confirm-icon.png'><br /><p>Are you sure to restart the exercise? <br /><b>Note: </b>This will delete your current Before Sketch and After Diagram.</p>", function (e) {
	    if (e) {
	    	$(".loader_bg").fadeIn();
	    	$(".loader").fadeIn();
	    	var navigation = Spine.Model.sub();
	    	navigation.configure("/user/navigation/restartExcersize", "jobReferenceNo", "emailId", "distinction");
	    	navigation.extend( Spine.Model.Ajax );
	    	//Populate the model
	    	var replyModel = new navigation({
	    		jobReferenceNo: sessionStorage.getItem("jobReferenceNo"),
	    		emailId: sessionStorage.getItem("jctEmail"),
	    		distinction: "N"
	    	});
	    	replyModel.save(); //POST
	    	
	    	navigation.bind("ajaxError", function(record, xhr, settings, error) {
	    		alertify.alert("Unable to connect to the server.");
	    		$(".loader_bg").fadeOut();
	    	    $(".loader").hide();
	    	});
	    	
	    	navigation.bind("ajaxSuccess", function(record, xhr, settings, error) {
	    		var jsonStr = JSON.stringify(xhr);
	    		var obj = jQuery.parseJSON(jsonStr);
	    		if (obj.statusCode == "200") {
	    			if (sessionStorage.getItem("total_json_add_task")) {
	    				sessionStorage.removeItem("total_json_add_task");
	    			}
	    			if (sessionStorage.getItem("total_count")) {
	    				sessionStorage.removeItem("total_count");
	    			}
	    			if (sessionStorage.getItem("div_intial_value")) {
	    				sessionStorage.removeItem("div_intial_value");
	    			}
	    			sessionStorage.setItem("next_page", obj.lastPage);
	    			sessionStorage.setItem("bsView", "D");
	    			sessionStorage.setItem("as1View", "D");
	    			sessionStorage.setItem("as2View", "D");
	    			sessionStorage.setItem("profileId", obj.profileId);
	    			sessionStorage.setItem("groupId", obj.groupId);
	    			sessionStorage.setItem("totalTime",obj.totalTime);
	    			sessionStorage.setItem("rowIdentity", obj.identifier);
	    			sessionStorage.setItem("jobTitle", obj.jobTitle);
	    			sessionStorage.setItem("isCompleted", 0);
	    			sessionStorage.setItem("pageSequence", 1);
	    			sessionStorage.setItem("myPage", "BS");
	    			if (sessionStorage.getItem("giveUserOption")) {
	    				sessionStorage.removeItem("giveUserOption");
	    			}
	    			window.location = "beforeSketch.jsp";
	    		} else {
	    			alertify.alert(ob.statusDesc);
	    			$(".loader_bg").fadeOut();
	    		    $(".loader").hide();
	    		}
	    	});
		    } 
		});
}

function saveNewActivePage(activePage) {
	//alert(activePage);
	var LogoutVO = Spine.Model.sub();
	LogoutVO.configure("/user/auth/updateLoginInfoRowData", "jobReferenceString", "rowId", "newActivePage");
	LogoutVO.extend( Spine.Model.Ajax );

	//Populate the model with data to transder
	var logoutModel = new LogoutVO({  
		jobReferenceString: sessionStorage.getItem("jobReferenceNo"),
		rowId: sessionStorage.getItem("rowIdentity"),
		newActivePage: activePage
	});
	logoutModel.save(); //POST
}
function showOldPDF(){
	window.open('showPdf.jsp', '_blank');	
}