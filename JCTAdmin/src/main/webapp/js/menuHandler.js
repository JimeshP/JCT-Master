var popup = null;
var firstResult=0;
/**
 * Function handles all entry point for reports.
 * @param type
 */
function populateReports(type){
	$(".loader_bg").fadeIn();
	$(".loader").fadeIn();
	if (sessionStorage.getItem("occupation")) {
		sessionStorage.removeItem("occupation");
	}
	sessionStorage.setItem("fnLSelected", "");
	sessionStorage.setItem("jLSelected", "");
	if(type == "BS") {
		firstResult=0;
		sessionStorage.setItem("firstResult",firstResult);
		populateBeforeSketchReport();
	} else if(type == "AS") {				
		firstResult=0;
		sessionStorage.setItem("firstResultAS",firstResult);
		populateAfterSketchReport();
	} else if(type == "AP") {
		firstResult=0;
		sessionStorage.setItem("firstResultAP",firstResult);
		populateActionPlanReport();
	} else if(type == "BSTOAS") {
		firstResult=0;
		sessionStorage.setItem("firstResultAsToBs",firstResult);
		populateBeforeSketchToAfterSketchReport();
	} else if(type == "MISC") {
		//firstResult=0;
		//sessionStorage.setItem("firstResultAsToBs",firstResult);
		populateMiscReport();
	} else if (type == "GP") {
		populateProfileDropDown();
	} 
	else {
		alertify.alert("We do not have this report for you.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide(); 
	}
}

/**
 * Function fetches miscellaneous report 
 * @param null
 */
function populateMiscReport(){
	var bsModel = Spine.Model.sub();
	bsModel.configure("/admin/reports/populateMiscReport", "none");
	bsModel.extend( Spine.Model.Ajax );
	//Populate the model
	var populateModel = new bsModel({  
		none: ""
	});
	populateModel.save(); //POST
	
	bsModel.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to contact the server.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	});
	
	bsModel.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		if(obj.statusCode == 200){
			sessionStorage.setItem("jsonStr",jsonStr); //contains whole population string
			window.location = "/admin/view/reportsMisc.jsp";
		}
	});
}


/**
 * Function fetches before sketch report
 * @param null
 */
function populateBeforeSketchReport(){
	var bsModel = Spine.Model.sub();
	bsModel.configure("/admin/reports/populateBeforeSketchReport", "recordIndex");
	bsModel.extend( Spine.Model.Ajax );
	
	//Populate the model
	var populateModel = new bsModel({  
		recordIndex: sessionStorage.getItem("firstResult")
	});
	populateModel.save(); //POST
	
	bsModel.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to contact the server.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	});
	
	bsModel.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		if(obj.statusCode == 200){
			sessionStorage.setItem("jsonStr",jsonStr); //contains whole population string
			window.location = "/admin/view/reportsBeforeSketch.jsp";
		}
	});
}


/**
 * Function fetches action plan report
 * @param null
 */
function populateActionPlanReport(){
	var bsModel = Spine.Model.sub();
	bsModel.configure("/admin/reports/populateActionPlanReport", "recordIndex");
	bsModel.extend( Spine.Model.Ajax );
	
	//Populate the model
	var populateModel = new bsModel({  
		recordIndex: sessionStorage.getItem("firstResultAP")
	});
	populateModel.save(); //POST
	
	bsModel.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to contact the server.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	});
	
	bsModel.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		if(obj.statusCode == 200){
			sessionStorage.setItem("jsonStr",jsonStr); //contains whole population string
			window.location = "/admin/view/reportsActionPlan.jsp";
		}
	});
}

/**
 * Function fetches after sketch report
 * @param null
 */
function populateAfterSketchReport(){
	var asModel = Spine.Model.sub();
	asModel.configure("/admin/reports/populateAfterSketchReport", "recordIndex");
	asModel.extend( Spine.Model.Ajax );
	
	//Populate the model
	var populateModel = new asModel({  
		recordIndex: sessionStorage.getItem("firstResultAS")
	});
	populateModel.save(); //POST
	
	asModel.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to contact the server.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	});
	
	asModel.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		if(obj.statusCode == 200){
			sessionStorage.setItem("jsonStr",jsonStr); //contains whole population string
			window.location = "/admin/view/reportsAfterSketch.jsp";
		}
	});
}


/**
 * Function fetches before sketch to after sketch report
 * @param null
 */
function populateBeforeSketchToAfterSketchReport(){
	var asModel = Spine.Model.sub();
	asModel.configure("/admin/reports/populateBeforeSketchToAfterSketchReport", "recordIndex");
	asModel.extend( Spine.Model.Ajax );
	
	//Populate the model
	var populateModel = new asModel({  
		recordIndex: sessionStorage.getItem("firstResultAsToBs")
	});
	populateModel.save(); //POST
	
	asModel.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to contact the server.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	});
	
	asModel.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		if(obj.statusCode == 200){
			sessionStorage.setItem("jsonStr",jsonStr); //contains whole population string
			window.location = "/admin/view/reportsBeforeSketchToAfterSketch.jsp";
		}
	});
}


/**
 * Function fetches user profile dropdown data
 * @param null
 */
function populateProfileDropDown(){
	var bsModel = Spine.Model.sub();
	bsModel.configure("/admin/reports/populateProfileDropDown", "recordIndex");
	bsModel.extend( Spine.Model.Ajax );
	
	//Populate the model
	var populateModel = new bsModel({  
		recordIndex: sessionStorage.getItem("firstResult")
	});
	populateModel.save(); //POST
	
	bsModel.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to contact the server.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	});
	
	bsModel.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		if(obj.statusCode == 200){
			sessionStorage.setItem("jsonStr",jsonStr); //contains whole population string
			window.location = "/admin/view/reportsGroupProfile.jsp";
		}
	});
}


$(document).ready(function() {
	populateMenu();
});



/**
 * Function to populate menu
 */

/*******Change 27.2.15***********/

function populateMenu() {
    var htmlMenuString = "";
    var hyperLink = "";
    var functionCall = "";
    var imageSrc = "";
    var menuItems = sessionStorage.getItem("menuItems");    
    var menuCounterSplit = menuItems.split("<!>");
    for (var index = 0; index<menuCounterSplit.length-1; index++ ) {    	
    var menuString = menuCounterSplit[index].split("`");
    var headingMenu = menuString[0];   // [0] bCoz only one sub menu is there
    var subMenuAndPages = menuString[1].split("+");  
    if (headingMenu == "Manage<br />User") { // Heading menu Logo setting
                    imageSrc = "/admin/img/icon-1.png";
    }  
    else if (headingMenu == "Manage<br />Users") { // Heading menu Logo setting
                    imageSrc = "/admin/img/icon-1.png";
    } 
    else if (headingMenu == "Application<br />Settings") {
                    imageSrc = "/admin/img/icon-2.png";
    }  else if (headingMenu == "Content<br />Configuration") {
                    imageSrc = "/admin/img/icon-3.png";
    } else if (headingMenu == "My<br />Account") {
                    imageSrc = "/admin/img/icon-4.png";
    } else if (headingMenu == "Reports") {
                    imageSrc = "/admin/img/icon-5.png";
    } else if (headingMenu == "Preview<br />Tool") {
                    imageSrc = "/admin/img/icon-5.png";
    } else if (headingMenu == "Refund<br />Request") {
                    imageSrc = "/admin/img/refund.png";
    }   
    
    if(subMenuAndPages.length -1 == 1) {                                  
	    var internalSplit = subMenuAndPages[0].split("~");
	     var pageSplit = internalSplit[1].split(",");
	      	if (internalSplit[0] == "Preview") {
            functionCall = "connectToTool()";
            htmlMenuString = htmlMenuString + "<li class='dropdown-submenu' onclick = "+ functionCall+"><a tabindex='-1' href='#'><span><img src='"+imageSrc+"' alt=''></span>"+headingMenu+"</a>";
	    }
	    else if (internalSplit[0] == "User Progress Report") {
	    	//functionCall = "populateUserProgressReports()";
            htmlMenuString = htmlMenuString + "<li class='dropdown-submenu'><a tabindex='-1' href='"+pageSplit[0]+"'><span><img src='"+imageSrc+"' alt=''></span>"+headingMenu+"</a>";
        }
	    else {
	        htmlMenuString = htmlMenuString + "<li class='dropdown-submenu'><a tabindex='-1' href='/admin/view/"+internalSplit[1]+"'><span><img src='"+imageSrc+"' alt=''></span>"+headingMenu+"</a>";
	    }   
	  
    } else {
    	htmlMenuString = htmlMenuString + "<li class='dropdown-submenu'><a tabindex='-1' href='#'><span><img src='"+imageSrc+"' alt=''></span>"+headingMenu+"</a>";
    	htmlMenuString = htmlMenuString + "<ul class='dropdown-menu cstm-sub'>";
    	for (var innerIndex = 0; innerIndex < subMenuAndPages.length-1; innerIndex++) {
    		var internalSplit = subMenuAndPages[innerIndex].split("~");
            if (headingMenu == "Reports") {
                var functionCall = "";
                var hyperLink = "";
               // alert(internalSplit[0]);
                if (internalSplit[0] == "Before Sketch") {
                	functionCall = "populateReports('BS')";
                    hyperLink = "<a href = '#' onclick = "+functionCall;
                } else if (internalSplit[0] == "After Diagram") {
                    functionCall = "populateReports('AS')";
                    hyperLink = "<a href = '#' onclick = "+functionCall;
                } else if (internalSplit[0] == "Before Sketch To After Diagram") {
                    functionCall = "populateReports('BSTOAS')";
                    hyperLink = "<a href = '#' onclick = "+functionCall;
                } else if (internalSplit[0] == "Action Plan") {
                    functionCall = "populateReports('AP')";
                    hyperLink = "<a href = '#' onclick = "+functionCall;
                } else if (internalSplit[0] == "Misc Report") {
                    functionCall = "populateReports('MISC')";
                    hyperLink = "<a href = '#' onclick = "+functionCall;
                    
                } else if (internalSplit[0] == "Group Profile Reports") {
                    functionCall = "populateReports('GP')";
                    hyperLink = "<a href = '#' onclick = "+functionCall;
                } else if (internalSplit[0] == "Merged Report") {
                    hyperLink = "<a href = '/admin/view/"+internalSplit[1]+"'";
                } else {
                    hyperLink = "<a href = '/admin/view/"+internalSplit[1]+"'";
                }
                htmlMenuString = htmlMenuString + "<li>"+hyperLink+">"+internalSplit[0]+"</a></li><li class='divider'></li>";
            } else {
                htmlMenuString = htmlMenuString + "<li><a href='/admin/view/"+internalSplit[1]+"' >"+internalSplit[0]+"</a></li><li class='divider'></li>";
            } 
    	}
    }                              
    htmlMenuString = htmlMenuString + "</ul></li><li class='divider'></li>";   
    }
    //htmlMenuString = htmlMenuString + "<li><a href='#'><span>&nbsp;</span> &nbsp;&nbsp;<span>&nbsp;&nbsp;</span></a></li>";
    /*if(sessionStorage.getItem("userRole") == 2) {
    	 //htmlMenuString = htmlMenuString + "<li><a href='#'><span>&nbsp;</span> &nbsp;&nbsp;<span>&nbsp;&nbsp;</span></a></li>";
    }*/ 
    document.getElementById("dynaMenuPopId").innerHTML = htmlMenuString;
}


/**
 * Function connects to the tool for preview
 */
function connectToTool() {
	$(".loader_bg").fadeIn();
	alertify.set({ buttonReverse: true });
	alertify.confirm("Preview tool will not save any data you enter.<br> Click OK to continue.", function (e) {
		if (e) {			
			var bsModel = Spine.Model.sub();
			bsModel.configure("/user/auth/facilitatorAuthenticate", "jctEmail");
			bsModel.extend( Spine.Model.Ajax );
			
			//Populate the model
			var populateModel = new bsModel({  
				jctEmail: sessionStorage.getItem("jctEmail")
			});
			populateModel.save(); //POST

			bsModel.bind("ajaxError", function(record, xhr, settings, error) {
				alertify.alert("Unable to contact the server.");
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			});
			
			bsModel.bind("ajaxSuccess", function(record, xhr, settings, error) {
				var jsonStr = JSON.stringify(xhr);
				var obj = jQuery.parseJSON(jsonStr);
				var statusCode = obj.statusCode;
				var url = obj.url;
				if((url == "/user/view/resultPage.jsp") || (url == "/user/view/resultPageFinal.jsp")) {
					sessionStorage.setItem("infoDispReq", "No");
				}
				if (obj.jctUserId) {
					sessionStorage.setItem("jctUserId", obj.jctUserId);
				}
					if (statusCode == 200) { //If user activated and login is success
						sessionStorage.setItem("accountExp", "NA"); // set as NA as he is a facilitator
						sessionStorage.setItem("linkClicked", "N");
						sessionStorage.setItem("name", obj.firstName);
						sessionStorage.setItem("jctEmail", sessionStorage.getItem("jctEmail"));
						sessionStorage.setItem("pageSequence", obj.pageSequence);
						sessionStorage.setItem("jobReferenceNo", obj.jobRefNo); //Job reference
						sessionStorage.setItem("next_page", obj.lastPage);
						sessionStorage.setItem("bsView", "D");
						sessionStorage.setItem("as1View", "D");
						sessionStorage.setItem("as2View", "D");
						sessionStorage.setItem("profileId", obj.profileId);
						
						sessionStorage.setItem("groupId", obj.groupId);
						sessionStorage.setItem("isCompleted", obj.isCompleted);
						sessionStorage.setItem("myPage", "BS");
						if (obj.source) {
							sessionStorage.setItem("src", obj.source);
						}
						if((obj.bsJson != "none") && (null != obj.bsJson)) {
							sessionStorage.setItem("total_json_add_task", obj.bsJson); //Job reference
							var myJsonObj = JSON.parse(obj.bsJson);
							sessionStorage.setItem("total_count", myJsonObj.length+1);
							sessionStorage.setItem("div_intial_value", obj.initialJson);
							sessionStorage.setItem("myPage", "BS");
						}
						sessionStorage.setItem("totalTime",obj.totalTime);
						sessionStorage.setItem("rowIdentity", obj.identifier);
						sessionStorage.setItem("jobTitle", obj.jobTitle);
						sessionStorage.setItem("isLogout", "N");
						
						if(obj.pageSequence == 1){
							
							sessionStorage.setItem("bktp", "Y");
							sessionStorage.setItem("pageSequence", 1);
							
							//sessionStorage.setItem("snapShotURLS", obj.jctBaseString);
							sessionStorage.setItem("myPage", "BS");
						}
						
						if(obj.pageSequence == 2){
							sessionStorage.setItem("snapShotURLS", obj.jctBaseString);
							sessionStorage.setItem("myPage", "BS");
						}
						if(obj.pageSequence == 3){
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
							sessionStorage.setItem("myPage", "AS");
						} if(obj.pageSequence == 4){
							sessionStorage.setItem("total_json_obj", obj.pageOneJson);
							sessionStorage.setItem("div_intial_value", obj.divInitialValue);
							sessionStorage.setItem("total_json_after_sketch_final", obj.afterSketch2TotalJsonObj);
							sessionStorage.setItem("total_count", obj.asTotalCount);
							sessionStorage.setItem("divHeight", obj.divHeight);
							sessionStorage.setItem("divWidth", obj.divWidth);
							sessionStorage.setItem("myPage", "AS");
						} if(obj.pageSequence == 5){
							sessionStorage.setItem("existing", obj.existing);
							sessionStorage.setItem("snapShotURLS", obj.snapShot);
							sessionStorage.setItem("myPage", "AS");
						} if(obj.pageSequence == 6) {
							sessionStorage.setItem("myPage", "AS");
						}
						sessionStorage.setItem("url", url);
						if (obj.excCompletionCount > 0) {
							if (url == "/user/view/finalPage.jsp") {
								sessionStorage.setItem("pageSequence", 8);
							}
							//window.location = url;
						} else {
							if (null == sessionStorage.getItem("isReset")) {
								//window.location = "view/landing-page.jsp";
							} else if (sessionStorage.getItem("isReset") != "Y"){
								//window.location = "landing-page.jsp";
							}else{
								//window.location = "view/landing-page.jsp";
							}
						}
					} 
				
					var params  = 'width='+screen.width;
					params += ', height='+screen.height;
					params += ', top=0, left=0';
					params += ', fullscreen=yes';
					params += ', scrollbars=1';
					
					//windowHandle = window.open("/user/view/beforeSketch.jsp", 'windowname4', params);
					windowHandle = window.open("/user/view/landing-page.jsp", 'windowname4', params);
					if (!windowHandle.closed)
					window.closed = false;
					$(".loader_bg").fadeOut();
				    $(".loader").hide();				
			});
		} else {
			$(".loader_bg").fadeOut();
		}
	});	//	function(e)
}