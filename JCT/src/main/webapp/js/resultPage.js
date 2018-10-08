/************************************************************************
 * Description: The js file is specifically for Action Plan page.		*
 * 																		*
 * Functions  : goPrevious() 											*
 * 					--> This function will help to 						*
 * 						navigate to previous page.						*
 * 						navigate to next page.							*
 * 				addEdits()												*
 * 					--> This will help to add edited diagrams  			*
 * 						when the page loads up.							*
 * 				addDiagWithoutEdits()									*
 * 					--> This will help to add not edited diagrams  		*
 * 						when the page loads up.							*
 * 																		*
 * Author	  : InterraIT												*
 * 																		*
 * 																		*
 ***********************************************************************/

var jrnAscii = "";
var jrnString = "";

/**
 * Function add to disable browser back button
 * while the page is loaded
 * @param null
 */
window.location.hash="";
window.location.hash="Again-No-back-button";//again because google chrome don't insert first hash into history
window.onhashchange=function(){window.location.hash="";};

window.setInterval(function(){
	updateTime();
}, 60000);

/**
 * Function to go to previous page
 */
function goPrevious() {
	$(".loader_bg").fadeIn();
	$(".loader").fadeIn();
	sessionStorage.setItem("pageSequence", 5);
	var navigation = Spine.Model.sub();
	navigation.configure("/user/navigation/goPrevious", "navigationSeq", "jobReferenceNos" ,"profileId", "linkClicked", "rowId");
	navigation.extend( Spine.Model.Ajax );
	//Populate the model
	var replyModel = new navigation({
		navigationSeq: "5",
		jobReferenceNos: sessionStorage.getItem("jobReferenceNo"),
		profileId: sessionStorage.getItem("profileId"),
		linkClicked: "N",
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
		if(sessionStorage.getItem("isLogout") == "N"){
			//sessionStorage.setItem("actionPlanQtnList", obj.list);
			//sessionStorage.setItem("existing", obj.existing);
			sessionStorage.setItem("snapShotURLS", obj.snapShot);
			sessionStorage.setItem("isCompleted", obj.isCompleted);
			window.location = "actionPlan.jsp";
		}
	});
	if (sessionStorage.getItem("infoDispReq")) {
		sessionStorage.removeItem("infoDispReq");
	}
}
/**
 * Function checks which radio option is selected and accordingly fires other functions
 */
function submitDecision(){
	if(document.getElementById("optionsRadios1").checked) {
		addDiagWithoutEdits();
	} else if(document.getElementById("optionsRadios2").checked){
		addEdits();
	} else if(document.getElementById("optionsRadios3").checked){
		disableStatus();
	}
}
/**
 * Function disables the diagrams and the diagrams will not be made public.
 */
function disableStatus(){
	$(".loader_bg").fadeIn();
	$(".loader").fadeIn();
	//Create a model
	var disableStatusModel = Spine.Model.sub();
	disableStatusModel.configure("/user/confirmation/disableStatus", "jobReferenceNo", "rowId");
	disableStatusModel.extend( Spine.Model.Ajax );
	
	//Populate the model
	var disableStatusPopulator = new disableStatusModel({  
		jobReferenceNo: sessionStorage.getItem("jobReferenceNo"),
		rowId: sessionStorage.getItem("rowIdentity")
	});
	//POST
	disableStatusPopulator.save();
	
	disableStatusModel.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
		if((statusCode == 500)){
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
			alertify.alert("<img src='../img/failed-icon.png'><br /><p>Unable to complete your request. Please contact administrator.</p>");
		} else {
			sessionStorage.setItem("pageSequence", 8);
			window.location = "finalPage.jsp";
		}
	});
	
	disableStatusModel.bind("ajaxError", function(record, xhr, settings, error) {
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
		alertify.alert("Unable to connect to the server.");
	});
}
/**
 * Function gives the user another chance to change the text (only) of the diagram and then the 
 * diagrams will be made search-able.
 */
function addEdits(){
	$(".loader_bg").fadeIn();
	$(".loader").fadeIn();
	sessionStorage.setItem("pageSequence", 4);
	//Create a model
	var diagEditModel = Spine.Model.sub();
	diagEditModel.configure("/user/confirmation/addEdits", "jobReferenceNo");
	diagEditModel.extend( Spine.Model.Ajax );
	
	//Populate the model
	var diagEditPopulator = new diagEditModel({  
		jobReferenceNo: sessionStorage.getItem("jobReferenceNo")
	});
	//POST
	diagEditPopulator.save();
	
	diagEditModel.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if((statusCode == 200)){
			if(sessionStorage.getItem("isLogout") == "N"){
				sessionStorage.setItem("pageSequence", 7); // Before sketch edit
				sessionStorage.setItem("isCompleted", obj.isCompleted);
				sessionStorage.setItem("total_json_add_task", obj.bsJson);
				var myJsonObj = JSON.parse(obj.bsJson);
				sessionStorage.setItem("total_count", myJsonObj.length+1);
				sessionStorage.setItem("div_intial_value", obj.initialJson);
				sessionStorage.setItem("bsEdit","resultPage.jsp");
				window.location = "beforeSketchEdit.jsp";
			} else { //It is logout
				//Create a model
				var LogoutVO = Spine.Model.sub();
				LogoutVO.configure("/user/auth/logout", "jobReferenceString", "rowId", "lastActivePage");
				LogoutVO.extend( Spine.Model.Ajax );

				//Populate the model with data to transder
				var logoutModel = new LogoutVO({  
					jobReferenceString: sessionStorage.getItem("jobReferenceNo"),
					rowId: sessionStorage.getItem("rowIdentity"),
					lastActivePage: "/user/view/resultPage.jsp"
				});
				logoutModel.save(); //POST
				
				LogoutVO.bind("ajaxError", function(record, xhr, settings, error) {
					$(".loader_bg").fadeOut();
				    $(".loader").hide();
					alertify.alert("Unable to connect to the server.");
				});
				
				LogoutVO.bind("ajaxSuccess", function(record, xhr, settings, error) {
					//var jsonStr = JSON.stringify(xhr);
					$(".loader_bg").fadeOut();
				    $(".loader").hide();
				});
			}
		} else if((statusCode == 500)){
			alertify.alert("<img src='../img/failed-icon.png'><br /><p>Unable to complete your request. Please contact administrator.</p>");
		}
	});
	
	diagEditModel.bind("ajaxError", function(record, xhr, settings, error) {
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
		alertify.alert("Unable to connect to the server.");
	});
}
/**
 * Function makes the diagrams search-able.
 */
function addDiagWithoutEdits(){
	$(".loader_bg").fadeIn();
	$(".loader").fadeIn();
	//Create a model
	var diagWithoutEditModel = Spine.Model.sub();
	diagWithoutEditModel.configure("/user/confirmation/addDiagWithoutEdits", "jobReferenceNo", "rowId");
	diagWithoutEditModel.extend( Spine.Model.Ajax );
	
	//Populate the model
	var diagWithoutEditPopulator = new diagWithoutEditModel({  
		jobReferenceNo: sessionStorage.getItem("jobReferenceNo"),
		rowId: sessionStorage.getItem("rowIdentity")
	});
	//POST
	diagWithoutEditPopulator.save();
	
	diagWithoutEditModel.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;		
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
		if ((statusCode == 500)) {
			alertify.alert("<img src='../img/failed-icon.png'><br /><p>Unable to complete your request. Please contact administrator.</p>");
		} else {
			sessionStorage.setItem("pageSequence", 8);
			window.location = "finalPage.jsp";
		}
	});
	
	diagWithoutEditModel.bind("ajaxError", function(record, xhr, settings, error) {
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
		alertify.alert("Unable to connect to the server.");
	});
}

$(document).ready(function() {
	//Enable or disable the information message in infoMsgDisp div
	if (!sessionStorage.getItem("infoDispReq")) {
		if (document.getElementById("infoMsgDisp")) {
			document.getElementById("infoMsgDisp").style.display = "block";
		}
	}
	
	sessionStorage.setItem("pageSequence", 6);
	if (sessionStorage.getItem('open')) {		
		window.open(generatePDFLink(), "", "width=700, height=600");
		sessionStorage.removeItem('open');
	}
});
/**
 * Function is used to download the pdf
 */
function downloadFile(){
	var link = document.getElementById('generateURL');
	//link.setAttribute("href", "../../user/confirmation/downloadFile/"+sessionStorage.getItem("jobReferenceNo")+".pdf");
	link.setAttribute("href", generatePDFLink());
	link.submit();
}
/**
 * Function generates the encoded version of the job reference number
 * @returns {String}
 */
function generatePDFLink(){
	jrnAscii = "";
	jrnString = "";
	retStr = "";
	var pdfLink = "../../user/confirmation/downloadFile/";
	var jRno = sessionStorage.getItem("jobReferenceNo");
	var jrNoLngth = jRno.length;
	for(var index=0; index<jrNoLngth; index++){
		jrnAscii = jrnAscii + jRno.charCodeAt(index);
	}
	var letters = replaceAsciiWithLetters(jrnAscii);
	
	pdfLink = pdfLink + letters + ".pdf";
	return pdfLink;
}
/**
 * Function replaces the ascii chars to letters
 * @param jrnAscii
 * @returns {String}
 */
function replaceAsciiWithLetters(jrnAscii){
	var myStr = "";
	var jrNoLngth = jrnAscii.length;
	for(var index=0; index<jrNoLngth; index++){
		myStr = myStr + getStr(jrnAscii[index]);
	}
	return myStr;
}
/**
 * Function replaces the index number with the string
 * @param indexNos
 * @returns {String}
 */
function getStr(indexNos) {
	var retStr = "";
	switch (parseInt(indexNos))
	  {
	  case 0:
		  retStr = "X";
		  break;
	  case 1:
		  retStr = "P";
		  break;
	  case 2:
		  retStr = "Y";
		  break;
	  case 3:
		  retStr = "A";
		  break;
	  case 4:
		  retStr = "L";
		  break;
	  case 5:
		  retStr = "B";
		  break;
	  case 6:
		  retStr = "M";
		  break;
	  case 7:
		  retStr = "T";
		  break;
	  case 8:
		  retStr = "S";
		  break;
	  case 9:
		  retStr = "H";
		  break;
	  }
	return retStr;
}
/**
 * Function stores the page info in the login info table and logs out
 */
function submitDecisionForLogout(){
	var LogoutVO = Spine.Model.sub();
	LogoutVO.configure("/user/auth/logout", "jobReferenceString", "rowId", "lastActivePage");
	LogoutVO.extend( Spine.Model.Ajax );

	//Populate the model with data to transder
	var logoutModel = new LogoutVO({  
		jobReferenceString: sessionStorage.getItem("jobReferenceNo"),
		rowId: sessionStorage.getItem("rowIdentity"),
		lastActivePage: "/user/view/resultPage.jsp"
	});
	logoutModel.save(); //POST
	
	LogoutVO.bind("ajaxError", function(record, xhr, settings, error) {
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
		alertify.alert("Unable to connect to the server.");
	});
	
	LogoutVO.bind("ajaxSuccess", function(record, xhr, settings, error) {
		//var jsonStr = JSON.stringify(xhr);
		sessionStorage.clear();
		window.location = "../signup-page.jsp";
	});

}
function showOldPDF(){
	window.open('showPdf.jsp', '_blank');	
}