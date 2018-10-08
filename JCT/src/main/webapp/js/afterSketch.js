/**
 * Function call while refresh the page
 */
window.onunload = function(event) {
	return saveMappingWhileRefresh();
};

window.setInterval(function(){
	updateTime();
}, 60000);

/**
 * Function add to populate the instruction description
 * while the page is loaded
 * @param null
 */
/**** 02.03.15****/

$(document).ready(function() {	
	
	var insDisplay = sessionStorage.getItem("as1View");
	if(null == insDisplay || insDisplay == "" || insDisplay == "D"){
		populatePopUp();
	} else {
		$('#myModal').modal('hide');
	}
	
	populateInstruction();
	if(ClientSession.get("strength_count") == 0 && ClientSession.get("value_count") == 0 && ClientSession.get("passion_count") == 0){
		ClientSession.set("total_json_obj" , null);
	}
	/* popover added to show the attribute description on mouse hover
	 * Dated : 16.07.2014
	 */
	/*$(".ckecked_strength").tooltip();
	$(".ckecked_value").tooltip();
	$(".ckecked_passion").tooltip();*/
});

/**
 * Function add to disable browser back button
 * while the page is loaded
 * @param null
 */
window.location.hash="";
window.location.hash="Again-No-back-button";//again because google chrome don't insert first hash into history
window.onhashchange=function(){window.location.hash="";};

/********** Added for instruction toggle ***********/
    $(".btn-slide").click(function(){
    $("#panel").slideToggle("slow");
    $(this).toggleClass("active"); 
    return false;
  });


var strengthHTML = "";
var passionHTML = "";
var valueHTML = "";
var strengthCheckedCount = 0;
var valueCheckedCount = 0;
var passionCheckedCount = 0;

var pageNavigation = 0; 
var backToPrevious = 0;

if(ClientSession.get("checked_strength") == null){
	var checkedStrength = 0;
} else {
	var checkedStrength = ClientSession.get("checked_strength");
}

/**
 * Function to To draw Strength element
 * @param objS
 */
var countStrength = 1;
function checkStrength(objS) {
	var nAgt = navigator.userAgent;
	if ((verOffset=nAgt.indexOf("MSIE"))!=-1) {
		objS.setAttribute('disabled', true);
	}
	if(objS.checked==true){	
		 $(".popover fade top in").popover('hide');
		/* find the description against the attribute name
		 * Dated : 16.07.2014
		 */
		var description = getAttrDecription(objS.value, "STR");				
		checkedStrength++;
		if(ClientSession.get("strength_count") == null){
			var divId = "divStrength_" + countStrength;	
			var textId = "textStrength_" + countStrength;
			var valS = objS.value;	
			//alert(valS);
			/* popover added to show the strength description on mouse hover
			 * Dated : 16.07.2014
			 */
		    var $element = $('<div id="' + divId +'" class="plus-sign commonDrag" data-toggle="popover" data-placement="bottom" title="" data-original-title="'+description+'" style="display:block;"><img src="../img/strength.png" width="110" height="110" style="no-repeat 0 9px;"/><div class= "plus-sign_span" type="text" id="'+textId+'"><p>'+valS+'</p></div>');
		} else if(ClientSession.get("strength_count") == 0 ){
			var divId = "divStrength_" + countStrength;	
			var textId = "textStrength_" + countStrength;
			var valS = objS.value;	
			//alert(valS);
			/* popover added to show the strength description on mouse hover
			 * Dated : 16.07.2014
			 */
		    var $element = $('<div id="' + divId +'" class="plus-sign commonDrag" data-toggle="popover" data-placement="bottom" title="" data-original-title="'+description+'" style="display:block;"><img src="../img/strength.png" width="110" height="110" style="no-repeat 0 9px;"/><div class= "plus-sign_span" type="text" id="'+textId+'"><p>'+valS+'</p></div>');
		} else {			
			countStrength = ClientSession.get("strength_count");
			var divId = "divStrength_" + countStrength;
			var textId = "textStrength_" + countStrength;
			var valS = objS.value;		
			//alert(valS);
			/* popover added to show the strength description on mouse hover
			 * Dated : 16.07.2014
			 */
		    var $element = $('<div id="' + divId +'" class="plus-sign commonDrag" data-toggle="popover" data-placement="bottom" title="" data-original-title="'+description+'" style="display:block;"><img src="../img/strength.png" width="110" height="110" style="no-repeat 0 9px;"/><div class= "plus-sign_span" type="text" id="'+textId+'"><p>'+valS+'</p></div>');	     		    
		}	
			countStrength++;	  	
		  	ClientSession.set("strength_count", countStrength);	 
		    $("#strength_pageWrap").append($element);
		    $element.popover(); 		   
		} else {	
			$(".popover fade top in").popover('hide');
			checkedStrength--;
		var obj = document.getElementsByClassName("plus-sign");		
		for (var i = 0; i<obj.length;i++){		
			var objectId = obj[i].id;// Strength_1
			var divCnt = objectId.substring(12);
			//var divValue = document.getElementById("textStrength_"+divCnt).innerHTML;
			var divValue  = $("#textStrength_" + divCnt).find('p').text();			
			if(divValue == objS.value ){
				var div = document.getElementById(objectId);
			    if (div) {
			    	div.setAttribute("style", "display:none");
			    }
			}						
		}		
	 }	
	if ((verOffset=nAgt.indexOf("MSIE"))!=-1) {
		 setTimeout(function() {
			 objS.removeAttribute("disabled");
		     }, 600);
	}	
	if(checkedStrength == 4){		
		$('.ckecked_strength input[type=checkbox]').not(':checked').attr("disabled",true);
	} else {
		$('.ckecked_strength input[type=checkbox]').not(':checked').attr("disabled",false);
	}
	ClientSession.set("checked_strength", checkedStrength);	  	
   	return false;
}



if(ClientSession.get("checked_value") == null){
	var checkedValue = 0;
} else {
	var checkedValue = ClientSession.get("checked_value");
}
/**
 * Function to To draw Value element
 * @param objV
 */
function checkValue(objV) {	
	$('.popover').css('display','none');
	var nAgt = navigator.userAgent;
	if ((verOffset=nAgt.indexOf("MSIE"))!=-1) {
		objV.setAttribute('disabled', true);
	}
	var countValue = 1;
	if(objV.checked==true){		
		 $(".popover fade top in").popover('hide');
		/* find the description against the attribute name
		 * Dated : 16.07.2014
		 */
		var description = getAttrDecription(objV.value, "VAL");	
		checkedValue++;	
		if(ClientSession.get("value_count") == null){
			var divId = "divValue_" + countValue;
			var textId = "textValue_" + countValue;
			var	valV = objV.value;			
			/* popover added to show the value description on mouse hover
			 * Dated : 16.07.2014
			 */
			var $element = $('<div id="' + divId +'" class="oval commonDrag" data-toggle="popover" data-placement="bottom" title="" data-original-title="'+description+'" style="display:block;"><img src="../img/value.png" width="130" height="100" style="no-repeat 0 9px;"/><div class="oval_span" id="'+textId+'"><p>'+valV+'</p></div>');	
		} else if(ClientSession.get("value_count") == 0){
			var divId = "divValue_" + countValue;
			var textId = "textValue_" + countValue;
			var	valV = objV.value;		
			/* popover added to show the value description on mouse hover
			 * Dated : 16.07.2014
			 */
			var $element = $('<div id="' + divId +'" class="oval commonDrag" data-toggle="popover" data-placement="bottom" title="" data-original-title="'+description+'" style="display:block;"><img src="../img/value.png" width="130" height="100" style="no-repeat 0 9px;"/><div class= "oval_span" id="'+textId+'"><p>'+valV+'</p></div>');
		} else {
			countValue = ClientSession.get("value_count");
			var divId = "divValue_" + countValue;
			var textId = "textValue_" + countValue;
			var	valV = objV.value;		
			/* popover added to show the value description on mouse hover
			 * Dated : 16.07.2014
			 */
			var $element = $('<div id="' + divId +'" class="oval commonDrag" data-toggle="popover" data-placement="bottom" title="" data-original-title="'+description+'" style="display:block;"><img src="../img/value.png" width="130" height="100" style="no-repeat 0 9px;"/><div class= "oval_span" id="'+textId+'"><p>'+valV+'</p></div>');
			}	
		countValue++;
	  	ClientSession.set("value_count", countValue);
	    $("#value_pageWrap").append($element);
	    $element.popover(); 
	} else {	
		 $(".popover fade top in").popover('hide');
		checkedValue--;
		var obj = document.getElementsByClassName("oval");
		for (var i = 0; i<obj.length;i++){
			var objectId = obj[i].id;// Strength_1
			var divCnt = objectId.substring(9);
			//var divValue = document.getElementById("textValue_"+divCnt).innerHTML;
			var divValue  = $("#textValue_" + divCnt).find('p').text();
			if(divValue == objV.value ){
				var div = document.getElementById(objectId);
			    if (div) {
			    	div.setAttribute("style", "display:none");
			    }
			}						
		}		
	  }
	if ((verOffset=nAgt.indexOf("MSIE"))!=-1) {
		 setTimeout(function() {
			 objV.removeAttribute("disabled");
		     }, 600);
	}	
	if(checkedValue == 4){		
		$('.ckecked_value input[type=checkbox]').not(':checked').attr("disabled",true);
	} else {
		$('.ckecked_value input[type=checkbox]').not(':checked').attr("disabled",false);
	}
	ClientSession.set("checked_value", checkedValue);	  	
    return false;
}


if(ClientSession.get("checked_passion") == null){
	var checkedPassion = 0;
} else {
	var checkedPassion = ClientSession.get("checked_passion");
}
var countPassion = 1;
/**
 * Function to To draw Passion element
 * @param objP
 */
function checkPassion(objP) {
	$('.popover').css('display','none');
	var nAgt = navigator.userAgent;
	if ((verOffset=nAgt.indexOf("MSIE"))!=-1) {
		objP.setAttribute('disabled', true);
	}
	if(objP.checked==true){
		/* find the description against the attribute name
		 * Dated : 16.07.2014
		 */
		var description = getAttrDecription(objP.value, "PAS");	
		checkedPassion++;
		if(ClientSession.get("passion_count") == null){
			var divId = "divPassion_" + countPassion;	
			var textId = "textPassion_" + countPassion;			
			var	valP = objP.value;
			/* popover added to show the passion description on mouse hover
			 * Dated : 16.07.2014
			 */
		    var $element = $('<div id="' + divId +'" class="up-triangle commonDrag" data-toggle="popover" data-placement="bottom" title="" data-original-title="'+description+'" style="display:block;"><img src="../img/passion.png" width="112" height="110" style="no-repeat 0 9px;"/><div class="up-triangle_span" id="'+textId+'"><p>'+valP+'</p></div>');	     
		} else if(ClientSession.get("passion_count") == 0){
			var divId = "divPassion_" + countPassion;	
			var textId = "textPassion_" + countPassion;			
			var	valP = objP.value;
			/* popover added to show the passion description on mouse hover
			 * Dated : 16.07.2014
			 */
		    var $element = $('<div id="' + divId +'" class="up-triangle commonDrag" data-toggle="popover" data-placement="bottom" title="" data-original-title="'+description+'" style="display:block;"><img src="../img/passion.png" width="112" height="110" style="no-repeat 0 9px;"/><div class = "up-triangle_span" id="'+textId+'"><p>'+valP+'</p></div>');
			}		
		else {
			countPassion = ClientSession.get("passion_count");
			var divId = "divPassion_" + countPassion;
			var textId = "textPassion_" + countPassion;			
			var	valP = objP.value;
			/* popover added to show the passion description on mouse hover
			 * Dated : 16.07.2014
			 */
		    var $element = $('<div id="' + divId +'" class="up-triangle commonDrag" data-toggle="popover" data-placement="bottom" title="" data-original-title="'+description+'" style="display:block;"><img src="../img/passion.png" width="112" height="110" style="no-repeat 0 9px;"/><div class="up-triangle_span" id="'+textId+'"><p>'+valP+'</p></div>');
			}
		countPassion++;
	  	ClientSession.set("passion_count", countPassion);
	    $("#passion_pageWrap").append($element);
	    $element.popover(); 
	} else {	
		checkedPassion--;
		var obj = document.getElementsByClassName("up-triangle");
		for (var i = 0; i<obj.length;i++){
			var objectId = obj[i].id;// Strength_1
			var divCnt = objectId.substring(11);
			//var divValue = document.getElementById("textPassion_"+divCnt).innerHTML;
			var divValue  = $("#textPassion_" + divCnt).find('p').text();
			if(divValue == objP.value ){
				var div = document.getElementById(objectId);
			    if (div) {
			    	div.setAttribute("style", "display:none");
			    }
			}						
		}		
	 }
	if ((verOffset=nAgt.indexOf("MSIE"))!=-1) {
		 setTimeout(function() {
			 objP.removeAttribute("disabled");
		     }, 600);
	}	
	if(checkedPassion == 4){		
		$('.ckecked_passion input[type=checkbox]').not(':checked').attr("disabled",true);
	} else {
		$('.ckecked_passion input[type=checkbox]').not(':checked').attr("disabled",false);
	}
	ClientSession.set("checked_passion", checkedPassion);	
    return false;
}


/**
 * Function to store the json string  
 * while navigate to next page
 * @param null
 */
function goToNextPage(){	
	sessionStorage.setItem("as1View","A");
	var check = true;	
	if(sessionStorage.getItem("src") != "Admin") {
		if(sessionStorage.getItem("isLogout")=="N"){
			/************* check number of Value **************/
			if ($('#value_pageWrap').children('.oval:visible').length<2) {
				check = false;
				alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please select at least two Values.</p>");
				return false;
			} 
			/************* check number of Strength **************/
			else if($('#strength_pageWrap').children('.plus-sign:visible').length<2) {
				check = false;
				alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please select at least two Strengths.</p>");
				return false;
			} 		
			/************* check number of Passion **************/
			else if ($('#passion_pageWrap').children('.up-triangle:visible').length<2) {
				check = false;
				alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please select at least two Passions.</p>");
				return false;
			}
		}
	}
	if(check == true) {
		
		/************* set the value to session for logout **************/
		ClientSession.set("checked_strength", checkedStrength);
		ClientSession.set("checked_value", checkedValue);	
		ClientSession.set("checked_passion", checkedPassion);
		
		pageNavigation = 1;
		var cntStrength = 1;
		var totalJson = [];	
		var idx = 0;		
		var divStrength = document.getElementById("divStrength_" + cntStrength);
		while(divStrength != null) {
				var unitJson = {};				
				unitJson["divId"] = document.getElementById("divStrength_" + cntStrength).id;	
				unitJson["textId"] = document.getElementById("textStrength_" + cntStrength).id;	
				//unitJson["divValue"] = document.getElementById("textStrength_"+ cntStrength).innerHTML;	
				unitJson["divValue"] = $("#divStrength_" + cntStrength).find('p').text(); 
				unitJson["style"] = document.getElementById("divStrength_" + cntStrength).getAttribute("style");
				unitJson["positionLeft"] = document.getElementById("divStrength_" + cntStrength).style.left;
				unitJson["positionTop"] = document.getElementById("divStrength_" + cntStrength).style.top;	
				totalJson[idx++] = unitJson;		
				divStrength = document.getElementById("divStrength_" + ++cntStrength);		
		}	
		
		var cntValue = 1;
		var divValue = document.getElementById("divValue_" + cntValue);
		while(divValue != null) {
			var unitJson = {};		
			unitJson["divId"] = document.getElementById("divValue_" + cntValue).id;	
			unitJson["textId"] = document.getElementById("textValue_"+ cntValue).id;		
			//unitJson["divValue"] = document.getElementById("textValue_"+ cntValue).innerHTML;
			unitJson["divValue"] = $("#divValue_" + cntValue).find('p').text(); 
			unitJson["style"] = document.getElementById("divValue_" + cntValue).getAttribute("style");
			unitJson["positionLeft"] = document.getElementById("divValue_" + cntValue).style.left;
			unitJson["positionTop"] = document.getElementById("divValue_" + cntValue).style.top;		
			totalJson[idx++] = unitJson;
			divValue = document.getElementById("divValue_" + ++cntValue);
		}

		var cntPassion = 1;
		var divPassion = document.getElementById("divPassion_" + cntPassion);
		while(divPassion != null) {
			var unitJson = {};		
			unitJson["divId"] = document.getElementById("divPassion_" + cntPassion).id;	
			unitJson["textId"] = document.getElementById("textPassion_"+ cntPassion).id;		
			//unitJson["divValue"] = document.getElementById("textPassion_"+ cntPassion).innerHTML;	
			unitJson["divValue"] = $("#divPassion_" + cntPassion).find('p').text(); 
			unitJson["style"] = document.getElementById("divPassion_" + cntPassion).getAttribute("style");
			unitJson["positionLeft"] = document.getElementById("divPassion_" + cntPassion).style.left;
			unitJson["positionTop"] = document.getElementById("divPassion_" + cntPassion).style.top;		
			totalJson[idx++] = unitJson;
			divPassion = document.getElementById("divPassion_" + ++cntPassion);
		}
		ClientSession.set("total_json_obj",JSON.stringify(totalJson));
		ClientSession.set("Page_navigation", pageNavigation);	
		countCheckedDiv();
		if(sessionStorage.getItem("isLogout")) {
			storeAfterSketchOne(totalJson, "None");
		} else {
			storeAfterSketchOne(totalJson, "Next");
		}
		
	}
}

/**
 * Function add to calculate the total time
 * @param null
 */
function getTotalTimeSpent(){
	var timeCounterSplit = document.getElementById("stwa").value.split(":");
	var hour = parseInt(timeCounterSplit[0]);
	var min = parseInt(timeCounterSplit[1]);
	var sec = parseInt(timeCounterSplit[2]);
	if(min != 0){
		min = min * 60;
	}
	if(hour != 0){
		hour = hour * 60 * 60;
	}
	var totalTimeInSec = hour+min+sec;
	return totalTimeInSec;
}
/**
 * Function sends the json string to the server for storing
 * @param jsonString
 */
function storeAfterSketchOne(jsonString, decider) {
	$(".loader_bg").fadeIn();
    $(".loader").fadeIn();
    var timeSpent = getTotalTimeSpent();
    sessionStorage.setItem("as1View","A");
    strengthHTML = "";
    valueHTML = "";
    passionHTML = "";
    strengthCheckedCount = 0;
    valueCheckedCount = 0;
    passionCheckedCount = 0;
    //Fetch Id for checked items for value, passion and strength
    for (var i=0;i<document.getElementsByName('strengthList').length;i++){
    	if(document.getElementsByName('strengthList')[i].checked){
    		strengthCheckedCount = strengthCheckedCount + 1;
    		strengthHTML = strengthHTML + document.getElementsByName('strengthList')[i].id+"#";
		}
	}
	
	for (var i=0;i<document.getElementsByName('valueList').length;i++){
		if(document.getElementsByName('valueList')[i].checked){
			valueCheckedCount = valueCheckedCount + 1;
			valueHTML = valueHTML + document.getElementsByName('valueList')[i].id+"#";
		}
	}
	
	for (var i=0;i<document.getElementsByName('passionList').length;i++){
		if(document.getElementsByName('passionList')[i].checked){
			passionCheckedCount = passionCheckedCount + 1;
			passionHTML = passionHTML + document.getElementsByName('passionList')[i].id+"#";
		}
	}
    
	/************** Comment out to resole the issue no 36 
	 * 			During logout from application, system prompts information 
	 * 			message for the task pending on the page.*****************/
		var checkedElems = JSON.parse(sessionStorage.getItem("checked_element"));
	    if(null == sessionStorage.getItem("checked_passion") || isNaN(sessionStorage.getItem("checked_passion")) ){
	    	sessionStorage.setItem("checked_passion" , 0);
	    }
	    if(null == sessionStorage.getItem("checked_strength") || isNaN(sessionStorage.getItem("checked_strength")) ){
	    	sessionStorage.setItem("checked_strength" , 0);
	    }
	    if(null == sessionStorage.getItem("checked_value") || isNaN(sessionStorage.getItem("checked_value")) ){
	    	sessionStorage.setItem("checked_value" , 0);
	    }	
	    
		var afterSketch1Model = Spine.Model.sub();
		afterSketch1Model.configure("/user/saveAfterSketchBasic", "jobReferenceNo", 
				"jsonStringFirstPage", "createdBy", "idChecked", "jobTitle", 
				"checkedElems", "passionCount", "valueCount", "strengthCount", 
				"checked_passion", "checked_strength", "checked_value", 
				"divInitialVal", "timeSpents", "profileId", "linkClicked", 
				"rowId", "isPrevious", "jctUserId");
		
		afterSketch1Model.extend( Spine.Model.Ajax );

		//Populate the model with data to transfer
		var modelPopulator = new afterSketch1Model({  
			jobReferenceNo: sessionStorage.getItem("jobReferenceNo"), 
			jsonStringFirstPage: jsonString,
			createdBy: sessionStorage.getItem("jctEmail"),
			idChecked: strengthHTML+valueHTML+passionHTML,
			jobTitle: sessionStorage.getItem("jobTitle"),
			checkedElems: checkedElems,
			passionCount: sessionStorage.getItem("passion_count"),
			valueCount: sessionStorage.getItem("value_count"),
			strengthCount: sessionStorage.getItem("strength_count"),
			checked_passion: sessionStorage.getItem("checked_passion"),
			checked_strength: sessionStorage.getItem("checked_strength"),
			checked_value: sessionStorage.getItem("checked_value"),
			divInitialVal: JSON.parse(sessionStorage.getItem("div_intial_value")),
			timeSpents: timeSpent,
			profileId: sessionStorage.getItem("profileId"),
			linkClicked: sessionStorage.getItem("linkClicked"),
			rowId: sessionStorage.getItem("rowIdentity"),
			isPrevious: decider,
			jctUserId: sessionStorage.getItem("jctUserId")
		});
		modelPopulator.save(); //POST
		afterSketch1Model.bind("ajaxError", function(record, xhr, settings, error) {			
		});
		
	    afterSketch1Model.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			if(sessionStorage.getItem("isLogout") == "N"){
				if (null != obj.pageOneJson) {
					ClientSession.set("total_json_obj",obj.pageOneJson);
				}
				if((obj.bsJson != "none") && (null != obj.bsJson)) {
					sessionStorage.setItem("total_json_add_task", obj.bsJson); //Job reference
					var myJsonObj = JSON.parse(obj.bsJson);
					//If no jsons for after sketch then the count should relate to only before case else After sketch
					if((null != obj.afterSketch2TotalJsonObj) && (obj.afterSketch2TotalJsonObj != "")){
						sessionStorage.setItem("total_json_after_sketch_final", obj.afterSketch2TotalJsonObj);
						sessionStorage.setItem("total_count", obj.asTotalCount);
						sessionStorage.setItem("divHeight", obj.divHeight);
						sessionStorage.setItem("divWidth", obj.divWidth);
					} else {
						sessionStorage.setItem("total_count", myJsonObj.length+1);
					}
					sessionStorage.setItem("div_intial_value", obj.initialJson);
				}
				if (decider == "Prev") {
					sessionStorage.setItem("pageSequence", 2);
					window.location = "Questionaire.jsp";
				} else {
					sessionStorage.setItem("pageSequence", 4);
					window.location = "afterSketch2.jsp";
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
					lastActivePage: "/user/view/afterSketch1.jsp"
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
		});
}

/**
 * Function to store the DivId of the checked chekboxes  
 * @param null
 */
function countCheckedDiv() {
	var totalCkecked = [];	
	var idx = 0;		
	
	for (var i=0;i<document.getElementsByName('strengthList').length;i++){
		var unitJ = {};	
		unitJ["divId"] = document.getElementsByName('strengthList')[i].id;	
		unitJ["divIdChecked"] = document.getElementsByName('strengthList')[i].checked;					
		totalCkecked[idx++] = unitJ;
	}
	
	for (var i=0;i<document.getElementsByName('valueList').length;i++){
		var unitJ = {};	
		unitJ["divId"] = document.getElementsByName('valueList')[i].id;	
		unitJ["divIdChecked"] = document.getElementsByName('valueList')[i].checked;					
		totalCkecked[idx++] = unitJ;
	}
	
	for (var i=0;i<document.getElementsByName('passionList').length;i++){
		var unitJ = {};	
		unitJ["divId"] = document.getElementsByName('passionList')[i].id;	
		unitJ["divIdChecked"] = document.getElementsByName('passionList')[i].checked;					
		totalCkecked[idx++] = unitJ;
	}
	ClientSession.set("checked_element",JSON.stringify(totalCkecked));
}

/**
 * Function to store the json string  
 * while navigate to prevoius page
 * @param null
 */
function backToPreviousPage() {
	sessionStorage.setItem("myPage", "BS");
	ClientSession.set("checked_strength", checkedStrength);
	ClientSession.set("checked_value", checkedValue);	
	ClientSession.set("checked_passion", checkedPassion);
	sessionStorage.setItem("as1View","A");
	sessionStorage.setItem("pageSequence", 2);
	 backToPrevious = 1;
		var cntStrength = 1;
		var totalJson = [];	
		var idx = 0;		
		var divStrength = document.getElementById("divStrength_" + cntStrength);
		while(divStrength != null) {
				var unitJson = {};		
				unitJson["divId"] = document.getElementById("divStrength_" + cntStrength).id;	
				unitJson["textId"] = document.getElementById("textStrength_" + cntStrength).id;	
				//unitJson["divValue"] = document.getElementById("textStrength_"+ cntStrength).innerHTML;	
				unitJson["divValue"] = $("#divStrength_" + cntStrength).find('p').text(); 
				unitJson["style"] = document.getElementById("divStrength_" + cntStrength).getAttribute("style");
				unitJson["positionLeft"] = document.getElementById("divStrength_" + cntStrength).style.left;
				unitJson["positionTop"] = document.getElementById("divStrength_" + cntStrength).style.top;		
				totalJson[idx++] = unitJson;		
				divStrength = document.getElementById("divStrength_" + ++cntStrength);		
		}	
		
		var cntValue = 1;
		var divValue = document.getElementById("divValue_" + cntValue);
		while(divValue != null) {
			var unitJson = {};		
			unitJson["divId"] = document.getElementById("divValue_" + cntValue).id;	
			unitJson["textId"] = document.getElementById("textValue_"+ cntValue).id;		
			//unitJson["divValue"] = document.getElementById("textValue_"+ cntValue).innerHTML;	
			unitJson["divValue"] = $("#divValue_" + cntValue).find('p').text(); 
			unitJson["style"] = document.getElementById("divValue_" + cntValue).getAttribute("style");
			unitJson["positionLeft"] = document.getElementById("divValue_" + cntValue).style.left;
			unitJson["positionTop"] = document.getElementById("divValue_" + cntValue).style.top;		
			totalJson[idx++] = unitJson;
			divValue = document.getElementById("divValue_" + ++cntValue);
		}

		var cntPassion = 1;
		var divPassion = document.getElementById("divPassion_" + cntPassion);
		while(divPassion != null) {
			var unitJson = {};		
			unitJson["divId"] = document.getElementById("divPassion_" + cntPassion).id;	
			unitJson["textId"] = document.getElementById("textPassion_"+ cntPassion).id;		
			//unitJson["divValue"] = document.getElementById("textPassion_"+ cntPassion).innerHTML;	
			unitJson["divValue"] = $("#divPassion_" + cntPassion).find('p').text(); 
			unitJson["style"] = document.getElementById("divPassion_" + cntPassion).getAttribute("style");
			unitJson["positionLeft"] = document.getElementById("divPassion_" + cntPassion).style.left;
			unitJson["positionTop"] = document.getElementById("divPassion_" + cntPassion).style.top;		
			totalJson[idx++] = unitJson;
			divPassion = document.getElementById("divPassion_" + ++cntPassion);
		}
		ClientSession.set("total_json_obj",JSON.stringify(totalJson));
		ClientSession.set("back_to_previous", backToPrevious);
		countCheckedDiv();
		
		//Fetch the questionaire
		var navigation = Spine.Model.sub();
		navigation.configure("/user/navigation/goPrevious", "navigationSeq", "jobReferenceNos", "profileId", "linkClicked", "rowId");
		navigation.extend( Spine.Model.Ajax );
		//Populate the model
			var replyModel = new navigation({
				navigationSeq: "2",
				jobReferenceNos: sessionStorage.getItem("jobReferenceNo"),
				profileId: sessionStorage.getItem("profileId"),
				linkClicked: sessionStorage.getItem("linkClicked"),
				rowId: sessionStorage.getItem("rowIdentity") 
			});
			replyModel.save(); //POST
			navigation.bind("ajaxError", function(record, xhr, settings, error) {
				alertify.alert("Unable to connect to the server.");
				return false;
			});
			
			navigation.bind("ajaxSuccess", function(record, xhr, settings, error) {
				var jsonStr = JSON.stringify(xhr);
				var obj = jQuery.parseJSON(jsonStr);
				if(sessionStorage.getItem("isLogout") == "N"){
					sessionStorage.setItem("questions", obj.list);
					sessionStorage.setItem("snapShotURLS", obj.jctBaseString);
					sessionStorage.setItem("isCompleted", obj.isCompleted);
					storeAfterSketchOne(totalJson, "Prev");
					//window.location = "Questionaire.jsp";
				}
			});
} 
/******************* CALL EVERYTIME WHEN THE PAGE APPEARS ************************/
/**
 * Function to populate Strength Element
 * @param null
 */
function populateStrengths(){
	var outputArea = document.getElementById('strengthPopulateDiv');
	var data = sessionStorage.getItem("strength");	
	var splitData = data.split(",");
	for(var index=0; index<splitData.length; index++){
		var innerSplit = splitData[index].split("\":\"");
		var id = innerSplit[0].replace("\"", "");
		if(index == 0){
			id = id.replace("{", "");
		}
		var fullValue = innerSplit[1].replace("\"", "");
		var valueSplit = fullValue.split("+++");
		var value = valueSplit[0];
		var description= valueSplit[1].replace("~~", ",");
		description = description.replace("~~", ",");
		if(splitData.length-index == 1){
			value = value.replace("}", "");
			description = description.replace("}", "");
		}
		description = value + ": "+ description;
		/* popover added to show the strength description on mouse hover
		 * Dated : 16.07.2014
		 */
		strengthHTML = strengthHTML + "<div class='col-md-12 item_value ckecked_strength' data-toggle='tooltip' data-placement='top' title='' data-original-title='"+description+"'><input type='checkbox' id='"+id+"' value='"+value+"' name='strengthList' onclick='checkStrength(this)'> &nbsp;"+value+"</div>";
		//strengthHTML = strengthHTML + "<div class='col-md-12 item_value ckecked_strength'><input type='checkbox' id='"+id+"' value='"+value+"' name='strengthList'  onclick='checkStrength(this)'> &nbsp;"+value+"</div>";
	}	
	outputArea.innerHTML = strengthHTML;
	$(".ckecked_strength").tooltip(); 	
}

/**
 * Function to populate Value Element
 * @param null
 */
function populateValue(){
	var outputArea = document.getElementById('valuePopulateDiv');
	var data = sessionStorage.getItem("value");
	var splitData = data.split(",");
	for(var index=0; index<splitData.length; index++){
		var innerSplit = splitData[index].split("\":\"");

		var id = innerSplit[0].replace("\"", "");
		if(index == 0){
			id = id.replace("{", "");
		}
		var fullValue = innerSplit[1].replace("\"", "");
		var valueSplit = fullValue.split("+++");
		var value = valueSplit[0];
		var description= valueSplit[1].replace("~~", ",");
		description = description.replace("~~", ",");
		//var value = innerSplit[1].replace("\"", "");
		if(splitData.length-index == 1){
			value = value.replace("}", "");
			description = description.replace("}", "");
		}
		description = value + ": "+ description;
		/* popover added to show the value description on mouse hover
		 * Dated : 16.07.2014
		 */
		valueHTML = valueHTML + "<div class='col-md-12 item_value ckecked_value' data-toggle='tooltip' data-placement='top' title='' data-original-title='"+description+"'><input type='checkbox' id='"+id+"' value='"+value+"' name='valueList' onclick='checkValue(this)' > &nbsp;"+value+"</div>";
		//valueHTML = valueHTML + "<div class='col-md-12 item_value ckecked_value'><input type='checkbox' id='"+id+"' value='"+value+"' name='valueList'  onclick='checkValue(this)' > &nbsp;"+value+"</div>";
	}
	outputArea.innerHTML = valueHTML;
	$(".ckecked_value").tooltip(); 
}

/**
 * Function to populate Passion Element
 * @param null
 */
function populatePassion(){
	var outputArea = document.getElementById('passionPopulateDiv');
	var data = sessionStorage.getItem("passion");
	var splitData = data.split(",");
	for(var index=0; index<splitData.length; index++){
		var innerSplit = splitData[index].split("\":\"");
		var id = innerSplit[0].replace("\"", "");
		if(index == 0){
			id = id.replace("{", "");
		}
		var fullValue = innerSplit[1].replace("\"", "");
		var valueSplit = fullValue.split("+++");
		var value = valueSplit[0];
		var description= valueSplit[1].replace("~~", ",");
		description = description.replace("~~", ",");
		//var value = innerSplit[1].replace("\"", "");
		if(splitData.length-index == 1){
			value = value.replace("}", "");
			description = description.replace("}", "");
		}
		description = value + ": "+ description;
		/* popover added to show the passion description on mouse hover
		 * Dated : 16.07.2014
		 */
		passionHTML = passionHTML + "<div class='col-md-12 item_value ckecked_passion' data-toggle='tooltip' data-placement='top' title='' data-original-title='"+description+"'><input type='checkbox' id='"+id+"' value='"+value+"' name='passionList' onclick='checkPassion(this)' > &nbsp;"+value+"</div>";
		//passionHTML = passionHTML + "<div class='col-md-12 item_value ckecked_passion'><input type='checkbox' id='"+id+"' value='"+value+"' name='passionList' onclick='checkPassion(this)' > &nbsp;"+value+"</div>";
	}
	outputArea.innerHTML = passionHTML;
	$(".ckecked_passion").tooltip(); 
}


/**
 * check if the json object is not null
 * parse the json string and draw the element
 * @param jsonString
 */
if(ClientSession.get("back_to_previous")==1){
	populateStrengths();
	populateValue();
	populatePassion();
	$('#strength_pageWrap').empty();
	$('#value_pageWrap').empty();
	$('#passion_pageWrap').empty();
	totalJson2 = JSON.parse(ClientSession.get("total_json_obj"));
	if(totalJson2!=null){
		var cnt = totalJson2.length;
		for (var i = 0; i<cnt;i++){
			var divIda = totalJson2[i]["divId"];
			var textIda = totalJson2[i]["textId"];	
			var divValuea = totalJson2[i]["divValue"];				
			var positionLefta = totalJson2[i]["positionLeft"];
			var positionTopa = totalJson2[i]["positionTop"];	
			var divStyle = totalJson2[i]["style"];
			if((divStyle == null) || (divStyle == "null") || (divStyle.search('block') != -1)){
				divStyle = "block;";
			}else{
				divStyle = "none;";
			}
			if(divIda.search("divStrength_") != -1){
				/* popover added to show the strength description on mouse hover
				 * Dated : 16.07.2014
				 */
				var description = getAttrDecription(divValuea, "STR");	
				//alert(divValuea);
				var $element2 = $('<div id="' + divIda +'" class="plus-sign commonDrag" data-toggle="popover" data-placement="bottom" title="" data-original-title="'+description+'" style="position: relative; top:'+positionTopa+'; left:'+positionLefta+'; display:'+divStyle+'"><img src="../img/strength.png" width="110" height="110" style="no-repeat 0 9px;"/><div class="plus-sign_span" id="'+textIda+'"><p>'+divValuea+'</p></div>');
				$("#strength_pageWrap").append($element2);
				$element2.popover(); 
			} else if(divIda.search("divValue_") != -1){	
				/* popover added to show the value description on mouse hover
				 * Dated : 16.07.2014
				 */
				var description = getAttrDecription(divValuea, "VAL");	
				var $element3 = $('<div id="' + divIda +'" class="oval commonDrag" data-toggle="popover" data-placement="bottom" title="" data-original-title="'+description+'" style="position: relative; top:'+positionTopa+'; left:'+positionLefta+'; display:'+divStyle+'"><img src="../img/value.png" width="130" height="100" style="no-repeat 0 9px;"/><div class="oval_span" id="'+textIda+'"><p>'+divValuea+'</p></div>');
				$("#value_pageWrap").append($element3);
				 $element3.popover(); 
			} else if(divIda.search("divPassion_") != -1){
				/* popover added to show the passion description on mouse hover
				 * Dated : 16.07.2014
				 */
				var description = getAttrDecription(divValuea, "PAS");	
				var $element4 = $('<div id="' + divIda +'" class="up-triangle commonDrag" data-toggle="popover" data-placement="bottom" title="" data-original-title="'+description+'" style="position: relative; top:'+positionTopa+'; left:'+positionLefta+'; display:'+divStyle+'"><img src="../img/passion.png" width="112" height="110" style="no-repeat 0 9px;"/><div class="up-triangle_span" id="'+textIda+'"><p>'+divValuea+'</p></div>');				
				$("#passion_pageWrap").append($element4);
				 $element4.popover(); 
			}	
		}  
	}
	totalJsonChk = JSON.parse(ClientSession.get("checked_element"));
	if(totalJsonChk!=null){
		var cnt = totalJsonChk.length;
		for (var i = 0; i<cnt;i++){
			var divId = totalJsonChk[i]["divId"];
			var divIdChk = totalJsonChk[i]["divIdChecked"];
			if(divIdChk==true){
				document.getElementById(divId).checked = true;
			}
		}
	}
}
/******************* This is called when the page is loaded****************/
$(document).ready(function() {
	var insDisplay = sessionStorage.getItem("as1View");
	if(null == insDisplay || insDisplay == "" || insDisplay == "D"){
		document.getElementById('panel').style.display = "block";
		document.getElementById('drp').setAttribute("class", "btn-slide active");
	} else {
		document.getElementById('panel').style.display = "none";
		document.getElementById('drp').setAttribute("class", "btn-slide");
	}
	if (!ClientSession.get("back_to_previous")){
		populateStrengths();
		populateValue();
		populatePassion();
	}
	$('#strength_pageWrap').empty();
	$('#value_pageWrap').empty();
	$('#passion_pageWrap').empty();
	totalJson2 = JSON.parse(ClientSession.get("total_json_obj"));
		if(totalJson2!=null){
		var cnt = totalJson2.length;
		for (var i = 0; i<cnt;i++){
			var divIda = totalJson2[i]["divId"];
			var textIda = totalJson2[i]["textId"];	
			var divValuea = totalJson2[i]["divValue"];	
			var positionLefta = totalJson2[i]["positionLeft"];
			var positionTopa = totalJson2[i]["positionTop"];	
			var divStyle = totalJson2[i]["style"];
			if((divStyle == null) || (divStyle == "null") || (divStyle.search('block') != -1)){
				divStyle = "block;";
			}else{
				divStyle = "none;";
			}
		if(divIda.search("divStrength_") != -1){
			/* popover added to show the strength description on mouse hover
			 * Dated : 16.07.2014
			 */
			var description = getAttrDecription(divValuea, "STR");	
			//alert(divValuea);
			var $element2 = $('<div id="' + divIda +'" class="plus-sign commonDrag" data-toggle="popover" data-placement="bottom" title="" data-original-title="'+description+'" style="position: relative; top:'+positionTopa+'; left:'+positionLefta+'; display:'+divStyle+'"><img src="../img/strength.png" width="110" height="110" style="no-repeat 0 9px;"/><div class= "plus-sign_span" id="'+textIda+'"><p>'+divValuea+'</p></div>');
			$("#strength_pageWrap").append($element2);
			 $element2.popover(); 
		} else if(divIda.search("divValue_") != -1){	
			/* popover added to show the value description on mouse hover
			 * Dated : 16.07.2014
			 */
			var description = getAttrDecription(divValuea, "VAL");	
			var $element3 = $('<div id="' + divIda +'" class="oval commonDrag" data-toggle="popover" data-placement="bottom" title="" data-original-title="'+description+'" style="position: relative; top:'+positionTopa+'; left:'+positionLefta+'; display:'+divStyle+'"><img src="../img/value.png" width="130" height="100" style="no-repeat 0 9px;"/><div class= "oval_span" type="text" id="'+textIda+'"><p>'+divValuea+'</p></div>');
			$("#value_pageWrap").append($element3);
			 $element3.popover(); 
		} else if(divIda.search("divPassion_") != -1){
			/* popover added to show the passion description on mouse hover
			 * Dated : 16.07.2014
			 */
			var description = getAttrDecription(divValuea, "PAS");	
			var $element4 = $('<div id="' + divIda +'" class="up-triangle commonDrag" data-toggle="popover" data-placement="bottom" title="" data-original-title="'+description+'" style="position: relative; top:'+positionTopa+'; left:'+positionLefta+'; display:'+divStyle+'"><img src="../img/passion.png" width="112" height="110" style="no-repeat 0 9px;"/><div class ="up-triangle_span" type="text" id="'+textIda+'"><p>'+divValuea+'</p></div>');
			$("#passion_pageWrap").append($element4);
			 $element4.popover(); 
		}
		}  
	}
	totalJsonChk = JSON.parse(ClientSession.get("checked_element"));
	if(totalJsonChk!=null){
		var cnt = totalJsonChk.length;
		for (var i = 0; i<cnt;i++){
			var divId = totalJsonChk[i]["divId"];
			var divIdChk = totalJsonChk[i]["divIdChecked"];
			if(divIdChk==true){
				document.getElementById(divId).checked = true;
			}
		}
	}
	populateFrmSession();
});

/**
 * Function add to find the checkboxes checked
 * @param null
 */
function populateFrmSession() {
	var checkedStrength = ClientSession.get("checked_strength");
	if(checkedStrength == 4){		
			$('.ckecked_strength input[type=checkbox]').not(':checked').attr("disabled",true);
		} else {
			$('.ckecked_strength input[type=checkbox]').not(':checked').attr("disabled",false);
		}
		
		
		var checkedValue = ClientSession.get("checked_value");
		if(checkedValue == 4){		
			$('.ckecked_value input[type=checkbox]').not(':checked').attr("disabled",true);
		} else {
			$('.ckecked_value input[type=checkbox]').not(':checked').attr("disabled",false);
		}
		
		var checkedPassion = ClientSession.get("checked_passion");
		if(checkedPassion == 4){		
			$('.ckecked_passion input[type=checkbox]').not(':checked').attr("disabled",true);
		} else {
			$('.ckecked_passion input[type=checkbox]').not(':checked').attr("disabled",false);
		}
}


/**
 * Function add to find the attribute description
 * against the attribute name
 * @param name
 * @param attrCode
 * Dated : 16.07.2014
 */
function getAttrDecription(name, attrCode) {
	var data = "";
	if (attrCode == "STR") {
		data = sessionStorage.getItem("strength");
	} else if (attrCode == "VAL") {
		data = sessionStorage.getItem("value");
	} else {
		data = sessionStorage.getItem("passion");
	}	
	var splitData = data.split(",");	 
	var description = "";
	for(var index=0; index<splitData.length; index++){
		var innerSplit = splitData[index].split("\":\"");		
		var fullValue = innerSplit[1].replace("\"", "");
		var valueSplit = fullValue.split("+++");
		var value = valueSplit[0];
		if(value == name) {			
			description= valueSplit[1].replace("~~", ",");;					
			description = description.replace("~~", ",");
			if(splitData.length-index == 1){
				value = value.replace("}", "");
				description = description.replace("}", "");
				break;
			}
		}	
	}
	return description;
}


/**
 * This Function save the data in session 
 * while the page refresh
 */
function saveMappingWhileRefresh(){	
		/************* set the value to session for logout **************/
		/*ClientSession.set("checked_strength", checkedStrength);
		ClientSession.set("checked_value", checkedValue);	
		ClientSession.set("checked_passion", checkedPassion);*/
		
		pageNavigation = 1;
		var cntStrength = 1;
		var totalJson = [];	
		var idx = 0;		
		var divStrength = document.getElementById("divStrength_" + cntStrength);
		while(divStrength != null) {
				var unitJson = {};				
				unitJson["divId"] = document.getElementById("divStrength_" + cntStrength).id;	
				unitJson["textId"] = document.getElementById("textStrength_" + cntStrength).id;	
				//unitJson["divValue"] = document.getElementById("textStrength_"+ cntStrength).innerHTML;	
				unitJson["divValue"] = $("#divStrength_" + cntStrength).find('p').text(); 
				unitJson["style"] = document.getElementById("divStrength_" + cntStrength).getAttribute("style");
				unitJson["positionLeft"] = document.getElementById("divStrength_" + cntStrength).style.left;
				unitJson["positionTop"] = document.getElementById("divStrength_" + cntStrength).style.top;	
				totalJson[idx++] = unitJson;		
				divStrength = document.getElementById("divStrength_" + ++cntStrength);		
		}	
		
		var cntValue = 1;
		var divValue = document.getElementById("divValue_" + cntValue);
		while(divValue != null) {
			var unitJson = {};		
			unitJson["divId"] = document.getElementById("divValue_" + cntValue).id;	
			unitJson["textId"] = document.getElementById("textValue_"+ cntValue).id;		
			//unitJson["divValue"] = document.getElementById("textValue_"+ cntValue).innerHTML;
			unitJson["divValue"] = $("#divValue_" + cntValue).find('p').text(); 
			unitJson["style"] = document.getElementById("divValue_" + cntValue).getAttribute("style");
			unitJson["positionLeft"] = document.getElementById("divValue_" + cntValue).style.left;
			unitJson["positionTop"] = document.getElementById("divValue_" + cntValue).style.top;		
			totalJson[idx++] = unitJson;
			divValue = document.getElementById("divValue_" + ++cntValue);
		}

		var cntPassion = 1;
		var divPassion = document.getElementById("divPassion_" + cntPassion);
		while(divPassion != null) {
			var unitJson = {};		
			unitJson["divId"] = document.getElementById("divPassion_" + cntPassion).id;	
			unitJson["textId"] = document.getElementById("textPassion_"+ cntPassion).id;		
			//unitJson["divValue"] = document.getElementById("textPassion_"+ cntPassion).innerHTML;	
			unitJson["divValue"] = $("#divPassion_" + cntPassion).find('p').text(); 
			unitJson["style"] = document.getElementById("divPassion_" + cntPassion).getAttribute("style");
			unitJson["positionLeft"] = document.getElementById("divPassion_" + cntPassion).style.left;
			unitJson["positionTop"] = document.getElementById("divPassion_" + cntPassion).style.top;		
			totalJson[idx++] = unitJson;
			divPassion = document.getElementById("divPassion_" + ++cntPassion);
		}
		ClientSession.set("total_json_obj",JSON.stringify(totalJson));
		ClientSession.set("Page_navigation", pageNavigation);	
		countCheckedDiv();
		/*if(sessionStorage.getItem("isLogout")) {
			storeAfterSketchOne(totalJson, "None");
		} else {
			storeAfterSketchOne(totalJson, "Next");
		}*/
		
}