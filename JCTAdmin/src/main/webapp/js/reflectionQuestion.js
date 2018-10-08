var mainHeight = $(".main_contain").height();
var manipulativeQtn = "";
/**
 * Function add the sub question input box dynamically
 * @param param
 */
function create(param, divStyle) {
var divCount = "";
var i;
if(divStyle != null) {
	if((divStyle.search('block') != -1)) {
	divCount = $('#sunQtnAdditionDiv').children('.single_form_item:visible').length;
	var jsonObj = [];
		var counter = 0;
		for (i = 1; i <= divCount; i += 1) {
		var unitJ = {};	
		var value = document.getElementById('qtnSubDescInptId_'+i).value;
		unitJ["ElementValue"+i] = value;
		jsonObj[counter++] = unitJ;
		}
		sessionStorage.setItem("previous_value",JSON.stringify(jsonObj)); 
	}
	}
    $("#sunQtnAdditionDiv").empty();
    for (i = 1; i <= param; i += 1) {   
    	//var mainQtnOrder = document.getElementById('qtnOrderInptId').value;
    	var subQtnOrder = "";
    	if( i == 1){
    		subQtnOrder = "a";   		
    	} else if( i == 2 ){
    		subQtnOrder = "b";   
    	} else if( i == 3 ){
    		subQtnOrder = "c";   
    	} else if( i == 4 ){
    		subQtnOrder = "d";   
    	}
    	 //$('#sunQtnAdditionDiv').append('<div class="single_form_item" ><div class="col-sm-4 "><label for="inputSubQtn" class="col-sm-12 control-label text_label">Sub Question '+mainQtnOrder+'.'+subQtnOrder+':</label></div><div class="col-sm-6"><input type="text" class="form-control-general-form-redefined" maxlength="50" value="" id="qtnSubDescInptId_'+i+'" onkeypress="return validateDescription(event)" ></div><div class="col-sm-2"></div><div class="clearfix"></div></div');   
    	 $('#sunQtnAdditionDiv').append('<div class="single_form_item" ><div class="col-sm-4 select_option"><label for="inputSubQtn" class="col-sm-12 control-label text_label">Sub Question '+i+':</label></div><div class="col-sm-6 select_option_radio"><input type="text" class="form-control-general-form-redefined" maxlength="50" value="" id="qtnSubDescInptId_'+i+'" onkeypress="return validateDescription(event)" ></div><div class="col-sm-2"></div><div class="clearfix"></div></div');   
    	 	
			
		
    	 /***************** To Disable the paste functionality ***************/
    	   /* $("#qtnSubDescInptId_"+i).bind("paste",function(e) {
    	        e.preventDefault();
    	  });*/
    }

	if(sessionStorage.getItem("previous_value") != null){
		var jsonObj = JSON.parse(sessionStorage.getItem("previous_value"));
		var count = jsonObj.length;
			for (var j = 0; j<count; j++){	
			var value = jsonObj[j]["ElementValue"+(j+1)];
			var div = document.getElementById('qtnSubDescInptId_'+(j+1));
			    if (div) {
				document.getElementById('qtnSubDescInptId_'+(j+1)).value = value;
				}
			
			}
		
		}
}

$('#inputNoOFSubQtn').change(function () {
var divStyle = document.getElementById("sunQtnAdditionDiv").getAttribute("style");
	/****** JIRA ID - JCTP-8 **********/
	var subQtnAddDiv = document.getElementById('sunQtnAdditionDiv');
	subQtnAddDiv.setAttribute("style", "display:block");	
	/***** ENDED ******/
    create($(this).val(), divStyle);
});

/**
 * Function calls when the page is loaded 
 * to populate the reflection question data
 */
$(document).ready(function() {	
	/*equalHeight($(".col"));*/
	document.getElementById('addHeaderDiv').innerHTML = "Add Reflection Question / Action Plan";
	disableAllfield();
	fetchReflectionQtnList('R', 0, "BS");
	sessionStorage.removeItem("previous_value"); 
	/***************** To Disable the paste functionality ***************/
    /*$("#qtnDescInptId").bind("paste",function(e) {
        e.preventDefault();
    });*/
});


/**
 * Function to view all the reflection question 
 * from database while landing to the page
 * @param objval
 */
function fetchReflectionQtnList(objVal, profileId, relatedPage) {	
	var userProf = Spine.Model.sub();
	userProf.configure("/admin/contentconfig/populateExistingRefQtn", "profileId", "relatedPage");
	userProf.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userProf({  
		profileId: profileId,
		relatedPage: relatedPage
	});
	modelPopulator.save(); //POST
	userProf.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	userProf.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if(statusCode == 200) {
			populateQtnDropDown(obj.noOfSubQtn);
			if(profileId == 0){
				populateDropDown(obj.userProfileMap);		
			}	
			if( null == objVal|| objVal == "" || objVal == 'R') {
				populateReflectionQtnTable(obj.existingRefQtnList, profileId);
			} else {
				populateActionPlanTable(obj.existingRefQtnList, profileId);
			}
			
		} else if(statusCode == 201) {
			document.getElementById("existingRefQtnListId").innerHTML = "<div align='center'>"+obj.statusDesc+"</div>";
		} else if(statusCode == 500) {
			//Show error message
		}
	});
}


/**
 * Function to show the reflection question 
 * on table view
 * @param refQtnListData
 * @param action
 */
function populateReflectionQtnTable(refQtnListData, profileIdVal) {	
	var memory = "";
	var profileId = "";
	var questionOrder = 0;
	var tableDiv = document.getElementById("existingRefQtnListId");
	if (null != refQtnListData) {
		var headingString = "<table width='94%' id='refQtnTableId' border=1 align='center' bordercolor='#78C0D3' class='tablesorter'><thead class='tab_header'><tr><th width='11%'><b>SL. No.</b></th><th width='16%'><b>Profile Name</b></th><th><b>Question</b></th><th><b>Sub Question</b></th><th width='12%'><b>Action</b></th></tr></thead><tbody>";
		var splitRec = refQtnListData.split("$$$");
		if(splitRec.length <= 2){
			document.getElementById("chgOrderBtnId").disabled = true;
		} 
		var counter = 1;
		for(var outerIndex = 0; outerIndex < splitRec.length-1; outerIndex++){	
			var trColor = "";
			if(outerIndex % 2 == 0) {
				trColor = "#D2EAF0";
			} else {
				trColor = "#F1F1F1";
			}
			var objSplit = splitRec[outerIndex].split("~");		
			if(objSplit[2] == 'BS') {
				headingString = headingString + "<tr align='center' class='row_width' bgcolor='"+trColor+"'><td>&nbsp; "+counter+".</td>";
				headingString = headingString + "<td align='center'>"+objSplit[3]+"</td>";	
				if(memory != objSplit[0] || profileId != objSplit[6]){
					questionOrder ++;
					memory = objSplit[0];
					profileId = objSplit[6];
					headingString = headingString + "<td align='left'>"+objSplit[0]+"</td>";			
				} else {				
					headingString = headingString + "<td align='center'>---</td>";
				}					
				headingString = headingString + "<td align='left'>"+objSplit[1]+"</td>";		
				//headingString = headingString + "<td class='table_col_txt_style'>&nbsp;&nbsp;&nbsp;<a class='delete_style' href='#' id='deleteBS_"+counter+"' onClick='deleteRefQtn(\""+objSplit[4]+"\", \""+objSplit[0]+"\",  \""+objSplit[6]+"\")'><img src = '../img/delete.png' /></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a class='edit_style' href='#' onClick='editRefQtn(\""+objSplit[4]+"\", \""+objSplit[0]+"\", \""+objSplit[3]+"\", \""+objSplit[1]+"\", \""+objSplit[6]+"\", \""+counter+"\")'><img src = '../img/edit.png' /></a></td>";
				headingString = headingString + "<td class='table_col_txt_style'><a class='edit_style' href='#' onClick='editRefQtn(\""+objSplit[4]+"\", \""+objSplit[0]+"\", \""+objSplit[3]+"\", \""+objSplit[1]+"\", \""+objSplit[6]+"\", \""+counter+"\")'><img src = '../img/edit.png' /></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a class='delete_style' href='#' id='deleteBS_"+counter+"' onClick='deleteRefQtn(\""+objSplit[4]+"\", \""+objSplit[0]+"\",  \""+objSplit[6]+"\")'><img src = '../img/delete.png' /></a></td>";
				headingString = headingString + "</tr>";
				counter = counter + 1;
			} 
			if(questionOrder < objSplit[7]) {
				questionOrder = objSplit[7];
			}					
		}		
		headingString = headingString + "</tbody></table>";
		tableDiv.innerHTML = headingString;
		//new SortableTable(document.getElementById('refQtnTableId'), 1);
		$("table").tablesorter(); 
	} else {		
		document.getElementById("chgOrderBtnId").disabled = true;
		tableDiv.innerHTML = "<div align='center'><img src='../img/no-record.png'><br />No Reflection Question To Display</div>";
	}	
	
	//equalHeight($(".col"));
}

/**
 * Function to show the action plan
 * on table view
 * @param refQtnListData
 */
function populateActionPlanTable(refQtnListData, profileIdVal) {
	var memory = "";
	var profileId = "";
	var questionOrder = 0;
	var mainQtnOrder = 0;
	var subQtnOrder = 0;
	var tableDiv = document.getElementById("existingRefQtnListId");
	if (null != refQtnListData) {
		var headingString = "<table width='94%' id='actionPlanTableId' border=1 align='center' bordercolor='#78C0D3' class='tablesorter'><thead class='tab_header'><tr><th width='11%'><b>SL. No.</b></th><th width='16%'><b>Profile Name</b></th><th><b>Question</b></th><th><b>Sub Question</b></th><th width='12%'><b>Action</b></th></tr></thead><tbody>";
		var splitRec = refQtnListData.split("$$$");
		if(splitRec.length <= 2){
			document.getElementById("chgOrderBtnId").disabled = true;
		} 
		var counter = 1;
		for(var outerIndex = 0; outerIndex < splitRec.length-1; outerIndex++){	
			var trColor = "";
			if(outerIndex % 2 == 0) {
				trColor = "#D2EAF0";
			} else {
				trColor = "#F1F1F1";
			}
			var objSplit = splitRec[outerIndex].split("~");	
			if(objSplit[2] == 'AS') {
				headingString = headingString + "<tr align='center' class='row_width' bgcolor='"+trColor+"'><td>&nbsp; "+counter+".</td>";
				headingString = headingString + "<td align='center'>"+objSplit[3]+"</td>";	
				if(memory != objSplit[0] || profileId != objSplit[6]){
					questionOrder ++;
					memory = objSplit[0];
					profileId = objSplit[6];
					headingString = headingString + "<td align='left'>"+objSplit[0]+"</td>";			
				} else {				
					headingString = headingString + "<td align='center'>---</td>";
				}					
				headingString = headingString + "<td align='left'>"+objSplit[1]+"</td>";				
				//headingString = headingString + "<td class='table_col_txt_style'>&nbsp;&nbsp;&nbsp;<a class='delete_style' href='#' id='deleteAS_"+counter+"' onClick='deleteActionPlan(\""+objSplit[4]+"\", \""+objSplit[0]+"\",  \""+objSplit[6]+"\")'><img src = '../img/delete.png' /></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a class='edit_style' href='#' onClick='editActionPlan(\""+objSplit[4]+"\", \""+objSplit[0]+"\", \""+objSplit[3]+"\", \""+objSplit[1]+"\", \""+objSplit[6]+"\", \""+counter+"\")'><img src = '../img/edit.png' /></a></td>";
				headingString = headingString + "<td class='table_col_txt_style'><a class='edit_style' href='#' onClick='editActionPlan(\""+objSplit[4]+"\", \""+objSplit[0]+"\", \""+objSplit[3]+"\", \""+objSplit[1]+"\", \""+objSplit[6]+"\", \""+counter+"\")'><img src = '../img/edit.png' /></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a class='delete_style' href='#' id='deleteAS_"+counter+"' onClick='deleteActionPlan(\""+objSplit[4]+"\", \""+objSplit[0]+"\",  \""+objSplit[6]+"\")'><img src = '../img/delete.png' /></a></td>";
				headingString = headingString + "</tr>";
				counter = counter + 1;
			} 
			if(questionOrder < objSplit[7]) {
				questionOrder = objSplit[7];
			}					
		}		
		headingString = headingString + "</tbody></table>";
		tableDiv.innerHTML = headingString;
		//new SortableTable(document.getElementById('actionPlanTableId'), 1);
		$("table").tablesorter(); 
	} else {		
		document.getElementById("chgOrderBtnId").disabled = true;
		tableDiv.innerHTML = "<div align='center'><img src='../img/no-record.png'><br />No Action Plan To Display</div>";
	}	
	//alert($(".main_contain").height());
	equalHeight($(".col"));
}

/**
 * Function add to check the  
 * selected radio button
 * @param obj
 */
function equalHeight(group) {
	   tallest = 0;
	   mmh = 15;
	   group.each(function() {
	      thisHeight = $(this).height();
	      //alert(thisHeight);
	      if(thisHeight > tallest) {
	         tallest = thisHeight;
	      }
	   });
	   group.height(tallest);
	}



function checkSelectedOption(obj) {		
	var subQtnDiv = document.getElementById('sudQuestionDiv');
	var refQtnAddDiv = document.getElementById('refQuestionAddDiv');
	var actnPlanAddDiv = document.getElementById('actionPlanAddDiv');
	if(obj.value == 'A'){		
		refQtnAddDiv.setAttribute("style", "display:none");
		actnPlanAddDiv.setAttribute("style", "display:block");
		document.getElementById('existing_list_Id').innerHTML = "Existing Action Plan";
		$(".chgOrderDivButton").addClass("chgOrderBtnActionClass");
		document.getElementById("chgOrderBtnId").disabled = true;
		fetchReflectionQtnList(obj.value, 0, "AS");
		document.getElementById('qtnDescInptId').value = "";
		document.getElementById('qtnDescInptId').setAttribute("disabled", "true");				
		document.getElementById('inputNoOFSubQtn').disabled = true;	
		
	} else {
		$(".chgOrderDivButton").removeClass("chgOrderBtnActionClass");
		refQtnAddDiv.setAttribute("style", "display:block");
		actnPlanAddDiv.setAttribute("style", "display:none");
		document.getElementById('existing_list_Id').innerHTML = "Existing Reflection Question";
		document.getElementById("chgOrderBtnId").disabled = true;
		fetchReflectionQtnList(obj.value, 0, "BS");
		equalHeight($(".col"));
		document.getElementById('qtnDescInptId').value = "";
		document.getElementById('qtnDescInptId').setAttribute("disabled", "true");				
		document.getElementById('inputNoOFSubQtn').disabled = true;	
	}
	create("", null);
	//equalHeight($(".col"));
}



/**
 * Function add to populate
 * the user profile drop-down value
 * @param userProfileMap
 */
function populateDropDown(userProfileMap){
	document.getElementById("inputUserProfileInpt").innerHTML = "";	
	var userProfileSelect = document.getElementById("inputUserProfileInpt");
	var defaultOpn = document.createElement("option");
	defaultOpn.text = "Select User Profile";
	defaultOpn.value = "";
	defaultOpn.className = "form-control-general";
	userProfileSelect.add(defaultOpn, null);
	for (var key in userProfileMap) {				
		var option = document.createElement("option");
		option.text = userProfileMap[key];
	    option.value = key;
	    option.className = "form-control-general";
	    try {
	    	userProfileSelect.add(option, null); //Standard 
	    }catch(error) {
	    	//regionSelect.add(option); // IE only
	    }
	}
}

/**
 * Function to add the
 * reflection question to the database
 * @param null
 */
function saveRefQtn() {	
	var relatedPage = "BS";
	var builder = "";
	var noOFSubQtn = document.getElementById("inputNoOFSubQtn").value;	
	for(var i = 1; i <= noOFSubQtn; i++){
		if(null != document.getElementById("qtnSubDescInptId_"+i).value.trim() && document.getElementById("qtnSubDescInptId_"+i).value.trim() !=""){
			builder = builder + document.getElementById("qtnSubDescInptId_"+i).value +"~~";			
		} 				
	}
	/*for(var i = 1; i <= noOFSubQtn; i++){
		if(null != document.getElementById("qtnSubDescInptId_"+i).value.trim() && document.getElementById("qtnSubDescInptId_"+i).value.trim() !=""){
			builder = builder + document.getElementById("qtnSubDescInptId_"+i).value +"~~";			
		} else {
			alertify.alert("Please enter sub question value.");
			return false;
		}					
	}*/
	var refQtnText = document.getElementById("qtnDescInptId").value;
	var userProfile = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value;
	var userProfileText = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].text;
	//var questionOrder = document.getElementById("qtnOrderInptId").value;
	if(validateActionPlan(refQtnText, builder)) {
		var newUserProf = Spine.Model.sub();
		newUserProf.configure("/admin/contentconfig/saveRefQtn", "refQtnDesc", "subQtn", "userProfileVal", "userProfileId","relatedPage", "createdBy");
		newUserProf.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new newUserProf({  
			refQtnDesc: refQtnText,
			subQtn: builder,
			userProfileVal: userProfileText,
			userProfileId: userProfile,
			relatedPage: relatedPage,
			createdBy: sessionStorage.getItem("jctEmail")
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
				alertify.alert("Reflection Question has been successfully created.");
				fetchReflectionQtnList('R', userProfile, "BS");
												
				document.getElementById('qtnDescInptId').value = "";	
				document.getElementById('qtnSubDescEditId').value = "";		
				document.getElementById('inputNoOFSubQtn').value = "";					
				//document.getElementById('qtnDescInptId').setAttribute("disabled", "true");				
				//document.getElementById('inputNoOFSubQtn').disabled = true;	
				create("", null);
			} else if(statusCode == 600) {
				alertify.alert("Reflection Question already exists.");
			} else if(statusCode == 602) {
				alertify.alert("One profile should have maximum of 6 reflection question.");
			} else if(statusCode == 500) {
				// Error Message
			}
		});
	}

}



/**
 * Function to edit the
 * reflection question
 * @param tablePkId
 * @param questnDesc
 * @param dropDownVal
 */
function editRefQtn(tablePkId, questnDesc, dropDownVal, subQtnDesc, profileId, counter) {		
	fetchReflectionQtnList('R' ,profileId , "BS");
	/******* ADDED FOR JCT PUBLIC VERSION ********/
	if(dropDownVal == "Default Profile") {
		var globalProfileChgDiv = document.getElementById('globalProfileChgDiv');
		globalProfileChgDiv.setAttribute("style", "display:block");	
	}	
	$("#refQuestionAddDiv").addClass("reflectionEdit");
	$('#globalProfileCheckBoxId').attr('checked', false);
	/****** JIRA ID - JCTP-8 **********/
	var subQtnAddDiv = document.getElementById('sunQtnAdditionDiv');
	subQtnAddDiv.setAttribute("style", "display:none");		
	
	/********** ENDED **********/	
	var chgOrderDiv = document.getElementById("chgOrderDiv");	
	chgOrderDiv.setAttribute("style", "display:none");
	$('#deleteBS_'+counter).addClass('disable_delete');

	document.getElementById('qtnDescInptId').value = questnDesc;
	document.getElementById('qtnDescInptId').setAttribute("readonly", "true");
	document.getElementById('inputUserProfileInpt').disabled = true;	
	var subQtnDivEdit = document.getElementById('sudQuestionEditDiv');
	subQtnDivEdit.setAttribute("style", "display:block");		
	document.getElementById('qtnSubDescEditId').value = subQtnDesc;
	document.getElementsByName("refQtnFld")[0].disabled = true;
	document.getElementsByName("refQtnFld")[1].disabled = true;
	document.getElementById('noOFSubQtnDiv').setAttribute("style", "display:none");
	document.getElementById("updateId").innerHTML = "Update Reflection Question";
	document.getElementById('addHeaderDiv').innerHTML = "Update Reflection Question / Action Plan";
	document.getElementById("addRefQtnBtnId").value = "Update";
	document.getElementById("addRefQtnBtnId").setAttribute("onclick", "updateRefQtn("+tablePkId+",'"+questnDesc+"')");
	var updateSection = document.getElementById("sudQuestionTotalEditDiv");
	updateSection.className = updateSection.className + " otherclass";
	//populateUserProfile
	var usrGrp = Spine.Model.sub();
	usrGrp.configure("/admin/contentconfig/populateUserProfileRefQtn", "none");
	usrGrp.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new usrGrp({  
		none: ""
	});
	modelPopulator.save(); //POST
	usrGrp.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	usrGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if(statusCode == 200) {
			var userProfileSelect = document.getElementById("inputUserProfileInpt");
			userProfileSelect.innerHTML = "";
			var defaultOpn = document.createElement("option");
			defaultOpn.text = "Select User Profile";
			defaultOpn.value = "";
			defaultOpn.className = "form-control-general";
			userProfileSelect.add(defaultOpn, null);
			for (var key in obj.userProfileMap) {				
				var option = document.createElement("option");
				option.text = obj.userProfileMap[key];
			    option.value = key;
			    option.className = "form-control-general";
			    if(option.text == dropDownVal) {
					option.setAttribute("selected", "selected");
					userProfileSelect.setAttribute("disabled", "true");
				}
			    try {
			    	userProfileSelect.add(option, null); //Standard 
			    }catch(error) {
			    	//regionSelect.add(option); // IE only
			    }
			}
		} else if(statusCode == 500) {
			// Error Message
		}
	});
}


/**
 * Function to update the
 * reflection question 
 * @param tablePkId
 * @param prevVal
 */
function updateRefQtn(tablePkId, prevVal) {
	//var reflectnQtn = document.getElementById("qtnDescInptId").value;
	var subQtn = document.getElementById("qtnSubDescEditId").value;
	var userProfile = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value;
	var userProfileText = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].text;
	var chechBoxVal = $("#globalProfileCheckBoxId").is(":checked");
	var cnfrmMsg = "";
	if (chechBoxVal == true) {
		cnfrmMsg = "Are you sure you want to make global profile change?";
	} else {
		cnfrmMsg = "Are you sure you want to update?";
	}
	alertify.set({ buttonReverse: true });
	alertify.confirm(cnfrmMsg, function (e) {
		if (e) {
			if(validateSubQtn(subQtn)) {
				var newUserProf = Spine.Model.sub();
				newUserProf.configure("/admin/contentconfig/updateRefQtn", "reflectnQtnVal", "subQtnVal" , "userProfileVal", "userProfileId", "tablePkId", "chechBoxVal");
				newUserProf.extend( Spine.Model.Ajax );
				//Populate the model with data to transfer
				var modelPopulator = new newUserProf({  
					//userGroupVal: userGroup,
					//reflectnQtnVal: prevVal,
					reflectnQtnVal: prevVal,
					subQtnVal: subQtn,
					userProfileVal: userProfileText,
					userProfileId: userProfile,
					tablePkId: tablePkId,
					chechBoxVal: chechBoxVal
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
						alertify.alert(obj.statusDesc);
						fetchReflectionQtnList('R', userProfile, "BS");
						
						document.getElementById('qtnDescInptId').value = "";
						document.getElementById('qtnDescInptId').removeAttribute("readonly", "true");
						document.getElementById('qtnDescInptId').removeAttribute("disabled", "true");
						document.getElementById('inputNoOFSubQtn').removeAttribute("readonly", "true");
						document.getElementById('inputNoOFSubQtn').removeAttribute("disabled", "true");
						
						var userProfileSelect = document.getElementById("inputUserProfileInpt");
						userProfileSelect.removeAttribute("disabled", "true");
						//document.getElementById('inputUserProfileInpt').setAttribute("readonly", "true");
						var subQtnDivEdit = document.getElementById('sudQuestionEditDiv');
						subQtnDivEdit.setAttribute("style", "display:none");		
						document.getElementById('qtnSubDescEditId').value = "";
						document.getElementsByName("refQtnFld")[0].disabled = false;
						document.getElementsByName("refQtnFld")[1].disabled = false;
						document.getElementById('noOFSubQtnDiv').setAttribute("style", "display:block");
						document.getElementById("updateId").innerHTML = "Reflection Question";
						document.getElementById("addHeaderDiv").innerHTML = "Add Reflection Question / Action Plan";
						/*document.getElementById("addActnPlanBtnId").value = "Add";
						document.getElementById("addActnPlanBtnId").setAttribute("onclick", "saveActionPlan()");*/
						document.getElementById("addRefQtnBtnId").value = "Add";
						document.getElementById("addRefQtnBtnId").setAttribute("onclick", "saveRefQtn()");
						var updateSection = document.getElementById("sudQuestionTotalEditDiv");
						updateSection.className = updateSection.className - " otherclass";
						
						var chgOrderDiv = document.getElementById("chgOrderDiv");	
						chgOrderDiv.setAttribute("style", "display:block");
						document.getElementById("chgOrderBtnId").disabled = false;
						
						/******* ADDED FOR JCT PUBLIC VERSION ********/
						var globalProfileChgDiv = document.getElementById('globalProfileChgDiv');
						globalProfileChgDiv.setAttribute("style", "display:none");	
						
					} else if(statusCode == 600) {
						alertify.alert("Reflection question already exists.");
					} else if(statusCode == 500) {
						//populateReflectionQtnTable(obj.existingRefQtnList, userProfile);
						fetchReflectionQtnList('R', userProfile, "BS");
						alertify.alert(obj.statusDesc);
					}
				});
			}
		}
		});	
}

/**
 * Function to validate the
 * reflection question text
 * @param reflectnQtn
 */
function validateSubQtn(subqtnDesc){
	var re=/^\s*$/;
	/****** validate sub question value ******/
	if (re.test(subqtnDesc)){
		alertify.alert("Please enter question value.");
		return false;
	} else if(subqtnDesc.search(/[?:"']{2,}/) > 0) {
		alertify.alert("Consecutive special character is not allowed");
		return false;
	} 		
	return true;
}


/**
 * Function to validate the
 * reflection question text
 * @param reflectnQtn
 */
function validateActionPlan(mainQtn, subQtnBuilder){
	var re=/^\s*$/;
	/****** Validate user profile is selected or not******/
	if(document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value == ""){
		alertify.alert("Please select User Profile.");
		return false;
	}
	/****** Validate main question value******/
	if (re.test(mainQtn)){
		alertify.alert("Please enter question value.");
		return false;
	} else if(mainQtn.search(/[?:"']{2,}/) > 0) {
		alertify.alert("Consecutive special character is not allowed");
		return false;
	}
	var builder = "";
	var noOFSubQtn = document.getElementById("inputNoOFSubQtn").value;	
	for(var i = 1; i <= noOFSubQtn; i++){
		if(null != document.getElementById("qtnSubDescInptId_"+i).value.trim() && document.getElementById("qtnSubDescInptId_"+i).value.trim() !=""){
			builder = builder + document.getElementById("qtnSubDescInptId_"+i).value +"~~";			
		} else {
			alertify.alert("Please enter sub question value.");
			return false;
		}					
	}
	/****** Validate question order******/
	/*if (re.test(questionOrder)){
		alertify.alert("Please enter question order.");
		return false;
	} else if(questionOrder == 0){
		alertify.alert("Question order cannot be 0.");
		return false;
	} else if(relatedPage == "BS"){
		var tbl = document.getElementById("refQtnTableId");
		if(tbl != null){
			var numRows = tbl.rows.length;
			var array = new Array();
			for (var i = 1; i < numRows; i++) {
			    var cells = tbl.rows[i].getElementsByTagName('td');
			    array.push(cells[2].innerHTML);
			}
			    if(array.indexOf(questionOrder) > -1){
			    	document.getElementById('qtnOrderInptId').value = "";	
					document.getElementById('inputNoOFSubQtn').value = "";	
					create("");
			    	alertify.alert("Question order already exists.");
					return false;
			    }
		}	
	} else if (relatedPage == "AS") {
		var tbl = document.getElementById("actionPlanTableId");
		if(tbl != null){
			var numRows = tbl.rows.length;
			var array = new Array();
			for (var i = 1; i < numRows; i++) {
			    var cells = tbl.rows[i].getElementsByTagName('td');
			    array.push(cells[2].innerHTML);
			}
			    if(array.indexOf(questionOrder) > -1){
			    	document.getElementById('qtnOrderInptId').value = "";	
					document.getElementById('inputNoOFSubQtn').value = "";	
					create("");
			    	alertify.alert("Question order already exists.");
					return false;
			}
		}	
	}*/
	
	/****** Validate sub question value******/
	var subQtnSplit = subQtnBuilder.split("~~");
	for(var index=0; index<subQtnSplit.length-1; index++){
		var eachSubQtn = subQtnSplit[index];
		if (re.test(eachSubQtn)){
			alertify.alert("Please enter sub question value.");
			return false;
		} else if(eachSubQtn.search(/[?:"']{2,}/) > 0) {
			alertify.alert("Consecutive special characters are not allowed");
			return false;
		}		
	}	
	
	
	return true;
}


/**
 * Function to delete the
 * reflection question
 * @param tablePkId
 * @param questnDesc
 */
function deleteRefQtn(tablePkId, questnDesc, profileId){	
	var relatedPage = "BS";
	alertify.set({ buttonReverse: true });
	alertify.confirm("Are you sure you want to delete this reflection question?", function (e) {
		if (e) {	    	
	    	var newUserProf = Spine.Model.sub();
	    	newUserProf.configure("/admin/contentconfig/deleteRefQtn", "tablePkId", "questnDesc", "profileId" , "relatedPage");
	    	newUserProf.extend( Spine.Model.Ajax );
	    	//Populate the model with data to transfer
	    	var modelPopulator = new newUserProf({  
	    		tablePkId: tablePkId,
	    		questnDesc: questnDesc,
	    		profileId: profileId,
	    		relatedPage: relatedPage
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
	    			//alertify.set({ delay: 2700 });
					alertify.alert("Reflection Question has been deleted successfully.");
					populateReflectionQtnTable(obj.existingRefQtnList, 0);
	    		} else if(statusCode == 500) {
	    			// Error Message
	    		}
	    	});
	    
		}
	});	
		
}

/**
 * Function to add the
 * action plan to the database
 * @param null
 */
function saveActionPlan() {	
	var relatedPage = "AS";
	var builder = "";
	var noOFSubQtn = document.getElementById("inputNoOFSubQtn").value;	
	for(var i = 1; i <= noOFSubQtn; i++){
		if(null != document.getElementById("qtnSubDescInptId_"+i).value.trim() && document.getElementById("qtnSubDescInptId_"+i).value.trim() !=""){
			builder = builder + document.getElementById("qtnSubDescInptId_"+i).value +"~~";			
		} /*else {
			alertify.alert("Please enter sub question.");
			return false;
		}	*/				
	}
	var refQtnText = document.getElementById("qtnDescInptId").value;
	var userProfile = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value;
	var userProfileText = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].text;
	//var questionOrder = document.getElementById("qtnOrderInptId").value;
	if(validateActionPlan(refQtnText, builder)) {
		var newUserProf = Spine.Model.sub();
		newUserProf.configure("/admin/contentconfig/saveActionPlan", "refQtnDesc", "subQtn", "userProfileVal", "userProfileId", "relatedPage", "createdBy");
		newUserProf.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new newUserProf({  
			refQtnDesc: refQtnText,
			subQtn: builder,
			userProfileVal: userProfileText,
			userProfileId: userProfile,			
			relatedPage: relatedPage,
			createdBy: sessionStorage.getItem("jctEmail")
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
				alertify.alert("Action Plan has been successfully created.");				
				fetchReflectionQtnList('A', userProfile, "AS");											
				document.getElementById('qtnDescInptId').value = "";	
				document.getElementById('qtnSubDescEditId').value = "";		
				document.getElementById('inputNoOFSubQtn').value = "";				
				//document.getElementById('qtnDescInptId').setAttribute("disabled", "true");				
				//document.getElementById('inputNoOFSubQtn').disabled = true;	
				create("", null);
			} else if(statusCode == 600) {
				alertify.alert("Action Plan already exists.");
			}  else if(statusCode == 602) {
				alertify.alert("One profile should have maximum of 6 action plan.");
			} else if(statusCode == 500) {
				// Error Message
			}
		});
	}
}


/**
 * Function to edit the
 * action plan questions
 * @param tablePkId
 * @param questnDesc
 * @param dropDownVal
 * @param subQtnDesc
 */
function editActionPlan(tablePkId, questnDesc, dropDownVal, subQtnDesc, profileId, counter){
	//fetchReflectionQtnList('R' ,profileId , "BS");
	fetchReflectionQtnList('A' ,profileId , "AS");
	/******* ADDED FOR JCT PUBLIC VERSION ********/
	if(dropDownVal == "Default Profile") {
		var globalProfileChgDiv = document.getElementById('globalProfileChgDiv');
		globalProfileChgDiv.setAttribute("style", "display:block");	
	}	
	$("#actionPlanAddDiv").addClass("reflectionEdit");
	$('#globalProfileCheckBoxId').attr('checked', false);
	/****** JIRA ID - JCTP-8 **********/
	var subQtnAddDiv = document.getElementById('sunQtnAdditionDiv');
	subQtnAddDiv.setAttribute("style", "display:none");	
	
	/********** ENDED **********/
	var chgOrderDiv = document.getElementById("chgOrderDiv");	
	chgOrderDiv.setAttribute("style", "display:none");
	$("#deleteAS_"+counter).addClass('disable_delete');

	/*$("#cancelActnPlanBtnId").addClass("search_btn_action");*/
	document.getElementById('qtnDescInptId').value = questnDesc;
	document.getElementById('qtnDescInptId').setAttribute("readonly", "true");
	document.getElementById('inputUserProfileInpt').disabled = true;	
	var subQtnDivEdit = document.getElementById('sudQuestionEditDiv');
	subQtnDivEdit.setAttribute("style", "display:block");		
	document.getElementById('qtnSubDescEditId').value = subQtnDesc;
	document.getElementsByName("refQtnFld")[0].disabled = true;
	document.getElementsByName("refQtnFld")[1].disabled = true;
	document.getElementById('noOFSubQtnDiv').setAttribute("style", "display:none");
	document.getElementById("updateId").innerHTML = "Update Action Plan";
	document.getElementById('addHeaderDiv').innerHTML = "Update Reflection Question / Action Plan";
	document.getElementById("addActnPlanBtnId").value = "Update";
	document.getElementById("addActnPlanBtnId").setAttribute("onclick", "updateActionPlan("+tablePkId+",'"+questnDesc+"')");
	var updateSection = document.getElementById("sudQuestionTotalEditDiv");
	updateSection.className = updateSection.className + " otherclass";
	//populateUserProfile
	var usrGrp = Spine.Model.sub();
	usrGrp.configure("/admin/contentconfig/populateUserProfileRefQtn", "none");
	usrGrp.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new usrGrp({  
		none: ""
	});
	modelPopulator.save(); //POST
	usrGrp.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	usrGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if(statusCode == 200) {
			var userProfileSelect = document.getElementById("inputUserProfileInpt");
			userProfileSelect.innerHTML = "";
			var defaultOpn = document.createElement("option");
			defaultOpn.text = "Select User Profile";
			defaultOpn.value = "";
			defaultOpn.className = "form-control-general";
			userProfileSelect.add(defaultOpn, null);
			for (var key in obj.userProfileMap) {				
				var option = document.createElement("option");
				option.text = obj.userProfileMap[key];
			    option.value = key;
			    option.className = "form-control-general";
			    if(option.text == dropDownVal) {
					option.setAttribute("selected", "selected");
					userProfileSelect.setAttribute("disabled", "true");
				}
			    try {
			    	userProfileSelect.add(option, null); //Standard 
			    }catch(error) {
			    	//regionSelect.add(option); // IE only
			    }
			}
		} else if(statusCode == 500) {
			// Error Message
		}
	});
}


/**
 * Function to update the
 * action plan question 
 * @param tablePkId
 * @param prevVal
 */
function updateActionPlan(tablePkId, prevVal) {
	//var reflectnQtn = document.getElementById("qtnDescInptId").value;
	var subQtn = document.getElementById("qtnSubDescEditId").value;
	var userProfile = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value;
	var userProfileText = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].text;
	var chechBoxVal = $("#globalProfileCheckBoxId").is(":checked");
	
	var cnfrmMsg = "";
	if (chechBoxVal == true) {
		cnfrmMsg = "Are you sure you want to make global profile change?";
	} else {
		cnfrmMsg = "Are you sure you want to update?";
	}
	alertify.set({ buttonReverse: true });
	alertify.confirm(cnfrmMsg, function (e) {
		if (e) {
			if(validateSubQtn(subQtn)) {
				var newUserProf = Spine.Model.sub();
				newUserProf.configure("/admin/contentconfig/updateActionPlan", "reflectnQtnVal", "subQtnVal" , "userProfileVal", "userProfileId", "tablePkId", "chechBoxVal");
				newUserProf.extend( Spine.Model.Ajax );
				//Populate the model with data to transfer
				var modelPopulator = new newUserProf({  
					//userGroupVal: userGroup,
					//reflectnQtnVal: prevVal,
					reflectnQtnVal: prevVal,
					subQtnVal: subQtn,
					userProfileVal: userProfileText,
					userProfileId: userProfile,
					tablePkId: tablePkId,
					chechBoxVal: chechBoxVal
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
						
						//alertify.alert("Action Plan has been edited successfully.");	
						alertify.alert(obj.statusDesc);
						
						fetchReflectionQtnList('A', userProfile, "AS");
						
						document.getElementById('qtnDescInptId').value = "";
						document.getElementById('qtnDescInptId').removeAttribute("readonly", "true");
						document.getElementById('qtnDescInptId').removeAttribute("disabled", "true");
						document.getElementById('inputNoOFSubQtn').removeAttribute("readonly", "true");
						document.getElementById('inputNoOFSubQtn').removeAttribute("disabled", "true");
						var userProfileSelect = document.getElementById("inputUserProfileInpt");
						userProfileSelect.removeAttribute("disabled", "true");
						
						//document.getElementById('inputUserProfileInpt').setAttribute("readonly", "true");
						var subQtnDivEdit = document.getElementById('sudQuestionEditDiv');
						subQtnDivEdit.setAttribute("style", "display:none");		
						document.getElementById('qtnSubDescEditId').value = "";
						document.getElementsByName("refQtnFld")[0].disabled = false;
						document.getElementsByName("refQtnFld")[1].disabled = false;
						document.getElementById('noOFSubQtnDiv').setAttribute("style", "display:block");
						document.getElementById("updateId").innerHTML = "Action Plan";
						document.getElementById("addHeaderDiv").innerHTML = "Add Reflection Question / Action Plan";
						document.getElementById("addActnPlanBtnId").value = "Add";
						document.getElementById("addActnPlanBtnId").setAttribute("onclick", "saveActionPlan()");
						var updateSection = document.getElementById("sudQuestionTotalEditDiv");
						updateSection.className = updateSection.className - " otherclass";
						
						var chgOrderDiv = document.getElementById("chgOrderDiv");	
						chgOrderDiv.setAttribute("style", "display:block");
						document.getElementById("chgOrderBtnId").disabled = false;
						
						/******* ADDED FOR JCT PUBLIC VERSION ********/
						var globalProfileChgDiv = document.getElementById('globalProfileChgDiv');
						globalProfileChgDiv.setAttribute("style", "display:none");	
						
					} else if(statusCode == 600) {
						alertify.alert("Action Plan already exists.");
					} else if(statusCode == 500) {
						//populateActionPlanTable(obj.existingRefQtnList, userProfile);
						fetchReflectionQtnList('A', userProfile, "AS");
						alertify.alert(obj.statusDesc);
					}
				});
			}			
		}
	});	

}

/**
 * Function to delete the
 * action plan question
 * @param tablePkId
 * @param questnDesc
 */
function deleteActionPlan(tablePkId, questnDesc, profileId){
	var relatedPage = "AS";
	alertify.set({ buttonReverse: true });
	alertify.confirm("Are you sure you want to delete the action plan?", function (e) {
		if (e) {	    	
	    	var newUserProf = Spine.Model.sub();
	    	newUserProf.configure("/admin/contentconfig/deleteActionPlan", "tablePkId", "questnDesc", "profileId" , "relatedPage");
	    	newUserProf.extend( Spine.Model.Ajax );
	    	//Populate the model with data to transfer
	    	var modelPopulator = new newUserProf({  
	    		tablePkId: tablePkId,
	    		questnDesc: questnDesc,
	    		profileId: profileId,
	    		relatedPage: relatedPage
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
	    			//alertify.set({ delay: 2700 });
					alertify.alert("Action Plan has been deleted successfully.");
					populateActionPlanTable(obj.existingRefQtnList, 0);
	    		} else if(statusCode == 500) {
	    			// Error Message
	    		}
	    	});
	    
		}
	});	
}

/**
 * Function add to populate
 * the no of sub question drop-down
 * @param noOfSubQtn
 */
function populateQtnDropDown(noOfSubQtn){	
	document.getElementById("inputNoOFSubQtn").innerHTML = "";
	var userProfileSelect = document.getElementById("inputNoOFSubQtn");		
	var defaultOptn = document.createElement("option");
	defaultOptn.text = "Select No of Sub Question";
	defaultOptn.value = "";
	defaultOptn.className = "form-control-general";
	userProfileSelect.add(defaultOptn, null);
	for(var i = 1; i<= noOfSubQtn;i++){		
		var option = document.createElement("option");
		option.text = i;
	    option.value = i;
	    option.className = "form-control-general";
	    try {
	    	userProfileSelect.add(option, null); //Standard 
	    }catch(error) {
	    	//regionSelect.add(option); // IE only
	    }	
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
 * Function call while click on reset button
 * to reset the page
 * @param null
 */
function resetValue() { 
	$("#refQuestionAddDiv").removeClass("reflectionEdit");
	$("#actionPlanAddDiv").removeClass("reflectionEdit");
	document.getElementsByName("refQtnFld")[0].disabled = false;
	document.getElementsByName("refQtnFld")[1].disabled = false;
	document.getElementById('inputUserProfileInpt').value = "";	
	document.getElementById('inputUserProfileInpt').disabled = false;
	document.getElementById("addActnPlanBtnId").value = "Add";
	document.getElementById("addRefQtnBtnId").value = "Add";
	document.getElementById("updateId").innerHTML = "Reflection Question / Action Plan";
	document.getElementById('addHeaderDiv').innerHTML = "Add Reflection Question / Action Plan";
	var updateSection2 = document.getElementById("sudQuestionTotalEditDiv");
	updateSection2.className = updateSection2.className - " otherclass";
	document.getElementById('qtnDescInptId').value = "";	
	document.getElementById('qtnSubDescEditId').value = "";		
	document.getElementById('inputNoOFSubQtn').value = "";	
	/*$("#cancelActnPlanBtnId").removeClass("search_btn_action");*/
	document.getElementById('sudQuestionEditDiv').setAttribute("style", "display:none");
	document.getElementById('noOFSubQtnDiv').setAttribute("style", "display:block");
	document.getElementById('globalProfileChgDiv').setAttribute("style", "display:none");
	
	var chgOrderDiv = document.getElementById("chgOrderDiv");	
	chgOrderDiv.setAttribute("style", "display:block");
	create("", null);
	disableAllfield();
	var selectedRadioButton = $('input[name="refQtnFld"]:checked').val();
	if(selectedRadioButton == "A"){
		fetchReflectionQtnList('A', 0, "AS");	
		document.getElementById('refQuestionAddDiv').setAttribute("style", "display:none");
		document.getElementById('actionPlanAddDiv').setAttribute("style", "display:block");
		document.getElementById("addActnPlanBtnId").setAttribute("onclick", "saveActionPlan()");
	} else {
		fetchReflectionQtnList('R', 0, "BS");	
		document.getElementById('refQuestionAddDiv').setAttribute("style", "display:block");
		document.getElementById('actionPlanAddDiv').setAttribute("style", "display:none");
		document.getElementById("addRefQtnBtnId").setAttribute("onclick", "saveRefQtn()");
	}
	document.getElementById('qtnDescInptId').removeAttribute("readonly");
	
}

/**
 * Function call to adjust the height
 * of the side menu bar
 * @param group
 */
function equalHeight(group) {
	   tallest = 0;
	   mmh = 15;
	   group.each(function() {
	      thisHeight = $(this).height();
	      //alert(thisHeight);
	      if(thisHeight > tallest) {
	         tallest = thisHeight;
	      }
	   });
	   group.height(tallest);
}

/**
 * Function call against the profile changed
 * @param obj
 */
function changeUserProfile(obj) {
	var relatedPage = "";
	var selectedRadioButton = $('input[name="refQtnFld"]:checked').val();
	if(selectedRadioButton == "A"){
		relatedPage = "AS";
	} else {
		relatedPage = "BS";
	}
	if(null == obj.value || obj.value == "") {
		document.getElementById('qtnDescInptId').setAttribute("disabled", "true");
		//document.getElementById('qtnOrderInptId').setAttribute("disabled", "true");
		document.getElementById('inputNoOFSubQtn').disabled = true;	
		document.getElementById("chgOrderBtnId").disabled=true;
		fetchReflectionQtnList(selectedRadioButton, 0, relatedPage);
	} else {		
		document.getElementById('qtnDescInptId').removeAttribute("disabled", "false");
		//document.getElementById('qtnOrderInptId').removeAttribute("disabled", "false");
		document.getElementById('inputNoOFSubQtn').disabled = false;	
		document.getElementById("chgOrderBtnId").disabled=false;
		fetchReflectionQtnList(selectedRadioButton, obj.value, relatedPage);
	}
	create("", null);
}

/**
 * Function add to allow only numbers
 * as a input except 0
 * @param event
 */
function numberOnly(event) {	
	 var c = event.which || event.keyCode;
	 if ( (c >= 48  && c <= 57) || c == 43) 
	        return true;
	    return false;	  
}


/**
 * Function add to disable all the fields while reset
 * @param null
 */
function disableAllfield() {
	document.getElementById('qtnDescInptId').setAttribute("disabled", "true");
	//document.getElementById('qtnOrderInptId').setAttribute("disabled", "true");
	document.getElementById('inputNoOFSubQtn').disabled = true;	
	document.getElementById("chgOrderBtnId").disabled = true;
}

/**
 * Function add to allow only a, b, c, d
 * as input in sub question order field
 * @param event
 */
function exceptAlphabetOnly(event){
	var c = event.which || event.keyCode;
	console.log(c);
	if ((c >96 &&  c < 101) || c == 8)
        return true;
    else
        return false; 	
}


/**
 * Function to view all the main question 
 * from database to a popup window
 * @param objval, profileId, relatedPage
 */
function fetchMainQuestionList(objVal, profileId, relatedPage) {
	var userProf = Spine.Model.sub();
	userProf.configure("/admin/contentconfig/populateMainQtn", "profileId", "relatedPage");
	userProf.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userProf({  
		profileId: profileId,
		relatedPage: relatedPage
	});
	modelPopulator.save(); //POST
	userProf.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	userProf.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if(statusCode == 200) {				
			//if( null == objVal|| objVal == "" || objVal == 'R') {
				//populateReflectionQtnTable(obj.existingRefQtnList, profileId);
				populateMainQtnTable(obj.mainQtnList);
			//} else {
				//populateActionPlanTable(obj.mainQtnList, profileId);
			//}
			
		} else if(statusCode == 201) {
			document.getElementById("mainQtnListId").innerHTML = "<div align='center'>"+obj.statusDesc+"</div>";
		} else if(statusCode == 500) {
			//Show error message
		}
	});
}

/**
 * Function to view all the main question
 * in a popup window to change the question order
 * @param null
 */
function changeOrder() {
	manipulativeQtn = "";
	var relatedPage = "";
	var selectedRadioButton = $('input[name="refQtnFld"]:checked').val();
	if(selectedRadioButton == "A"){
		relatedPage = "AS";
	} else {
		relatedPage = "BS";
	}
	var userProfile = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value;
	var userProfileText = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].text;
	document.getElementById("profileDescId").innerHTML = "Profile: "+userProfileText;
	var subQtnList = document.getElementById("subQtnListId");	
	subQtnList.setAttribute("style", "display:none");
	var mainQtnList = document.getElementById("mainQtnListId");	
	mainQtnList.setAttribute("style", "display:block");
	
	var subQtnSaveBtn = document.getElementById("subQtnSaveBtnId");	
	subQtnSaveBtn.setAttribute("style", "display:none");
	
	var mainQtnSaveBtn = document.getElementById("mainQtnSaveBtnId");	
	mainQtnSaveBtn.setAttribute("style", "display:block");
	
	var chgSubQtnLink = document.getElementById("chgSubQtnLinkId");	
	chgSubQtnLink.setAttribute("style", "display:block");
	
	/*var backBtn = document.getElementById("backBtnId");	
	backBtn.setAttribute("style", "display:none");*/
	
	var headerDescMain = document.getElementById("headerDescMainId");	
	headerDescMain.setAttribute("style", "display:block");
	
	var headerDescSub = document.getElementById("headerDescSubId");	
	headerDescSub.setAttribute("style", "display:none");	
	
	var mainQtnDesc = document.getElementById("mainQtnDescId");	
	mainQtnDesc.setAttribute("style", "display:none");
	
	fetchMainQuestionList(selectedRadioButton, userProfile, relatedPage);
}


/**
 * Function to populate the main question list
 * in a table view
 * @param refQtnListData
 */
function populateMainQtnTable(refQtnListData) {
	var tableDiv = document.getElementById("mainQtnListId");
	if (null != refQtnListData) {
		var headingString = "<table width='90%' id='table-1' border='1' align='center' class='tablesorter'><thead align='center' class='tab_header'><tr><th><b>Main Question</b></th></tr></thead><tbody>";
		var splitRec = refQtnListData.split("$$$");
		var counter = 1;
		for(var outerIndex = 0; outerIndex < splitRec.length-1; outerIndex++){			
			var objSplit = splitRec[outerIndex].split("~");		
				headingString = headingString + "<tr align='center' id="+counter+" class='custom_table_row'><td align='left' onclick='populateMainQtnsForSubQtn(\""+objSplit[0]+"\",\""+counter+"\")'>"+objSplit[0]+"</td>";	
				headingString = headingString + "</tr>";
				counter = counter + 1;							
		}		
		headingString = headingString + "</tbody></table>";
		tableDiv.innerHTML = headingString;
	} 
	 $("#table-1").tableDnD();
	 $("table").tablesorter(); 
}

/**
 * Function to view all the sub question
 * in a popup window to change the sub question order
 * @param null
 */
function changeSubOrder() {	
	if(null == manipulativeQtn || manipulativeQtn == ""){
		alertify.alert("Please select main question first.");
		return false;
	} else {
		var relatedPage = "";
		var selectedRadioButton = $('input[name="refQtnFld"]:checked').val();
		if(selectedRadioButton == "A"){
			relatedPage = "AS";
		} else {
			relatedPage = "BS";
		}
		var userProfile = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value;
		var userProfileText = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].text;
		document.getElementById("profileDescId").innerHTML = "Profile: "+userProfileText;
		
		
		
		
		var mainQtnList = document.getElementById("mainQtnListId");	
		mainQtnList.setAttribute("style", "display:none");
		
		var subQtnList = document.getElementById("subQtnListId");	
		subQtnList.setAttribute("style", "display:block");
		
		var subQtnSaveBtn = document.getElementById("subQtnSaveBtnId");	
		subQtnSaveBtn.setAttribute("style", "display:block");
		
		var mainQtnSaveBtn = document.getElementById("mainQtnSaveBtnId");	
		mainQtnSaveBtn.setAttribute("style", "display:none");
		
		var chgSubQtnLink = document.getElementById("chgSubQtnLinkId");	
		chgSubQtnLink.setAttribute("style", "display:none");		
		
		/*var backBtn = document.getElementById("backBtnId");	
		backBtn.setAttribute("style", "display:block");*/
		
		var headerDescMain = document.getElementById("headerDescMainId");	
		headerDescMain.setAttribute("style", "display:none");
		
		var headerDescSub = document.getElementById("headerDescSubId");	
		headerDescSub.setAttribute("style", "display:block");
		
		var mainQtnDesc = document.getElementById("mainQtnDescId");	
		mainQtnDesc.setAttribute("style", "display:block");
		
		fetchSubQuestionList(selectedRadioButton, userProfile, relatedPage, manipulativeQtn);
		
	}	
}

/**
 * Function to view all the sub question 
 * from database to a popup window
 * @param objval, profileId, relatedPage, mainQuestion
 */
function fetchSubQuestionList(radioBtnVal, profileId, relatedPage, mainQuestion) {
	var userProf = Spine.Model.sub();
	userProf.configure("/admin/contentconfig/populateSubQtn", "profileId", "relatedPage", "mainQuestion");
	userProf.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userProf({  
		profileId: profileId,
		relatedPage: relatedPage,
		mainQuestion: mainQuestion
	});
	modelPopulator.save(); //POST
	userProf.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	userProf.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if(statusCode == 200) {				
			var splitRec = obj.subQtnList.split("$$$");
			if(splitRec.length == 2){
				changeOrder();
				alertify.alert("Cannot change the order of sub question.");				
			} else {
				populateSubQtnTable(obj.subQtnList, mainQuestion);
			}				
		} else if(statusCode == 201) {
			document.getElementById("subQtnListId").innerHTML = "<div align='center'>"+obj.statusDesc+"</div>";
		} else if(statusCode == 500) {
			//Show error message
		}
	});
}


/**
 * Function to populate the sub question list
 * in a table view
 * @param refQtnListData, mainQuestion
 */
function populateSubQtnTable(refQtnListData, mainQuestion) {	
	document.getElementById("mainQtnDescId").innerHTML = "Main Question: "+mainQuestion;
	var tableDiv = document.getElementById("subQtnListId");
	if (null != refQtnListData) {
		var headingString = "<table width='90%' id='table-2' border='1' align='center' class='tablesorter'><thead align='center' class='tab_header'><tr style='cursor: default;'><th><b>Sub Question</b></th></tr><tbody>";
		var splitRec = refQtnListData.split("$$$");
		var counter = 1;
		for(var outerIndex = 0; outerIndex < splitRec.length-1; outerIndex++){			
			var objSplit = splitRec[outerIndex].split("~");		
				headingString = headingString + "<tr align='center' id="+counter+" class='custom_table_row' onmousedown='highlight()'><td align='left'>"+objSplit[0]+"</td>";								
				headingString = headingString + "</tr>";
				counter = counter + 1;						
		}		
		headingString = headingString + "</tbody></table>";
		tableDiv.innerHTML = headingString;
		
	} 
	 $("#table-2").tableDnD();
	 $("table").tablesorter(); 
}

function highlight() {
	$("tr").mousedown(function(){
		$(this).addClass("selected").siblings().removeClass("selected");
	});
}
/**
 * Function to store the main question value
 * in a global variable
 * @param v
 */
function populateMainQtnsForSubQtn(mainQtn, trId) {
	manipulativeQtn = "";
	manipulativeQtn = mainQtn;
	$("tr").click(function(){
		$(this).addClass("selected").siblings().removeClass("selected");
	});
}

/**
 * Function to save the main question
 * by main question order
 * @param null
 */
function saveMainQtnOrder() {
	var relatedPage = "";
	var selectedRadioButton = $('input[name="refQtnFld"]:checked').val();
	if(selectedRadioButton == "A"){
		relatedPage = "AS";
	} else {
		relatedPage = "BS";
	}
	var userProfile = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value;
	var userProfileText = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].text;
	var builder = "";
	$('#table-1 tbody td').each(function() {
	    var cellText = $(this).html();  
	    builder = builder + cellText +"~~";	
	});
	var newUserProf = Spine.Model.sub();
	newUserProf.configure("/admin/contentconfig/saveMainQtnOrder", "builder", "userProfile", "relatedPage");
	newUserProf.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new newUserProf({  
		builder: builder,
		userProfile: userProfile,
		relatedPage: relatedPage
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
			alertify.alert("Main Question Order has been changed successfully.");
			manipulativeQtn = "";
			populateMainQtnTable(obj.mainQtnList);
			fetchReflectionQtnList(selectedRadioButton, userProfile, relatedPage);
			//fetchMainQuestionList(selectedRadioButton, userProfile, relatedPage);		
		}  else if(statusCode == 500) {
			// Error Message
		}
	});
}

/**
 * Function to save the sub question
 * by sub question order
 * @param null
 */
function saveSubQtnOrder() {
	var relatedPage = "";
	var selectedRadioButton = $('input[name="refQtnFld"]:checked').val();
	if(selectedRadioButton == "A"){
		relatedPage = "AS";
	} else {
		relatedPage = "BS";
	}
	var userProfile = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value;
	var userProfileText = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].text;
	var builder = "";
	$('#table-2 tbody td').each(function() {
	    var cellText = $(this).html();  
	    builder = builder + cellText +"~~";	
	});
	var newUserProf = Spine.Model.sub();
	newUserProf.configure("/admin/contentconfig/saveSubQtnOrder", "builder", "userProfile", "relatedPage", "manipulativeQtn");
	newUserProf.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new newUserProf({  
		builder: builder,
		userProfile: userProfile,
		relatedPage: relatedPage,
		manipulativeQtn: manipulativeQtn
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
			//alertify.set({ delay: 2700 });
			alertify.alert("Sub Question has been changed successfully.");
			populateSubQtnTable(obj.subQtnList, manipulativeQtn);
			fetchReflectionQtnList(selectedRadioButton, userProfile, relatedPage);
			//fetchSubQuestionList(selectedRadioButton, userProfile, relatedPage);
			}
			else if(statusCode == 500) {
			// Error Message
		}
	});
}

/**
 * Function to navigate to previous window
 * @param null
 */
function backToPrevious() {
	var mainQtnList = document.getElementById("mainQtnListId");	
	mainQtnList.setAttribute("style", "display:block");
	
	var subQtnList = document.getElementById("subQtnListId");	
	subQtnList.setAttribute("style", "display:none");
	
	var subQtnSaveBtn = document.getElementById("subQtnSaveBtnId");	
	subQtnSaveBtn.setAttribute("style", "display:none");
	
	var mainQtnSaveBtn = document.getElementById("mainQtnSaveBtnId");	
	mainQtnSaveBtn.setAttribute("style", "display:block");
	
	var chgSubQtnLink = document.getElementById("chgSubQtnLinkId");	
	chgSubQtnLink.setAttribute("style", "display:block");		
	
	var headerDescMain = document.getElementById("headerDescMainId");	
	headerDescMain.setAttribute("style", "display:block");
	
	var headerDescSub = document.getElementById("headerDescSubId");	
	headerDescSub.setAttribute("style", "display:none");
	
	var mainQtnDesc = document.getElementById("mainQtnDescId");	
	mainQtnDesc.setAttribute("style", "display:none");
	
	/*var backBtn = document.getElementById("backBtnId");	
	backBtn.setAttribute("style", "display:none");*/
}