var storedFileSize = 0;
$(document).ready(function() {
	document.getElementById('fileType').setAttribute("disabled", "true");
	//document.getElementById('instructionHeaderText').setAttribute("disabled", "true");
	populateUI();
});


/**
 * Function to check the extension of the uploaded file
 * @param null
 */
function validateVideoFile() {
	var parts = document.getElementById('filename').value.split('.');
	var userProfile = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value;
	var relatedPage = document.getElementById("fileType").options[document.getElementById("fileType").selectedIndex].value;
	//var headerName = document.getElementById("instructionHeaderText").value;
	
	if( userProfile == 0 ){
		alertify.alert('Please select the user profile.');
		return false;
	}
	if( relatedPage == 0 ){
		alertify.alert('Please select Related Page.');
		return false;
	}
	var instructionDesc = CKEDITOR.instances.inputText.getData().trim();
	if (null == instructionDesc || instructionDesc == "" || instructionDesc == "&nbsp;" || instructionDesc.length == 0 ){
		alertify.alert("Please enter instuction data.");
		return false;
	}
	
	if ((null == instructionDesc || instructionDesc == "" || instructionDesc == "&nbsp;" ) && (null == document.getElementById('filename').value || document.getElementById('filename').value == "")){
		alertify.alert("Either enter text or upload video to continue.");
		return false;
	}
	
	if((null != document.getElementById('filename').value) && (document.getElementById('filename').value != "")) {
		if (parts[parts.length-1] != 'mp4') {
			alertify.alert('Please upload mp4 format only.');
		    return false;
		}else {
		if (storedFileSize > 15) {
				alertify.alert('Video file is too big.');
				return false;
		}
			document.getElementById("hiddenFileName").value = document.getElementById('filename').value;
			return true;
		}
	}
}
	

/**
 * Function to view all the reflection question 
 * from database while landing to the page
 * @param objval
 */
function populateUI() {	
	var userProf = Spine.Model.sub();
	userProf.configure("/admin/populateExistingVideos", "none");
	userProf.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userProf({  
		none: ""
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
			populateProfileDropDown(obj.userProfileMap);		
		} 
	});
	$( 'html, body, document' ).stop().animate( { scrollTop: 0 }, 'slow' );
}

/**
 * Function add to populate
 * the user profile drop-down value
 * @param userProfileMap
 */
function populateProfileDropDown(userProfileMap){
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

function changeVideo(obj) {
	var pageType = document.getElementById("fileType").options[document.getElementById("fileType").selectedIndex].value;
	if (pageType != "0") {
		document.getElementById("existing_data_div").style.display = "block";
	} else {
		document.getElementById("existing_data_div").style.display = "none";
	}
	var plot = document.getElementById("videoSection");
	plot.innerHTML = "<div align='center'><img src='../img/Processing.gif' /></div>";
	var profileId = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value;
	if (pageType != "0" && profileId != "") {
		var userProf = Spine.Model.sub();
		userProf.configure("/admin/populateVideosBySelection", "profileId", "page");
		userProf.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new userProf({  
			profileId: profileId,
			page: pageType
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
				plotTextAndVideo(obj);						
			} 
		});
	} else {
		document.getElementById("videoSection").innerHTML = "";
	}
	$( 'html, body, document' ).stop().animate( { scrollTop: 0 }, 'slow' );
}

function plotVideo (videoLink, desc) {
	// Changes by Rajan to match different URL context
	if(videoLink) {
		//adding /admin url  - HARDCODING
		videoLink = "https://s3-us-west-2.amazonaws.com/jobcrafting/" + videoLink;
	}
	// Changes by Rajan end
	document.getElementById("existing_data_div").style.display = "block";
	var plot = document.getElementById("videoSection");
	plot.innerHTML = "";
	if(videoLink === null) {
		document.getElementById('existing_list_Id').innerHTML = "";
		plot.innerHTML = "<div align='center'><img src='/admin/img/no-record.png'><br />"+desc+"</div>";
	} else {
		document.getElementById('existing_list_Id').innerHTML = "Existing Video";
		plot.innerHTML = "<div align='center'><video width='25%' poster='../img/frame.jpg' controls preload='auto'><source src='"+videoLink+"' type='video/mp4'></video>";
	}	
}


function plotTextAndVideo (obj) {
	document.getElementById("existing_data_div").style.display = "block";
	var plot = document.getElementById("videoSection");	
	plot.innerHTML = "";
	CKEDITOR.instances.inputText.setData('');
	CKEDITOR.instances.inputTextAfterVideo.setData('');
	//document.getElementById('instructionHeaderText').value = obj.instructionHeader;
	setTimeout(function(){ 
		if (obj.instructionBeforeVideo){
			CKEDITOR.instances['inputText'].insertHtml(obj.instructionBeforeVideo);
		} else {
			CKEDITOR.instances.inputText.setData('');
		}
	}, 100);
	if (obj.videoLink){
		// Changes by Rajan to match different URL context
		var videoLink = obj.videoLink;
		//adding /admin url  - HARDCODING
		videoLink = "https://s3-us-west-2.amazonaws.com/jobcrafting/" + videoLink;
		document.getElementById('existing_list_Id').innerHTML = "Existing Video";
		document.getElementById("hiddenFileName").value = obj.videoLink;	
		//Earlier code // plot.innerHTML = "<div align='center'><video width='25%' poster='../img/frame.jpg' controls preload='auto'><source src='"+obj.videoLink+"' type='video/mp4'></video>";
		plot.innerHTML = "<div align='center'><video width='25%' poster='../img/frame.jpg' controls preload='auto'><source src='"+videoLink+"' type='video/mp4'></video>";
		// Changes by Rajan end
	} else {
		plot.innerHTML = "";		
		document.getElementById("existing_data_div").style.display = "none";
	}
	setTimeout(function(){ 
		if (obj.instructionAfterVideo){
			CKEDITOR.instances['inputTextAfterVideo'].insertHtml(obj.instructionAfterVideo);
		} else {
			CKEDITOR.instances.inputTextAfterVideo.setData('');
		}
	}, 100);
	/*** Adjusting the footer div ***/
	var maxHeight = "";
	if(navigator.userAgent.match(/ipad/i) ){
		 maxHeight = $("#existing_data_div").outerHeight()+130;	
		} else {
		 maxHeight = $("#existing_data_div").outerHeight()+40;	
		}
	//var maxHeight = $("#existing_data_div").outerHeight()+50;	
	$(".footer").css({'margin-top' : maxHeight+'px'});
	/*******************************/
}

function getFileSize(fileid) {
	storedFileSize = 0;
 try {
 var fileSize = 0;
 fileSize = $("#" + fileid)[0].files[0].size; //size in kb
 fileSize = fileSize / 1048576; //size in mb 
 storedFileSize = fileSize;
 }
 catch (e) {

 }
}

function manageDiv(selection, hidingDiv, headerText, action) {
	if (document.getElementById(selection).style.display == "none") {
		document.getElementById(selection).style.display = "block";
		document.getElementById(hidingDiv).style.display = "none";
	} if (hidingDiv == "videoDivId") {
		document.getElementById("videoSection").style.display = "none";
		document.getElementById("existing_data_div").style.display = "none";
	} else {
		document.getElementById("videoSection").style.display = "block";
	}
	if(action == "textVideo") {
		document.getElementById('videoTextDivId').setAttribute("style", "display:block");	
		document.getElementById('fileType').setAttribute("disabled", "true");
	} else if(action == "onlyVideo") {
		document.getElementById('videoTextDivId').setAttribute("style", "display:none");
		document.getElementById('fileType').setAttribute("disabled", "true");
		document.getElementById("existing_data_div").style.display = "none";
	} else {
		document.getElementById('videoTextDivId').setAttribute("style", "display:none");
		document.getElementById('fileType1').setAttribute("disabled", "true");
		document.getElementById("existing_data_div").style.display = "none";
	}
	
	document.getElementById('headerId').innerHTML = headerText;
	document.getElementById('inputUserProfileInpt1').value = "";	
	document.getElementById('inputUserProfileInpt').value = "";	
	document.getElementById('fileType').value = 0;	
	document.getElementById('fileType1').value = 0;	
	CKEDITOR.instances.editor1.setData('');
	CKEDITOR.instances.inputText.setData('');
	$( 'html, body, document' ).stop().animate( { scrollTop: 0 }, 'fast' );
}

function changeTextIns(obj) {
	var userProfile = document.getElementById("inputUserProfileInpt1").options[document.getElementById("inputUserProfileInpt1").selectedIndex].value;
	var relatedPage = document.getElementById("fileType1").options[document.getElementById("fileType1").selectedIndex].value;
	if ((userProfile != "") && (relatedPage != "0")) {
		var newUserProf = Spine.Model.sub();
		newUserProf.configure("/admin/populateTextInsBySelection", "profileId", "page");
		newUserProf.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new newUserProf({  
			profileId: userProfile,
			page: relatedPage
		});
		modelPopulator.save(); //POST
		newUserProf.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
		});
		newUserProf.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = JSON.parse(jsonStr);
			if (obj.instruction.trim().length > 0) {
				CKEDITOR.instances['editor1'].insertHtml(obj.instruction);
				document.getElementById("addInstructionBtnId").value = "Update";
			} else {
				CKEDITOR.instances.editor1.setData('');
				if (document.getElementById("addInstructionBtnId").value == "Update") {
					document.getElementById("addInstructionBtnId").value = "Add";
				}
			}
		});
	} else {
		CKEDITOR.instances.editor1.setData('');
		if (document.getElementById("addInstructionBtnId").value == "Update") {
			document.getElementById("addInstructionBtnId").value = "Add";
		}
	}
}

function saveInstruction () {
	var insAreaText = CKEDITOR.instances.editor1.getData();
	var userProfile = document.getElementById("inputUserProfileInpt1").options[document.getElementById("inputUserProfileInpt1").selectedIndex].value;
	var relatedPage = document.getElementById("fileType1").options[document.getElementById("fileType1").selectedIndex].value;
	if (validateInstruction (insAreaText, userProfile, relatedPage)) {
		var newUserProf = Spine.Model.sub();
		newUserProf.configure("/admin/savePopupInstruction", "insAreaTextVal", "userProfileId", "relatedPage");
		newUserProf.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new newUserProf({  
			insAreaTextVal: insAreaText,
			userProfileId: userProfile,
			relatedPage: relatedPage
		});
		modelPopulator.save(); //POST
		newUserProf.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
		});
		newUserProf.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			//var obj = jQuery.parseJSON(jsonStr);
			var obj = JSON.parse(jsonStr);
			alertify.alert(obj.statusDesc);
			/*if (document.getElementById("addInstructionBtnId").value == "Update") {
				document.getElementById("addInstructionBtnId").value = "Add";
			}*/
			var userProfileSelect = document.getElementById("inputUserProfileInpt1");
			userProfileSelect.setAttribute("disabled", "true");
			var relatedPage = document.getElementById("fileType1");
			relatedPage.setAttribute("disabled", "true"); 
			var addInstructionBtnId = document.getElementById("addInstructionBtnId");
			addInstructionBtnId.setAttribute("disabled", "true"); 
			
		});
	} 
}

function validateInstruction (instructionDesc, userProfile, relatedPage) {
	if(userProfile == ""){
		alertify.alert("Please select User Profile.");
		return false;
	}
	if(relatedPage == "0"){
		alertify.alert("Please select Related Page.");
		return false;
	}
	if (null == instructionDesc || instructionDesc == "" || instructionDesc == "&nbsp;"){
		alertify.alert("Please enter instuction data.");
		return false;
	}
	return true;
}
function cancelInstruction() {
	/*var userProfileSelect = document.getElementById("inputUserProfileInpt1");
	userProfileSelect.removeAttribute("disabled", "true");
	var relatedPage = document.getElementById("fileType1");
	relatedPage.removeAttribute("disabled", "true"); */
	
	var addInstructionBtnId = document.getElementById("addInstructionBtnId");
	addInstructionBtnId.removeAttribute("disabled", "true"); 
	document.getElementById("inputUserProfileInpt1").selectedIndex = "0";
	document.getElementById("fileType1").selectedIndex = "0";
	document.getElementById('fileType1').setAttribute("disabled", "true");
	CKEDITOR.instances.editor1.setData('');
	if (document.getElementById("addInstructionBtnId").value == "Update") {
		document.getElementById("addInstructionBtnId").value = "Add";
	}
}
function resetuploadVDOfield(){
	document.getElementById("inputUserProfileInpt").selectedIndex = "0";
	document.getElementById("fileType").selectedIndex = "0";
	document.getElementById('fileType').setAttribute("disabled", "true");
	document.getElementById("existing_data_div").style.display = "none";	
	$(".form-control").val("");
	CKEDITOR.instances.inputText.setData('');
	CKEDITOR.instances.inputTextAfterVideo.setData('');
	//document.getElementById("instructionHeaderText").value = "";
	$( 'html, body, document' ).stop().animate( { scrollTop: 0 }, 'slow' );
}

/**
 * Function to disable the related page dropdown
 * if user profile is not selected
 * and enable the related page dropdown
 * if user profile is selected
 * @param obj
 */
function disablePageFieldVideo(obj) {
	if(null == obj.value || obj.value == "") {
		document.getElementById('fileType').setAttribute("disabled", "true");
		//document.getElementById('instructionHeaderText').setAttribute("disabled", "true");
	} else {
		document.getElementById('fileType').removeAttribute("disabled", "true");
		//document.getElementById('instructionHeaderText').removeAttribute("disabled", "true");
	}
}

/**
 * Function to disable the related page dropdown
 * if user profile is not selected
 * and enable the related page dropdown
 * if user profile is selected
 * @param obj
 */
function disablePageFieldText(obj) {
	if(null == obj.value || obj.value == "") {
		document.getElementById('fileType1').setAttribute("disabled", "true");
	} else {
		document.getElementById('fileType1').removeAttribute("disabled", "true");
	}
}

function removeVideo() {
	document.getElementById("existing_data_div").style.display = "none";
	document.getElementById("hiddenFileName").value = "";	
}