var mainHeight = $(".main_contain").height();
/**
 * Function calls when the page is loaded 
 * to populate the drop-down for user profile
 */
/*window.onload = function () {
if($(document).scrollTop() !=0|| $('html').scrollTop() !=0 || $('body').scrollTop() !=0){
	//alert("window here");
	$( 'html, body, document' ).stop().animate( { scrollTop: 0 }, 'fast' );
	}
};*/

$(document).ready(function() {	
	document.getElementById('inputPageInpt').setAttribute("disabled", "true");
	fetchDropDownValues();
	/*$(function() {
        $("#editor1").htmlarea({
           toolbar: [
						["bold", "italic", "underline"],
						["orderedlist", "unorderedlist"],
						["indent", "outdent"],
					 ]
        	toolbar: [
						[ 'Source', '-', 'Bold', 'Italic', 'Underline', '-', 'BulletedList', 'NumberedList' ],
						[ 'JustifyLeft', 'JustifyCenter', 'JustifyRight', 'JustifyBlock' , 'IndentLeft', 'Outdent','Indent'],
						[ 'Font', 'FontSize' ],
						[ 'TextColor', '-', 'About' ]
					],
			});
	});*/
	
});

/**
 * Function to fetch the 
 * user profile drop down
 * @param null
 */
function fetchDropDownValues() {
	var userProf = Spine.Model.sub();
	userProf.configure("/admin/contentconfig/populateUserProfileGeneral", "none");
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
		//var obj = jQuery.parseJSON(jsonStr);
		var obj = JSON.parse(jsonStr);
		var statusCode = obj.statusCode;
		if(statusCode == 200) {
			populateDropDown(obj.userProfileMap);										
		} else if(statusCode == 201) {
			//Show error message
		} else if(statusCode == 500) {
			//Show error message
		}
	});
	$( 'html, body, document' ).stop().animate( { scrollTop: 0 }, 'slow' );
}


/**
 * Function to disable the related page dropdown
 * if user profile is not selected
 * and enable the related page dropdown
 * if user profile is selected
 * @param obj
 */
function disablePageField(obj) {
	if(null == obj.value || obj.value == "") {
		document.getElementById('inputPageInpt').setAttribute("disabled", "true");
	} else {
		document.getElementById('inputPageInpt').removeAttribute("disabled", "true");
	}
}


/**
 * Function to fetch the instruction data from the db
 * for the selected profile and related page
 * @param objVal
 */
function fetchInstructionData(objVal) {
	var userProfile = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value;
	var relatedPageVal = objVal.value;
	var userProf = Spine.Model.sub();
	userProf.configure("/admin/contentconfig/populateExistingInstruction", "userProfileId", "relatedPageText");
	userProf.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userProf({  
		userProfileId: userProfile,
		relatedPageText: relatedPageVal
	});
	modelPopulator.save(); //POST
	userProf.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	userProf.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		//var obj = jQuery.parseJSON(jsonStr);
		var obj = JSON.parse(jsonStr);
		var statusCode = obj.statusCode;
		if(statusCode == 200) {			
			populateInstructionDataTable(obj.existingInstruction, obj.existingInstructionId, obj.existingInstructionType , obj.existingInstructionVideo);			
		} else if(statusCode == 201) {
			populateNone();
		} else if(statusCode == 500) {
			//Show error message
		}
	});
}


/**
 * Function to show the instruction data
 * for the selected profile and related page
 * @param instructionData
 * @param instructionId
 */
function populateInstructionDataTable(instructionData, instructionId, insType, videoLink) {
	// Changes by Rajan to match different URL context
	if(videoLink) {
		//adding /admin url  - HARDCODING
		videoLink = "https://s3-us-west-2.amazonaws.com/jobcrafting/" + videoLink;
	}
	// Changes by Rajan end
	CKEDITOR.instances['editor1'].insertHtml(instructionData);
	/*$('#editor1').htmlarea('html', instructionData);*/
	document.getElementById("instructionUpdateDiv").setAttribute("style", "display:block");
	document.getElementById("instructionAddDiv").setAttribute("style", "display:none");
	var userProfileSelect = document.getElementById("inputUserProfileInpt");
	userProfileSelect.setAttribute("disabled", "true");
	var relatedPage = document.getElementById("inputPageInpt");
	relatedPage.setAttribute("disabled", "true");
	//document.getElementById("updateInstructionBtnId").setAttribute("onclick", "updateInstruction("+instructionId+")");
	
	if(insType == 'TEXTANDVIDEO') {
		$('#incVideo').prop('checked', true);
		document.getElementById("videoDivId").style.display = "block";
	} else {
		$('#incVideo').prop('checked', false);
		document.getElementById("videoDivId").style.display = "none";
	}
			
	var plot = document.getElementById("videoSection");	
	plot.innerHTML = "";
	if (videoLink){
		//document.getElementById('existing_list_Id').innerHTML = "Existing Video";
		//document.getElementById("hiddenFileName").value = obj.videoLink;	
		
		document.getElementById("videoSectionDiv").style.display = "block";
		document.getElementById("hiddenFileName").value = videoLink;	
		plot.innerHTML = "<div align='center'><video width='60%' poster='../img/frame.jpg' controls preload='auto'><source src='"+videoLink+"' type='video/mp4'></video>";
	} else {
		plot.innerHTML = "";
		document.getElementById("hiddenFileName").value = "";	
		document.getElementById("videoSectionDiv").style.display = "none";
	}
}

/**
 * Function call if the instruction data is not in db 
 * for the selected profile and related page
 * @param null
 */
function populateNone(){
	/*$('#editor1').htmlarea('html', '&nbsp;');	*/
	CKEDITOR.instances['editor1'].insertHtml('&nbsp;');
	document.getElementById("instructionUpdateDiv").setAttribute("style", "display:none");
	document.getElementById("instructionAddDiv").setAttribute("style", "display:block");
	var userProfileSelect = document.getElementById("inputUserProfileInpt");
	userProfileSelect.removeAttribute("disabled", "true");
	var relatedPage = document.getElementById("inputPageInpt");
	relatedPage.removeAttribute("disabled", "true");
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
	document.getElementById('inputPageInpt').setAttribute("disabled", "false");
}

/**
 * Function to save the
 * instruction to the database
 * @param null
 */
function saveInstruction() {
	var insAreaText = CKEDITOR.instances.editor1.getData();
	var userProfile = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value;
	var userProfileText = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].text;	
	var relatedPageVal = document.getElementById("inputPageInpt").options[document.getElementById("inputPageInpt").selectedIndex].text;	
	if(validateInstruction(insAreaText)) {
		var newUserProf = Spine.Model.sub();
		newUserProf.configure("/admin/contentconfig/saveInstruction", "insAreaTextVal", "userProfileVal", "userProfileId", "relatedPageText", "createdBy");
		newUserProf.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new newUserProf({  
			insAreaTextVal: insAreaText,
			userProfileVal: userProfileText,
			userProfileId: userProfile,
			relatedPageText: relatedPageVal,
			createdBy: sessionStorage.getItem("jctEmail")
		});
		modelPopulator.save(); //POST
		newUserProf.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
		});
		newUserProf.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			//var obj = jQuery.parseJSON(jsonStr);
			var obj = JSON.parse(jsonStr);
			var statusCode = obj.statusCode;
			if(statusCode == 200) {
				//alertify.set({ delay: 2700 });
				alertify.alert("Instruction has been successfully created.");
				cancelInstruction();
			} else if(statusCode == 500) {
				// Error Message
			}
		});
	}
}


/**
 * Function to update the
 * instruction to the database
 * NOT IN USE
 * @param instructionId
 */
function updateInstruction(instructionId) {
	var insAreaText = CKEDITOR.instances.editor1.getData();			
	if(validateInstruction(insAreaText)) {
		var newUserProf = Spine.Model.sub();
		newUserProf.configure("/admin/contentconfig/updateInstruction", "insAreaTextVal", "instructionId");
		newUserProf.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new newUserProf({  
			insAreaTextVal: insAreaText,
			instructionId: instructionId
		});
		modelPopulator.save(); //POST
		newUserProf.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
		});
		newUserProf.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			//var obj = jQuery.parseJSON(jsonStr);
			var obj = JSON.parse(jsonStr);
			var statusCode = obj.statusCode;
			if(statusCode == 200) {
				alertify.alert("Instruction has been updated successfully created.");
				cancelInstruction();
			} else if(statusCode == 500) {
				// Error Message
			}
		});
	}
}



/**
 * Function to validate 
 * instruction data
 * @param instructionDesc
 */
function validateInstruction(instructionDesc){	
	if(document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value == ""){
		alertify.alert("Please select User Profile.");
		return false;
	}
	if(document.getElementById("inputPageInpt").options[document.getElementById("inputPageInpt").selectedIndex].value == ""){
		alertify.alert("Please select Related Page.");
		return false;
	}
	if (null == instructionDesc || instructionDesc == "" || instructionDesc == "&nbsp;"){
		alertify.alert("Please enter instuction data.");
		return false;
	}
	return true;
}

/**
 * Function call if user click on cancel button
 * to set the page as previous state
 * @param null
 */
function cancelInstruction() {
	var userProfileSelect = document.getElementById("inputUserProfileInpt");
	userProfileSelect.removeAttribute("disabled", "true");	
	CKEDITOR.instances.editor1.setData('');
	//document.getElementById('inputPageInpt').value = "";
	document.getElementById("inputPageInpt").selectedIndex = "0";
	document.getElementById('inputPageInpt').setAttribute("disabled", "true");
	document.getElementById("instructionUpdateDiv").setAttribute("style", "display:none");
	document.getElementById("instructionAddDiv").setAttribute("style", "display:block");
	fetchDropDownValues();

	$('#incVideo').prop('checked', false);
		
	var plot = document.getElementById("videoSection");	
	plot.innerHTML = "";	
	plot.innerHTML = "";		
	document.getElementById("videoSectionDiv").style.display = "none";	
}




/**
 * Function to check the extension of the uploaded file
 * @param null
 */
function validateVideoFile() {
	document.getElementById("createdBy").value = sessionStorage.getItem("jctEmail");
	var parts = document.getElementById('filename').value.split('.');
	var userProfile = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].value;
	var relatedPage = document.getElementById("inputPageInpt").options[document.getElementById("inputPageInpt").selectedIndex].value;
	var checkbox = document.getElementById("incVideo");
	
	if( userProfile == 0 ){
		alertify.alert('Please select the user profile.');
		return false;
	} else {
		document.getElementById("hiddenProfileInput").value = userProfile;
		document.getElementById("hiddenProfileValInput").value = document.getElementById("inputUserProfileInpt").options[document.getElementById("inputUserProfileInpt").selectedIndex].text;
	}
	if( relatedPage == 0 ){
		alertify.alert('Please select Related Page.');
		return false;
	} else {
		document.getElementById("hiddenPageInput").value = relatedPage;
	}
	
	
	var instructionDesc = CKEDITOR.instances.editor1.getData().trim();
	if (null == instructionDesc || instructionDesc == "" || instructionDesc == "&nbsp;" || instructionDesc.length == 0 ){
		alertify.alert("Please enter instuction data.");
		return false;
	}
	
	if ((null == instructionDesc || instructionDesc == "" || instructionDesc == "&nbsp;" ) && (null == document.getElementById('filename').value || document.getElementById('filename').value == "")){
		alertify.alert("Either enter text or upload video to continue.");
		return false;
	}
	if (checkbox.checked) {
		if((null != document.getElementById('filename').value) && (document.getElementById('filename').value != "")) {
			if (parts[parts.length-1] != 'mp4') {
				alertify.alert('Please upload mp4 format only.');
			    return false;
			} else if (storedFileSize > 200) {
				alertify.alert('Video file is too big.');
				return false;
			}
			document.getElementById("hiddenFileName").value = document.getElementById('filename').value;
			return true;
		} else {
			alertify.alert('Please select a video file.');
			return false;
		}
	}
	return true;
}

function toogleCheckbox() {
	var checkbox = document.getElementById("incVideo");
	if (checkbox.checked) {
		document.getElementById("videoDivId").style.display = "block";
		document.getElementById("videoSectionDiv").style.display = "block";
	} else {
		document.getElementById("videoDivId").style.display = "none";
		document.getElementById("videoSectionDiv").style.display = "none";
		
	}
}

var storedFileSize = 0;
function getFileSize(fileid) {
	storedFileSize = 0;
	try {
		var fileSize = 0;
		fileSize = $("#" + fileid)[0].files[0].size; //size in kb
		fileSize = fileSize / 1048576; //size in mb 
		storedFileSize = fileSize;
	} catch (e) {
	
	}
}