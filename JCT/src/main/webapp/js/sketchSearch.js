var firstRslt;
var maxRslt;
var imageText;
var nextLine;

/**
 * Function disables the search link depending upon page
 */
$(document).ready(function() {
	var myPage = sessionStorage.getItem("myPage");
	if (myPage == "BS") {
		document.getElementById("fancybox-manual-b").style.display = "block";
		document.getElementById("fancybox-manual-a").style.display = "none";
	} else {
		document.getElementById("fancybox-manual-b").style.display = "block";
		document.getElementById("fancybox-manual-a").style.display = "block";
	}
});

/**
 * function is used search the before sketch diagram
 * for the specific job title.
 */
function searchBeforeSketch(){
	if (document.getElementById("occupationBS")) {
		document.getElementById("occupationBS").value = "";
	}
	if (sessionStorage.getItem("occupationTitle")) {
		sessionStorage.removeItem("occupationTitle");
		sessionStorage.setItem("occupationTitle", "All Occupations");
	}
	document.getElementById('searchingPanel').style.display="block";
	document.getElementById("headerDiv").style.display="none";
	document.getElementById('imageDiv').style.display="block";
	document.getElementById('imageEnlargeDiv').style.display="none";
	document.getElementById('imageDiv').innerHTML = "";
	if(sessionStorage.getItem("firstResult")){
		sessionStorage.removeItem("firstResult");
	}
	if(sessionStorage.getItem("maxResults")){
		sessionStorage.removeItem("maxResults");
	}
	imageText = "<br />";
	nextLine = 0;
	$("#fancybox-manual-b").click(function() {
		$.fancybox.open({
			href : '#sketchSearchDiv',
			closeClick : false,
			openEffect : 'elastic',
			openSpeed  : 150,
			closeEffect : 'elastic',
			closeSpeed  : 150,
			padding : 5,
			closeClick : true,
			modal: true,
			showCloseButton : true,
			afterShow : function() {
				$('.fancybox-skin').append('<a title="Close" class="fancybox-item fancybox-close" href="javascript:jQuery.fancybox.close();"></a>');
			}
		});
	});
	if(null == sessionStorage.getItem("firstResult")){
		firstRslt = 0;
		sessionStorage.setItem("firstResult", firstRslt);
		
	}else{
		firstRslt = parseInt(sessionStorage.getItem("firstResult")) + 9;
	}
	if(null == sessionStorage.getItem("maxResults")){
		maxRslt = 1;
		sessionStorage.setItem("maxResults", maxRslt);
	}else{
		maxRslt = parseInt(sessionStorage.getItem("maxResults")) + 9;
	}
	searchAllDiagram('BS');
}

function searchAllDiagram(distinction) {
	
	document.getElementById('bsTable').style.display="none";
	document.getElementById('bsTable2').style.display="block";
	document.getElementById('asAvailability').style.display="none";				//		no button on first job-search page
	document.getElementById('bsAvailability1').style.display="none";
	document.getElementById('popupBackButton').innerHTML="&nbsp;";
	document.getElementById("bsSearchButtonId").disabled=true;
	imageText = "";
	
	// Before Sketch
	if (distinction == "BS") {
		document.getElementById('imageDiv').innerHTML = "";
	} else {
		document.getElementById('ASimageDiv').innerHTML = "";
	}
	
	//Create a model
	var beforeSketch = Spine.Model.sub();
	beforeSketch.configure("/user/search/searchAllDiagrams", "distinction", "firstResult", "maxResults");
	beforeSketch.extend( Spine.Model.Ajax );
	
	//Populate the model
	var populateModel = new beforeSketch({  
		distinction: distinction,
		firstResult: sessionStorage.getItem("firstResult"),
		maxResults: maxRslt
	});
	populateModel.save(); //POST
	
	beforeSketch.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	
	beforeSketch.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		if (distinction == "BS") {
			plotBeforeSketch (obj);
		} else {
			plotAfterSketch (obj);
		}
	});
	
}

function plotBeforeSketch(obj) {
	document.getElementById("headerDiv").style.display = "block";
	document.getElementById("bsSearchButtonId").disabled=false;
	if(obj.statusCode == 200){
		sessionStorage.setItem("totalDiagramCount",obj.totalDiagramCount);
		var baseObj = obj.sketchStringList;
		
		var size = baseObj.length;
			if(size > 0){
			for(var index=0; index<size; index++){
				nextLine = nextLine + 1;
				imageText = imageText + "<div class='col-md-4'><p style='text-align: center;'>"+baseObj[index+2]+"</p>&nbsp;&nbsp;&nbsp;&nbsp;<img src="+baseObj[index]+" class='thumbnails_search' onclick='openImage("+baseObj[index]+",\""+baseObj[index+1]+"\",\""+baseObj[index+1]+"\")'/> &nbsp;&nbsp;&nbsp;&nbsp;</div>";
				index = index + 2;
			}
			
			/** pagination starts**/
			var indFirst = parseInt(sessionStorage.getItem("firstResult")) + 1;
			var indSecond = 0;
			if(sessionStorage.getItem("firstResult") == 0){
				indSecond = 9;
			} else {
				var total = parseInt(sessionStorage.getItem("totalDiagramCount"));
				indSecond = parseInt(sessionStorage.getItem("firstResult")) + 9;
				if(indSecond >= total){
					indSecond = parseInt(sessionStorage.getItem("firstResult")) + (parseInt(size) / 3);
				} 
			}
			if(indSecond > (parseInt(size) / 3)){
				indSecond = parseInt(sessionStorage.getItem("totalDiagramCount"));
			}
			/** pagination ends **/
			document.getElementById("imageDiv").style.display="block";
			document.getElementById('imageDiv').innerHTML = imageText;
			var itemCount = "("+indFirst+" to "+indSecond + " out of "+parseInt(sessionStorage.getItem("totalDiagramCount"))+" records)";
			var plottingString = itemCount+"<p><div class='single_form_item'><p><input type='button' class='next_prev_btn btn_admin btn_admin-primary btn_admin-sm search_btn prev_btn_remove_diagram' id='prevBtnId' value='<< PREV' onclick='backToPrevious(\"BS\")' /> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type='button' id='nextBtnId' class='btn_admin btn_admin-primary btn_admin-sm search_btn next_btn_remove_diagram' value='NEXT >>' onclick='goToNext(\"BS\")' /></p></div></div></p>";
			document.getElementById('contoller').style.display="block";
			document.getElementById('contoller').innerHTML = plottingString;
			
			var title = sessionStorage.getItem("occupationTitle");
			if (null == title)
				title = "All Occupations";
			document.getElementById("bsHeaderInitial").innerHTML = title;
			
			
			if(sessionStorage.getItem("firstResult") == 0){
				$('#prevBtnId').attr('disabled','disabled');
			} 
			if ((parseInt(sessionStorage.getItem("firstResult")) + (parseInt(size) / 3)) == (parseInt(sessionStorage.getItem("totalDiagramCount")))) {
				$('#nextBtnId').attr('disabled','disabled');
			} else if ((parseInt(sessionStorage.getItem("firstResult")) + (parseInt(size) / 3)) > (parseInt(sessionStorage.getItem("totalDiagramCount")))) {
				$('#nextBtnId').attr('disabled','disabled');
			}
		} else {
			document.getElementById("bsHeaderInitial").innerHTML = sessionStorage.getItem("occupationTitle");
			document.getElementById("headerDiv").style.display="block";
			document.getElementById('imageDiv').style.display="block";
			document.getElementById('imageDiv').innerHTML  = "<br /><br /><br /><div align='center'><img src='../img/not-found.png'><br />No examples for this occupation yet.</div>";
		}
			/** DEFECT ID: 371 **/
			document.getElementById("bsSearchButtonId").disabled=false;
	} else {
		/** DEFECT ID: 371 **/
		document.getElementById("bsSearchButtonId").disabled=false;
	}
}

function plotAfterSketch(obj) {
	document.getElementById("ASheaderDiv").style.display = "block";
	document.getElementById("asSearchButtonId").disabled=false;
	if(obj.statusCode == 200){
		sessionStorage.setItem("totalDiagramCount",obj.totalDiagramCount);
		var baseObj = obj.sketchStringList;
		var size = baseObj.length;
			if(size > 0){
			for(var index=0; index<size; index++){
				nextLine = nextLine + 1;
				//imageText = imageText + "<img src="+baseObj[index]+" class='thumbnails' onclick='openImageAS("+baseObj[index]+",\""+baseObj[index+1]+"\",\""+baseObj[index+1]+"\")'/> &nbsp;&nbsp;&nbsp;&nbsp;";
				imageText = imageText + "<div class='col-md-4'><p style='text-align: center;'>"+baseObj[index+2]+"</p>&nbsp;&nbsp;&nbsp;&nbsp;<img src="+baseObj[index]+" class='thumbnails_search' onclick='openImageAS("+baseObj[index]+",\""+baseObj[index+1]+"\",\""+baseObj[index+1]+"\")'/> &nbsp;&nbsp;&nbsp;&nbsp;</div>";
				/*if(nextLine%3 == 0){
					imageText = imageText + "<br /><br />";
				}*/
				index = index + 2;
			}
			
			/** pagination starts**/
			var indFirst = parseInt(sessionStorage.getItem("firstResult")) + 1;
			var indSecond = 0;
			if(sessionStorage.getItem("firstResult") == 0){
				indSecond = 9;
			} else {
				var total = parseInt(sessionStorage.getItem("totalDiagramCount"));
				indSecond = parseInt(sessionStorage.getItem("firstResult")) + 9;
				if(indSecond >= total){
					indSecond = parseInt(sessionStorage.getItem("firstResult")) + (parseInt(size) / 3);
				} 
			}
			if(indSecond > (parseInt(size) / 3)){
				indSecond = parseInt(sessionStorage.getItem("totalDiagramCount"));
			}
			/** pagination ends **/
			
			document.getElementById("ASimageDiv").style.display="block";
			document.getElementById('ASimageDiv').innerHTML = imageText;
			var itemCount = "("+indFirst+" to "+indSecond + " out of "+parseInt(sessionStorage.getItem("totalDiagramCount"))+" records)";
			var plottingString = itemCount+"<P><div class='single_form_item'><p><input type='button' class='next_prev_btn btn_admin btn_admin-primary btn_admin-sm search_btn prev_btn_remove_diagram' id='prevBtnIdAS' value='<< PREV' onclick='backToPrevious(\"AS\")' /> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type='button' id='nextBtnIdAS' class='btn_admin btn_admin-primary btn_admin-sm search_btn next_btn_remove_diagram' value='NEXT >>' onclick='goToNext(\"AS\")' /></p></div></div></P>";
			document.getElementById('AScontoller').style.display="block";
			document.getElementById('AScontoller').innerHTML = plottingString;
			var title = sessionStorage.getItem("occupationTitle");
			if (null == title)
				title = "All Occupations";
			document.getElementById("asHeaderInitial").innerHTML = title;
			
			if(sessionStorage.getItem("firstResult") == 0){
				$('#prevBtnIdAS').attr('disabled','disabled');
			} 
			if ((parseInt(sessionStorage.getItem("firstResult")) + (parseInt(size) / 3)) == (parseInt(sessionStorage.getItem("totalDiagramCount")))) {
				$('#nextBtnIdAS').attr('disabled','disabled');
			} else if ((parseInt(sessionStorage.getItem("firstResult")) + (parseInt(size) / 3)) > (parseInt(sessionStorage.getItem("totalDiagramCount")))) {
				$('#nextBtnIdAS').attr('disabled','disabled');
			}
		} else {
			document.getElementById("asHeaderInitial").innerHTML = sessionStorage.getItem("occupationTitle");
			document.getElementById("ASheaderDiv").style.display="block";
			document.getElementById('ASimageDiv').style.display="block";
			document.getElementById('ASimageDiv').innerHTML  = "<br /><br /><br /><div align='center'><img src='../img/not-found.png'><br />No examples for this occupation yet.</div>";
		}
			/** DEFECT ID: 371 **/
			document.getElementById("asSearchButtonId").disabled=false;
	} else {
		/** DEFECT ID: 371 **/
		document.getElementById("asSearchButtonId").disabled=false;
	}

}

function goToNext(distinction) {
	var indSecond = parseInt(sessionStorage.getItem("firstResult")) + 9;
	sessionStorage.setItem("firstResult",indSecond);
	document.getElementById('contoller').innerHTML = "";
	searchAllDiagram(distinction);
}

function backToPrevious(distinction) {	
	var indSecond = parseInt(sessionStorage.getItem("firstResult")) - 9;
	sessionStorage.setItem("firstResult",indSecond);
	searchAllDiagram(distinction);
}
/**
 * Function fetches and populate the job level drop down on the searching lightbox
 * @param type
 */
/*function populateDropDown(type) {
	//Create a model
	var master = Spine.Model.sub();
	master.configure("/user/auth/populateSignupMasterData", "none");
	master.extend( Spine.Model.Ajax );
	
	//Populate the model
	var populateMaster = new master({  
		none: ""
	});
	populateMaster.save(); //POST
	master.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to fetch the dropdown values.");
	});
	
	master.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if(statusCode == "200"){
			var index = 1;
			var jobLevelSelect = document.getElementById("inputJobLevel"+type);
			jobLevelSelect.innerHTML = "";
			for (var key in obj.jobLevel) {
				//Create a option element dynamically
				var option2 = document.createElement("option");
				if(index == 1) {
					var option = document.createElement("option");
					option.text = "Select Job Level";
				    option.value = "";
				    option.className = "form-control-general";
				    try {
				    	jobLevelSelect.add(option, null); //Standard 
				    } catch(error) {
				    	//jobLevelSelect.add(option); // IE only
				    }
				} 
					option2.text = obj.jobLevel[key];
				    option2.value = obj.jobLevel[key];
				    option2.className = "form-control-general";
			    try {
			    	jobLevelSelect.add(option2, null); //Standard 
			    } catch(error) {
			    	//jobLevelSelect.add(option); // IE only
			    }
			    index = index + 1;
			}
			
		}
	});
}*/
/**
 * Function searhes the before sketch diagram
 */
/*function searchBeforeSketchDiagram() {	
	*//** DEFECT ID: 371 **//*
	var jobLevel = document.getElementById("inputJobLevelBS").options[document.getElementById("inputJobLevelBS").selectedIndex].value;
	
	if(jobLevel == "Select Job Level" || jobLevel == ""){	//********************** If No Job-level is selected **************************
		document.getElementById("headerDiv").style.display="none";
		document.getElementById("imageDiv").style.display="none";
		document.getElementById("imageBS").innerHTML="Please select search Criteria";
		
		return false;	
	}
	else{
	document.getElementById('imageBS').innerHTML = "<img src='../img/Processing.gif'/>";
	document.getElementById('bsTable').style.display="none";
	document.getElementById('bsTable2').style.display="block";
		
	document.getElementById('asAvailability').style.display="none";				//		no button on first job-search page
	document.getElementById('bsAvailability1').style.display="none";
	//document.getElementById('popupBackButton').style.display="none";
	document.getElementById('popupBackButton').innerHTML="&nbsp;";
	document.getElementById("bsSearchButtonId").disabled=true;
	imageText = "<br />";
	document.getElementById('imageDiv').innerHTML = "";
	//var jobLevel = document.getElementById("inputJobLevelBS").options[document.getElementById("inputJobLevelBS").selectedIndex].value;
	sessionStorage.setItem("selectedJobTitle", jobLevel);
	//Create a model
	var beforeSketch = Spine.Model.sub();
	beforeSketch.configure("/user/search/searchBeforeSketch", "jobTitle", "firstResult", "maxResults");
	beforeSketch.extend( Spine.Model.Ajax );
	if(null == sessionStorage.getItem("firstResult")){
		firstRslt = 0;
	}else{
		firstRslt = parseInt(sessionStorage.getItem("maxResults")) + 1;
	}
	if(null == sessionStorage.getItem("maxResults")){
		maxRslt = 20;
	}else{
		maxRslt = parseInt(sessionStorage.getItem("maxResults")) + 20;
	}
	sessionStorage.setItem("firstResult", firstRslt);
	sessionStorage.setItem("maxResults", maxRslt);
	//Populate the model
	var populateModel = new beforeSketch({  
		//jobTitle: sessionStorage.getItem("jobTitle"),
		jobTitle: jobLevel,
		firstResult: firstRslt,
		maxResults: maxRslt
	});
	populateModel.save(); //POST
	
	beforeSketch.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	
	beforeSketch.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		document.getElementById('imageBS').innerHTML = "";
		document.getElementById("headerDiv").style.display = "block";
		if(obj.statusCode == 200){
			var baseObj = obj.sketchStringList;
			var size = baseObj.length;
				if(size > 0){
				for(var index=0; index<size; index++){
					nextLine = nextLine + 1;
					imageText = imageText + "<img src="+baseObj[index]+" class='thumbnails' onclick='openImage("+baseObj[index]+",\""+baseObj[index+1]+"\",\""+baseObj[index+2]+"\")'/> &nbsp;&nbsp;&nbsp;&nbsp;";
					if(nextLine%3 == 0){
						imageText = imageText + "<br /><br /><br />";
					}
					index = index + 2;
				}
				document.getElementById("imageDiv").style.display="block";
				document.getElementById('bsHeaderInitial').innerHTML = sessionStorage.getItem("selectedJobTitle");
				document.getElementById('imageBS').innerHTML = "";
				document.getElementById('imageDiv').innerHTML = imageText;
				//document.getElementById('contoller').innerHTML = "<< Previous &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Next >>";
			} else {
				document.getElementById("headerDiv").style.display="block";
				document.getElementById('bsHeaderInitial').innerHTML = sessionStorage.getItem("selectedJobTitle");
				document.getElementById('imageBS').innerHTML  = "No results to display .";
			}
				*//** DEFECT ID: 371 **//*
				document.getElementById("bsSearchButtonId").disabled=false;
		}else{
			alert('handle the exception');
			*//** DEFECT ID: 371 **//*
			document.getElementById("bsSearchButtonId").disabled=false;
		}
	});
	}		//		my else
}*/
/**
 * Function searches the after sketch diagram
 */
/*function searchAfterSketchDiagram() {
	var jobLevel = document.getElementById("inputJobLevelAS").options[document.getElementById("inputJobLevelAS").selectedIndex].value;
	
	if(jobLevel == "Select Job Level" || jobLevel == ""){	//********************** If No Job-level is selected **************************
		document.getElementById("ASheaderDiv").style.display="none";
		document.getElementById("ASimageDiv").style.display="none";
		document.getElementById("imageAS").innerHTML="Please select search Criteria";
		
		return false;	
	}
	else{
		document.getElementById('imageAS').innerHTML = "<img src='../img/Processing.gif'/>";
		document.getElementById('asTable').style.display="none";
		document.getElementById('asTable2').style.display="block";
		
		*//** DEFECT ID: 371 **//*
		document.getElementById("asSearchButtonId").disabled = true;
		imageText = "<br />";
		//document.getElementById('ASimageDiv').innerHTML = "";
		if(document.getElementById('ASimageDiv')) {
			document.getElementById('ASimageDiv').innerHTML = "";
		}
		
			
		sessionStorage.setItem("selectedJobTitle", jobLevel);
		//Create a model
		var afterSketch = Spine.Model.sub();
		afterSketch.configure("/user/search/searchAfterSketch", "jobTitle", "firstResult", "maxResults");
		afterSketch.extend( Spine.Model.Ajax );
		if(null == sessionStorage.getItem("firstResult")){
			firstRslt = 0;
		}else{
			firstRslt = parseInt(sessionStorage.getItem("maxResults")) + 1;
		}
		if(null == sessionStorage.getItem("maxResults")){
			maxRslt = 20;
		}else{
			maxRslt = parseInt(sessionStorage.getItem("maxResults")) + 20;
		}
		sessionStorage.setItem("firstResult", firstRslt);
		sessionStorage.setItem("maxResults", maxRslt);
		//Populate the model
		var populateModel = new afterSketch({  
			//jobTitle: sessionStorage.getItem("jobTitle"),
			jobTitle: jobLevel,
			firstResult: firstRslt,
			maxResults: maxRslt
		});
		populateModel.save(); //POST
		
		afterSketch.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
		});
		
		afterSketch.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			document.getElementById('imageAS').innerHTML = "";
			document.getElementById("ASimageDiv").style.display="block";
			document.getElementById("ASheaderDiv").style.display = "block";
			if(obj.statusCode == 200){
				var baseObj = obj.sketchStringList;
				var size = baseObj.length;
				if(size > 0){
					for(var index=0; index<size; index++){
						nextLine = nextLine + 1;
						imageText = imageText + "<img src="+baseObj[index]+" class='thumbnails' onclick='openImageAS("+baseObj[index]+",\""+baseObj[index+1]+"\",\""+baseObj[index+2]+"\")'/> &nbsp;&nbsp;&nbsp;&nbsp;";
						if(nextLine%3 == 0){
							imageText = imageText + "<br /><br /><br />";
						}
						index = index + 2;
					}
					
					document.getElementById('asHeaderInitial').innerHTML = sessionStorage.getItem("selectedJobTitle");
					document.getElementById('imageAS').innerHTML = "";
					document.getElementById('ASimageDiv').innerHTML = imageText;
				} else {
					document.getElementById('asHeaderInitial').innerHTML = sessionStorage.getItem("selectedJobTitle");				
					document.getElementById('imageAS').innerHTML  = "No results to display .";
				}
				*//** DEFECT ID: 371 **//*
				document.getElementById("asSearchButtonId").disabled = false;
			}else{
				alert('handle the exception');
				
				*//** DEFECT ID: 371 **//*
				document.getElementById("asSearchButtonId").disabled = false;
			}
		});
	}		//		my else
	
}*/

/**
 * function is used search the before sketch diagram
 * for the specific job title.
 * 
 * JUSTIN FEEDBACK
 */
function searchAfterSketch() {
	if (document.getElementById("occupationAS")) {
		document.getElementById("occupationAS").value = "";
	}
	if (sessionStorage.getItem("occupationTitle")) {
		sessionStorage.removeItem("occupationTitle");
		sessionStorage.setItem("occupationTitle", "All Occupations");
	}
	document.getElementById('ASsearchingPanel').style.display="block";
	document.getElementById("ASheaderDiv").style.display="none";
	document.getElementById('ASimageDiv').style.display="block";
	document.getElementById('ASimageEnlargeDiv').style.display="none";
	document.getElementById('ASimageDiv').innerHTML = "";
	if(sessionStorage.getItem("firstResult")){
		sessionStorage.removeItem("firstResult");
	}
	if(sessionStorage.getItem("maxResults")){
		sessionStorage.removeItem("maxResults");
	}
	imageText = "<br />";
	nextLine = 0;
	$("#fancybox-manual-a").click(function() {
		$.fancybox.open({
			href : '#ASsketchSearchDiv',
			closeClick : false,
			openEffect : 'elastic',
			openSpeed  : 150,
			closeEffect : 'elastic',
			closeSpeed  : 150,
			padding : 5,
			closeClick : true,
			modal: true,
			showCloseButton : true,
			afterShow : function() {
				$('.fancybox-skin').append('<a title="Close" class="fancybox-item fancybox-close" href="javascript:jQuery.fancybox.close();"></a>');
			}
		});
	});
	if(null == sessionStorage.getItem("firstResult")){
		firstRslt = 0;
		sessionStorage.setItem("firstResult", firstRslt);
		
	}else{
		firstRslt = parseInt(sessionStorage.getItem("firstResult")) + 9;
	}
	if(null == sessionStorage.getItem("maxResults")){
		maxRslt = 1;
		sessionStorage.setItem("maxResults", maxRslt);
	}else{
		maxRslt = parseInt(sessionStorage.getItem("maxResults")) + 9;
	}
	searchAllDiagram('AS');	
}
/**
 * Function opens the selected before sketch image
 * @param img
 * @param jobReferenceNo
 * @param availability
 */
function openImage(img, jobReferenceNo, availability){
	document.getElementById('searchingPanel').style.display="none";
	document.getElementById('imageDiv').style.display="none";
	document.getElementById('contoller').style.display="none";
	document.getElementById('imageEnlargeDiv').style.display="block";
	document.getElementById('bsTable').style.display="block";
	document.getElementById('bsTable2').style.display="none";
	document.getElementById('bsHeader').innerHTML = sessionStorage.getItem("selectedJobTitle");
	
	//document.getElementById('backButton').innerHTML = "<input type='button' class='popupBackButton' onClick='goBack()' value='Back'>";
	document.getElementById('mainPic').innerHTML = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br /><img src="+img+" class='thumbnailPopup'/>";
	document.getElementById("asAvailability").style.display = "block";
	document.getElementById("bsAvailability1").style.display = "none";
	//Now check if the aftersketch diagram is available for the selected diagram
	var pageSequence = sessionStorage.getItem("pageSequence");
	var myPage = sessionStorage.getItem("myPage");
	document.getElementById("asAvailability").style.display = "block";			//			Changed, 14th july
	document.getElementById("popupBackButton").style.display = "block";
			
	var link = document.getElementById("asAvailability");
	link.onclick = function () {
		//disable the link
		document.getElementById("asAvailability").disabled = true;
		var master = Spine.Model.sub();
		master.configure("/user/search/getSelectedDiagrams", "jobReferenceNos", "distinction");
		master.extend( Spine.Model.Ajax );
				
		//Populate the model
		var populateMaster = new master({  
			jobReferenceNos: jobReferenceNo,
			distinction: "AS"
		});
		populateMaster.save(); //POST
		master.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to fetch after sketch diagram.");
			document.getElementById("asAvailability").disabled = false;
		});
		populateMaster.bind("ajaxSuccess", function(record, xhr, settings, error) {
			document.getElementById("asAvailability").disabled = false;
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			document.getElementById('mainPic').innerHTML = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br /><img src="+obj.selectedDiagram+" class='thumbnailPopup'/>";
			document.getElementById("asAvailability").style.display = "none";
			document.getElementById("bsAvailability1").style.display = "block";
			var link2 = document.getElementById("bsAvailability1");
			
			link2.onclick = function () {
				document.getElementById("bsAvailability1").disabled = true;
				var master = Spine.Model.sub();
				master.configure("/user/search/getSelectedDiagrams", "jobReferenceNos", "distinction");
				master.extend( Spine.Model.Ajax );
				
				//Populate the model
				var populateMaster = new master({  
					jobReferenceNos: jobReferenceNo,
					distinction: "BS"
				});
				populateMaster.save(); //POST
				master.bind("ajaxError", function(record, xhr, settings, error) {
					alertify.alert("Unable to fetch after sketch diagram.");
					document.getElementById("bsAvailability1").disabled = false;
				});
				populateMaster.bind("ajaxSuccess", function(record, xhr, settings, error) {
					var jsonStr = JSON.stringify(xhr);
					var obj = jQuery.parseJSON(jsonStr);
					document.getElementById('mainPic').innerHTML = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br /><img src="+obj.selectedDiagram+" class='thumbnailPopup'/>";
					document.getElementById("asAvailability").style.display = "block";
					document.getElementById("bsAvailability1").style.display = "none";
					document.getElementById("bsAvailability1").disabled = false;
				});
			};
					
		});
	};
}
/**
 * Function opens the selected after sketch image
 * @param img
 * @param jobReferenceNo
 * @param availability
 */
function openImageAS(img, jobReferenceNo, availability){
	document.getElementById('ASsearchingPanel').style.display="none";
	document.getElementById('ASimageDiv').style.display="none";
	document.getElementById('AScontoller').style.display="none";
	document.getElementById('ASimageEnlargeDiv').style.display="block";
	document.getElementById('asTable').style.display="block";
	document.getElementById('asTable2').style.display="none";
	document.getElementById('asHeader').innerHTML = sessionStorage.getItem("selectedJobTitle");
	document.getElementById('ASmainPic').innerHTML = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br /><img src="+img+" class='thumbnailPopup'/>";
	document.getElementById("bsAvailability").style.display = "block";
	document.getElementById("asAvailability1").style.display = "none";
	document.getElementById("bsAvailability").style.display = "block";			//			Changed, 14th july
	document.getElementById("ASpopupBackButton").style.display = "block";
	document.getElementById("bsAvailability").disabled = false;
	
	var link = document.getElementById("bsAvailability");
	link.onclick = function () {
		//disable the link
		document.getElementById("bsAvailability").disabled = true;
		var master = Spine.Model.sub();
		master.configure("/user/search/getSelectedDiagrams", "jobReferenceNos", "distinction");
		master.extend( Spine.Model.Ajax );
				
		//Populate the model
		var populateMaster = new master({  
			jobReferenceNos: jobReferenceNo,
				distinction: "BS"
		});
		populateMaster.save(); //POST
		master.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to fetch after sketch diagram.");
			document.getElementById("asAvailability").disabled = false;
		});
		populateMaster.bind("ajaxSuccess", function(record, xhr, settings, error) {
			document.getElementById("asAvailability1").disabled = false;
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			document.getElementById('ASmainPic').innerHTML = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br /><img src="+obj.selectedDiagram+" class='thumbnailPopup'/>";
			document.getElementById("bsAvailability").style.display = "none";
			document.getElementById("asAvailability1").style.display = "block";
			var link2 = document.getElementById("asAvailability1");
			link2.onclick = function () {
				document.getElementById("asAvailability1").disabled = true;
				var master = Spine.Model.sub();
				master.configure("/user/search/getSelectedDiagrams", "jobReferenceNos", "distinction");
				master.extend( Spine.Model.Ajax );
				
				//Populate the model
				var populateMaster = new master({  
					jobReferenceNos: jobReferenceNo,
					distinction: "AS"
				});
				populateMaster.save(); //POST
				master.bind("ajaxError", function(record, xhr, settings, error) {
					alertify.alert("Unable to fetch after sketch diagram.");
					document.getElementById("asAvailability1").disabled = false;
				});
				populateMaster.bind("ajaxSuccess", function(record, xhr, settings, error) {
					var jsonStr = JSON.stringify(xhr);
					var obj = jQuery.parseJSON(jsonStr);
					document.getElementById('ASmainPic').innerHTML = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br /><img src="+obj.selectedDiagram+" class='thumbnailPopup'/>";
					document.getElementById("asAvailability1").style.display = "none";
					document.getElementById("bsAvailability").style.display = "block";
					document.getElementById("bsAvailability").disabled = false;
				});
			};
		});
	};
}
/**
 * Function helps to return to the main search window from selected before sketch image 
 * on which all the results are being displayed
 */
function goBack(){
	document.getElementById('searchingPanel').style.display="block";
	document.getElementById('imageDiv').style.display="block";
	document.getElementById('contoller').style.display="block";
	document.getElementById('imageEnlargeDiv').style.display="none";
	document.getElementById('bsTable').style.display="none";
	document.getElementById('bsTable2').style.display="block";
}
/**
 * Function helps to return to the main search window from selected after sketch image 
 * on which all the results are being displayed
 */
function goBackAS(){
	document.getElementById('ASsearchingPanel').style.display="block";
	document.getElementById('ASimageDiv').style.display="block";
	document.getElementById('AScontoller').style.display="block";
	document.getElementById('ASimageEnlargeDiv').style.display="none";
	document.getElementById('asTable').style.display="none";
	document.getElementById('asTable2').style.display="block";

}



//////////////////////////////// MY ACCOUNT //////////////////////////////////
function openMyAccountDetails() {
	var src = "Tool";
	if (sessionStorage.getItem("src")) {
		src = sessionStorage.getItem("src");
	}
	// Enable the open my account details feature only if the user has logged in from the tool
	if (src == "Tool") {
	
	$("#fancybox-manual-my-account").click(function() {
		if(document.getElementById("inputEPass")) {
			document.getElementById("inputEPass").value = "";
		}
		if(document.getElementById("inputNPass")) {
			document.getElementById("inputNPass").value = "";
		}
		if(document.getElementById("inputCPass")) {
			document.getElementById("inputCPass").value = "";
		}
		if(document.getElementById("passwordResetLink")) {
			document.getElementById("passwordResetLink").style.display = "block";
		}
		if(document.getElementById("passwordResetArea")) {
			document.getElementById("passwordResetArea").style.display = "none";
		}
		$.fancybox.open({
			href : '#myAccountDiv',
			closeClick : false,
			openEffect : 'elastic',
			openSpeed  : 150,
			closeEffect : 'elastic',
			closeSpeed  : 150,
			padding : 5,
			closeClick : true,
			modal: true,
			showCloseButton : true,
			afterShow : function() {
				$('.fancybox-skin').append('<a title="Close" class="fancybox-item fancybox-close" href="javascript:jQuery.fancybox.close();"></a>');
			}
		});
	});
	fetchMyAccountDetails();
	}
}
/**
 * Function fetches the account details of the user
 */
function fetchMyAccountDetails() {
	document.getElementById("myAcValDiv").innerHTML = "";
	document.getElementById("myAcValDiv").innerHTML = "<div align='center'><img src='../img/Processing.gif'/></div>";
	//Create a model
	var master = Spine.Model.sub();
	master.configure("/user/search/populateMyAccount", "emailId");
	master.extend( Spine.Model.Ajax );
	
	//Populate the model
	var populateMaster = new master({  
		emailId: sessionStorage.getItem("jctEmail")
	});
	populateMaster.save(); //POST
	
	master.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	
	master.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		if ( obj.statusCode == 200 ) {
			populateDetails(obj, "success");
		} else if ( obj.statusCode == 500 ) {
			populateDetails(obj, "failure");
		}
	});
}
/**
 * Function displays the account details
 * @param obj
 * @param distinction
 */
function populateDetails(obj, distinction) {
	var myAcValDiv = document.getElementById('myAcValDiv');
	if (distinction == "success") {
		document.getElementById("myAccountheaderDiv").innerHTML = "My Account";
		var tab = "<table width='90%'><tr><th>First Name: </th><td>"+obj.firstName+"</td></tr>"+
				  "<tr><th>Last Name: </th><td>"+obj.lastName+"</td></tr>"+
				  //"<tr><th>Region: </th><td>"+obj.region+"</td></tr>"+
				  //"<tr><th>Gender: </th><td>"+obj.sex+"</td></tr>"+
				  //"<tr><th>Tenure: </th><td>"+obj.tenure+"</td></tr>"+
				  //"<tr><th>Supervise People: </th><td>"+obj.supervisePeople+"</td></tr>"+
				  //"<tr><th>Function Group: </th><td>"+obj.functionGroup+"</td></tr>"+
				  //"<tr><th>Job Level: </th><td>"+obj.jobLevel+"</td></tr>"+
				  "<tr><th width='34%'>Account Expiration Date: </th><td>"+new Date(obj.jctAccountExpirationDate).toLocaleDateString()+"<input type='button' onclick='submitForm()' id='renewSubscription' value='Renew Subscription' class='btn btn-sm renew_subs_btn'></td></tr>" +
				  "<tr><td colspan='3' style='width:100%' class='myAcccuont_note_renew' id='renewSubscriptionNoteId'><b>Note:</b> Your current subscription was purchased for you by a facilitator, "+obj.faciFirstName+" "+obj.faciLastName+" ("+obj.faciEmail+"). You can purchase a renewal for yourself using the button above, or you can ask your facilitator to renew your subscription.</td></tr>"+
				  "<tr><th>Occupation Title: </th><td>"+obj.onetDesc+"</td></tr>" +
				  //"<tr><th>Customer Id: </th><td>"+obj.customerId+"</td></tr></table>";
				  "<tr><th>Email id: </th><td>"+sessionStorage.getItem("jctEmail")+"</td></tr></table>";
			
		myAcValDiv.innerHTML = tab;
	} else {
		myAcValDiv.innerHTML = "Something went wrong.";
	}
	var customerId = obj.customerId;
    var customerIdStr = customerId.substr(0, 2);
	if(customerIdStr == "99") {		 
		document.getElementById('renewSubscriptionNoteId').setAttribute("style", "display:");
	} else {
		document.getElementById('renewSubscriptionNoteId').setAttribute("style", "display:none");
	}
}

function searchOccupation(distinction) {
	var occupation = "";
	if (distinction == "BS") {
		occupation = document.getElementById("occupationBS").value;
		
		if (occupation.trim().length == 0) {
			alertify.alert("<img src='../img/alert-icon.png'><br /><p> Please provide search words that describe your occupation.</p>");
			document.getElementById("occupationBS").focus();
			return;
		} else {
			document.getElementById('contoller').style.display="none";		
			document.getElementById('contollerOccupation').style.display="none";
			occupation = occupation.replace(/,+/g, ' ');
			var val = occupation.replace(/ +/g, ',');
			if (val.indexOf(",") == 0) {
				val = val.substring(1, val.length);
			}
			val = val.replace(/([*()\+\\{}\[\]\|\/?])+/g, '');
			if (val == "") {
				//alertify.alert("Only special character is not valid.");
				//return false;
				document.getElementById("contollerOccupation").style.display = "none";
				document.getElementById("imageDiv").innerHTML="<div align='center'><br /><br /><img src='../img/not-found.png'><br /><p>The occupation you have entered is not in the system.<br />Please enter search words and select an occupation from the list provided.</p></div>";
				return false;
			} else {
				document.getElementById("imageDiv").innerHTML = "<br /><br /><br /><br /><br /><br /><div align='center'><img src='../img/Processing.gif'/></div>";
			}
		}
		
	} else {
		occupation = document.getElementById("occupationAS").value;
		if (occupation.trim().length == 0) {
			alertify.alert("<img src='../img/alert-icon.png'><br /><p> Please provide search words that describe your occupation.</p>");
			document.getElementById("occupationAS").focus();
			return;
		}
		document.getElementById('AScontoller').style.display="none";		
		document.getElementById('AScontollerOccupation').style.display="none";
		occupation = occupation.replace(/,+/g, ' ');
		var val = occupation.replace(/ +/g, ',');
		if (val.indexOf(",") == 0) {
			val = val.substring(1, val.length);
		}
		val = val.replace(/([*()\+\\{}\[\]\|\/?])+/g, '');
		if (val == "") {
			//alertify.alert("Only special character is not valid.");
			document.getElementById("AScontollerOccupation").style.display = "none";
			document.getElementById("ASimageDiv").innerHTML="<div align='center'><br /><br /><img src='../img/not-found.png'><br /><p>The occupation you have entered is not in the system.<br />Please enter search words and select an occupation from the list provided.</p></div>";
			return false;
		} else {
			document.getElementById("ASimageDiv").innerHTML = "<br /><br /><br /><br /><br /><br /><div align='center'><img src='../img/Processing.gif'/></div>";
		}
		
	}
	
	// Search the o* net list
	var stub = Spine.Model.sub();
	stub.configure("/user/auth/searchOccupationList", "searchString");
	stub.extend( Spine.Model.Ajax );
	occupation = occupation.replace(/,+/g, ' ');
	var val = occupation.replace(/ +/g, ',');
	if (val.indexOf(",") == 0) {
		val = val.substring(1, val.length);
	}
	val = val.replace(/([*()\+\\{}\[\]\|\/?])+/g, '');
	if (val == "") {
		alertify.alert("Only special character is not valid.");
		return false;
	}
	//Populate the model with data to transfer
	var modelPopulation = new stub({ 
		searchString: val
	});
	modelPopulation.save(); //POST the data
	
	stub.bind("ajaxError", function(record, xhr, settings, error) {
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
		alertify.alert("Unable to connect to the server.");
	});
	stub.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if (statusCode == "200") {
			plotDataNow(obj.dtoList, distinction);
		} else {
			if (distinction == "BS") {
				document.getElementById("contollerOccupation").style.display = "none";
				//document.getElementById("imageDiv").innerHTML="<div align='center'><br /><br /><img src='../img/not-found.png'><br /><p><b>"+occupation+"</b> "+obj.statusMsg+"</p></div>";
				document.getElementById("imageDiv").innerHTML="<div align='center'><br /><br /><img src='../img/not-found.png'><br /><p>"+obj.statusMsg+"</p></div>";
			} else {
				document.getElementById("AScontollerOccupation").style.display = "none";
				//document.getElementById("ASimageDiv").innerHTML="<div align='center'><br /><br /><img src='../img/not-found.png'><br /><p><b>"+occupation+"</b> "+obj.statusMsg+"</p></div>";
				document.getElementById("ASimageDiv").innerHTML="<div align='center'><br /><br /><img src='../img/not-found.png'><br /><p>"+obj.statusMsg+"</p></div>";
			}
		}
	});	
}

function plotDataNow(dtoList,distinction) {
	var plottingDiv = "";
	if (distinction == "BS") {
		if (document.getElementById("imageDiv")) {
			document.getElementById("imageDiv").innerHTML = "";
		}
		document.getElementById('contollerOccupation').style.display="block";
		plottingDiv = document.getElementById("imageDiv");
	} else {
		if (document.getElementById("ASimageDiv")) {
			document.getElementById("ASimageDiv").innerHTML = "";
		}
		document.getElementById('AScontollerOccupation').style.display="block";
		
		plottingDiv = document.getElementById("ASimageDiv");
	}
	
	var str = "<table width='75%' class='custom_table_onet'><thead><tr class='tab_header'></tr></thead><tbody>";
	
	for (var index = 0; index < dtoList.length; index++) {
		var trColor = "";
		var titleId = "title_" + index;
		var descId = "desc_" + index;
		//str = str + "<tr id='"+titleId+"' bgcolor="+trColor+"><td class='custom_td_onet' ><input type='radio' class='custom_radio_onet' name='occupation' value='"+dtoList[index].code+"' onclick='showDescription(\""+descId+"\", \""+dtoList.length+"\", \""+dtoList[index].title+"\")'>"+dtoList[index].title+"</td></tr><tr id= '"+descId+"' style='display:none;'><td class='onet_desc_custom_cls' >"+dtoList[index].desc+"</td></tr>";
		var showHideId = "show_hide_" + index; 
		str = str + "<tr id='"+titleId+"' bgcolor="+trColor+"><td class='custom_td_onet' ><input type='radio' onclick='storeTitleVal(\""+dtoList[index].title+"\")' class='custom_radio_onet' name='occupation' value='"+dtoList[index].code+"'>"+dtoList[index].title+"</td><td id='"+showHideId+"' class='hide_onet' onclick='showDescription(\""+descId+"\", \""+dtoList.length+"\", \""+showHideId+"\", \""+dtoList[index].title+"\")'></td></tr><tr id= '"+descId+"' style='display:none;'><td class='onet_desc_custom_cls' >"+dtoList[index].desc+"</td></tr>";
	}
	str = str + "</tbody></table>";
	plottingDiv.innerHTML = str;
}

function storeTitleVal(title) {
	sessionStorage.setItem("occupationTitle", title);
}
/*function showDescription(descId, length, title) {
	if (undefined != $("input[name=occupation]:checked").val()) {
		for(var i=0; i<length; i++) {
			var tempId="desc_"+i;
			if (document.getElementById(tempId).style.display == "block") {
				document.getElementById(tempId).style.display = "none";
				break;
			}
		}
		sessionStorage.setItem("occupationTitle", title);
		document.getElementById(descId).style.display = "block";
		
	} else {
		alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please select an occupation from the list</p>");
		return false;
	}
	
}*/



function showDescription(descId, length, showHideId, title) {
	//if (undefined != $("input[name=occupation]:checked").val()) {
		if($('#'+showHideId).attr('class') == "hide_onet") {
			$("#" + showHideId).removeClass("hide_onet");
			$("#" + showHideId).addClass("show_onet");
			document.getElementById(descId).style.display = "block";
		} else {
			$("#" + showHideId).removeClass("show_onet");
			$("#" + showHideId).addClass("hide_onet");
			document.getElementById(descId).style.display = "none";
		}
		sessionStorage.setItem("occupationTitle", title);
		//document.getElementById(descId).style.display = "block";
		
	/*} else {
		alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please select an occupation from the list</p>");
		return false;
	}*/
	
	/*if($('#'+showHideId).attr('class') == "hide_onet") {
		$("#" + showHideId).removeClass("hide_onet");
		$("#" + showHideId).addClass("show_onet");
		document.getElementById(descId).style.display = "block";
	} else {
		$("#" + showHideId).removeClass("show_onet");
		$("#" + showHideId).addClass("hide_onet");
		document.getElementById(descId).style.display = "none";
	}*/
	
	
}


function processSelectedOccupation(distinction) {
	if (undefined == $("input[name=occupation]:checked").val()) {
		alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please select an occupation from the list</p>");
	} else {
		if (distinction == "BS") {
			document.getElementById('contollerOccupation').style.display="none";
		} else {
			document.getElementById('AScontollerOccupation').style.display="none";
		}
		var code =$("input[name=occupation]:checked").val();
		goBack();
		if(sessionStorage.getItem("firstResult")){
			sessionStorage.removeItem("firstResult");
			firstRslt = 0;
			sessionStorage.setItem("firstResult", firstRslt);
		}
		if(sessionStorage.getItem("maxResults")){
			sessionStorage.removeItem("maxResults");
			maxRslt = 1;
			sessionStorage.setItem("maxResults", maxRslt);
		}
		if (distinction == "BS") {
			document.getElementById("imageDiv").innerHTML = "<div align='center' id='occupationPlottingDiv'><br /><br /><img src='../img/Processing.gif' /></div>";
		} else {
			document.getElementById("ASimageDiv").innerHTML = "<div align='center' id='ASoccupationPlottingDiv'><br /><br /><img src='../img/Processing.gif' /></div>";
		}
		searchSelected(distinction, code);
	}
}

function searchSelected(distinction, code) {
	// Before Sketch
	if (document.getElementById('contoller')) {
		document.getElementById('contoller').style.display="none";
		document.getElementById('imageDiv').innerHTML = "<br /><br /><br /><br /><br /><br /><div align='center'><img src='../img/Processing.gif'/></div>";
	}
	
	// After Sketch
	if (document.getElementById('AScontoller')) {
		document.getElementById('AScontoller').style.display="none";
		document.getElementById('ASimageDiv').innerHTML = "<br /><br /><br /><br /><br /><br /><div align='center'><img src='../img/Processing.gif'/></div>";
	}
	
	imageText = "";
	
	//Create a model
	var beforeSketch = Spine.Model.sub();
	beforeSketch.configure("/user/search/searchAllDiagramsWithCode", "distinction", "firstResult", "maxResults", "code");
	beforeSketch.extend( Spine.Model.Ajax );
	
	//Populate the model
	var populateModel = new beforeSketch({  
		distinction: distinction,
		firstResult: sessionStorage.getItem("firstResult"),
		maxResults: maxRslt,
		code: code
	});
	populateModel.save(); //POST
	
	beforeSketch.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	
	beforeSketch.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		if (distinction == "BS") {
			plotBeforeSketch (obj);
		} else {
			plotAfterSketch (obj);
		}
	});
}

function searchBeforeSketchAll() {
	document.getElementById("occupationBS").value= "";
	document.getElementById("contollerOccupation").style.display = "none";
	searchBeforeSketch();
}

function searchAfterSketchAll() {
	document.getElementById("occupationAS").value= "";
	document.getElementById("AScontollerOccupation").style.display = "none";
	searchAfterSketch();
}