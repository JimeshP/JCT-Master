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
function submitDecision() {
	var correct = "Correct";
	var change = "Change It";
	if (sessionStorage.getItem("correctInfo")) {
		correct = sessionStorage.getItem("correctInfo");
	}
	if (sessionStorage.getItem("changeInfo")) {
		change = sessionStorage.getItem("changeInfo");
	}
	if(document.getElementById("addWithoutEditsId").checked) {
		alertify.set({ buttonReverse: true });
		alertify.set({ labels: {
		    ok     : correct,
		    cancel : change
		} });
		alertify.confirm("<img src='../img/confirm-icon.png'><br /><p>Is your occupation still the following? <br />"+sessionStorage.getItem("demoInfoMsg")+"</p>", function (e) {
		    if (e) {
		    	addDiagWithoutEdits();
		    } else {
		    	populateEditableData();
		    }
		});
	} else if(document.getElementById("addAfterEditId").checked) {
		alertify.set({ buttonReverse: true });
		alertify.set({ labels: {
		    ok     : correct,
		    cancel : change
		} });
		alertify.confirm("<img src='../img/confirm-icon.png'><br /><p>Is your occupation still the following? <br />"+sessionStorage.getItem("demoInfoMsg")+"</p>", function (e) {
		    if (e) {
		    	addEdits();
		    } else {
		    	populateEditableData();
		    }
		});
	} else if(document.getElementById("keepPreviousDiagId").checked){
		keepPrevious();
	}
}

function keepPrevious() {
	//alert('test');
	var master = Spine.Model.sub();
	master.configure("/user/confirmation/keepPrevious", "jobReferenceNos");
	master.extend( Spine.Model.Ajax );
	
	//Populate the model
	var populateMaster = new master({  
		jobReferenceNos: sessionStorage.getItem("jobReferenceNo")
	});
	populateMaster.save(); //POST
	master.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to fetch before sketch diagram.");
	});
	populateMaster.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		if (obj.statusCode == 200 || obj.statusCode == "200") {
			sessionStorage.setItem("pageSequence", 8);
			window.location = "finalPage.jsp";
		} else {
			alertify.alert("<img src='../img/failed-icon.png'><br /><p>Unable to complete your request. Please contact administrator.</p>");
		}
	});
}

function searchDiags() {
	$("#fancybox-manual-b-final").click(function() {
		$.fancybox.open({
			href : '#sketchSearchDivFinal',
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
	populateBSDiagram();
}

/*function populateEditableData() {
	$.fancybox.open({
		href : '#rectificationId',
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
	
	var master = Spine.Model.sub();
	master.configure("/user/auth/populateUserDetails", "emailId");
	master.extend( Spine.Model.Ajax );
	
	//Populate the model
	var populateMaster = new master({  
		emailId: sessionStorage.getItem("jctEmail")
	});
	populateMaster.save(); //POST
	master.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to your account details.");
	});
	master.bind("ajaxSuccess", function(record, xhr, settings, error) {
		//alertify.alert("Unable to fetch before sketch diagram.");
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		if (obj.statusCode == "200") {
			var htmlString = "<div class='form_area_middle'><div class='form_area_wraper_justified_page'><div id='headerDiv' class='popupHeader'>"+obj.statusDesc+"</div><form class='form-horizontal' name='signupFrm' autocomplete='on' novalidate enctype='multipart/form-data' data-remote='true'>";
			//Region
			htmlString = htmlString + "<div class='single_form_item'><div class='col-md-4 '><label for='inputFN' class='col-md-12 control-label align-right'>Region:</label></div>";
			htmlString = htmlString + "<div class='col-md-5'><select class='form-control-general' id='inputLocation'><option class='form-control-general' id='chooseRegnId' value=''>Select Your Region</option>";
			var regionSelect = document.getElementById("inputLocation");
			for (var key in obj.region) {				
				//Create a option element dynamically
				var option = document.createElement("option");
				if(option.value == obj.accountVO.region){
					htmlString = htmlString + "<option id='"+obj.region[key]+"' selected class='form-control-general' value='"+obj.region[key]+"'>"+obj.region[key]+"</option>";
				} else {
					htmlString = htmlString + "<option id='"+obj.region[key]+"' class='form-control-general' value='"+obj.region[key]+"'>"+obj.region[key]+"</option>";
				}
				option.className = "form-control-general";
			    try {
			    	regionSelect.add(option, null); //Standard 
			    }catch(error) {
			    	//regionSelect.add(option); // IE only
			    }
			}
			
			htmlString = htmlString + "</select></div>";
			htmlString = htmlString + "<div class='col-md-3'>&nbsp;</div><div class='clearfix'></div>";
			
			//function group
			htmlString = htmlString + "<div class='single_form_item'><div class='col-md-4 '><label for='inputFN' class='col-md-12 control-label align-right'>Function Group:</label></div>";
			htmlString = htmlString + "<div class='col-md-5'><select class='form-control-general' id='inputFunctionGroup'>";
			var regionSelect = document.getElementById("inputFunctionGroup");
			for (var key in obj.functionGroup) {				
				//Create a option element dynamically
				var option = document.createElement("option");
				if(option.value == obj.accountVO.functionGroup){
					htmlString = htmlString + "<option id='"+obj.functionGroup[key]+"' selected class='form-control-general' value='"+obj.functionGroup[key]+"'>"+obj.functionGroup[key]+"</option>";
				} else {
					htmlString = htmlString + "<option id='"+obj.functionGroup[key]+"' class='form-control-general' value='"+obj.functionGroup[key]+"'>"+obj.functionGroup[key]+"</option>";
				}
				option.className = "form-control-general";
			    try {
			    	regionSelect.add(option, null); //Standard 
			    }catch(error) {
			    	//regionSelect.add(option); // IE only
			    }
			}
			htmlString = htmlString + "</select></div>";
			htmlString = htmlString + "<div class='col-md-3'>&nbsp;</div><div class='clearfix'></div>";
			
			//level
			htmlString = htmlString + "<div class='single_form_item'><div class='col-md-4 '><label for='inputFN' class='col-md-12 control-label align-right'>Job Level:</label></div>";
			htmlString = htmlString + "<div class='col-md-5'><select class='form-control-general' id='inputJobLevel'>";
			var regionSelect = document.getElementById("inputJobLevel");
			for (var key in obj.jobLevel) {				
				//Create a option element dynamically
				var option = document.createElement("option");
				if(option.value == obj.accountVO.jobLevel){
					htmlString = htmlString + "<option id='"+obj.jobLevel[key]+"' selected class='form-control-general' value='"+obj.jobLevel[key]+"'>"+obj.jobLevel[key]+"</option>";
				} else {
					htmlString = htmlString + "<option id='"+obj.jobLevel[key]+"' class='form-control-general' value='"+obj.jobLevel[key]+"'>"+obj.jobLevel[key]+"</option>";
				}
				option.className = "form-control-general";
			    try {
			    	regionSelect.add(option, null); //Standard 
			    }catch(error) {
			    	//regionSelect.add(option); // IE only
			    }
			}
			htmlString = htmlString + "</select></div>";
			htmlString = htmlString + "<div class='col-md-3'>&nbsp;</div><div class='clearfix'></div>";
			htmlString = htmlString + "</form></div></div>";
			
			htmlString = htmlString + "<div class='single_form_item'>"+
                     "<p><input type='button' class='btn btn-primary btn-sm' value='Update' id='saveDetails' onclick='updateUserDeatils()' /></p>"+
                   "</div>";
				   
			document.getElementById("rectificationId").innerHTML = htmlString;
			if (obj.accountVO.region != "") {
				document.getElementById(obj.accountVO.region).selected=true;
			} else {
				document.getElementById('chooseRegnId').selected=true;
			}
			if (obj.accountVO.functionGroup != "") {
				document.getElementById(obj.accountVO.functionGroup).selected=true;
			}
			if (obj.accountVO.jobLevel != "") {
				document.getElementById(obj.accountVO.jobLevel).selected=true;
			}
			
		} else {
			alertify.alert("<img src='../img/failed-icon.png'><br /><p>Are you sure you want to cancel? <br />"+obj.statusDesc+"</p>");
		}
	});
}*/


function populateEditableData() {
	$.fancybox.open({
		href : '#rectificationId',
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
	
	var master = Spine.Model.sub();
	master.configure("/user/auth/populateUserDetails", "emailId");
	master.extend( Spine.Model.Ajax );
	
	//Populate the model
	var populateMaster = new master({  
		emailId: sessionStorage.getItem("jctEmail")
	});
	populateMaster.save(); //POST
	master.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to your account details.");
	});
	master.bind("ajaxSuccess", function(record, xhr, settings, error) {
		//alertify.alert("Unable to fetch before sketch diagram.");
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		if (obj.statusCode == "200") {
			var htmlString = "<div class='form_area_middle'><div class='form_area_wraper_justified_page'><div id='headerDiv' class='popupHeader'>"+obj.statusDesc+"</div><form class='form-horizontal' name='signupFrm' autocomplete='on' novalidate enctype='multipart/form-data' data-remote='true'>";
			htmlString = htmlString + "<div class='single_form_item'><div class='col-md-3 '><label for='inputFN' class='col-md-12 control-label align-right mandatory_field'>Occupation:</label></div>";						
			htmlString = htmlString + "<div class='col-md-6'><input type='text' class='form-control-general' id='ocptnSrchTxtId' value='"+obj.occupationVal+"'/></div>";
			htmlString = htmlString + "<div class='col-md-3'><input type='button' class='btn btn-primary btn-sm searchClass' value='Search' id='searchOccupation' onclick='searchOccupationSearch()'/></div>";
			htmlString = htmlString + "<div class='clearfix'></div>";
			
			
			
			//htmlString = htmlString + "<div id='occupationDiv' class='onet_scroll'><div id='headerPanel' class='popupHeader custom_header'>Select the option that best describes your occupation. Feel free to try different search words to find a better fit</div>";
			
			htmlString = htmlString + "<div id='occupationDiv' class='onetData'>";
			//htmlString = htmlString + "<div align='center' id='occupationPlottingDiv'><br /><br /><img src='../img/Processing.gif' /></div>";
			htmlString = htmlString + "<div align='center' id='occupationPlottingDiv'><br /></div>";
			htmlString = htmlString + "<div align='center' id='dataPlottingDiv'></div>";
		
			htmlString = htmlString + "</form></div></div>";
			
			htmlString = htmlString + "<div class='single_form_item'>"+
                     "<p><input type='button' class='btn btn-primary btn-sm' value='Update' id='saveDetails' onclick='updateUserDeatils()' /></p>"+
                   "</div>";
				   
			document.getElementById("rectificationId").innerHTML = htmlString;
			
		/*	if (obj.accountVO.region != "") {
				document.getElementById(obj.accountVO.region).selected=true;
			} else {
				document.getElementById('chooseRegnId').selected=true;
			}
			if (obj.accountVO.functionGroup != "") {
				document.getElementById(obj.accountVO.functionGroup).selected=true;
			}
			if (obj.accountVO.jobLevel != "") {
				document.getElementById(obj.accountVO.jobLevel).selected=true;
			}*/
			
			
			// $(".onet_scroll").css({ height: "10px" });
		} else {
			alertify.alert("<img src='../img/failed-icon.png'><br /><p>Are you sure you want to cancel? <br />"+obj.statusDesc+"</p>");
		}
	});
	//document.getElementById(occupationDiv).style.display = "none";
}

function updateUserDeatils() {
	//var location = document.getElementById("inputLocation").options[document.getElementById("inputLocation").selectedIndex].value;
	//var functionGrp = document.getElementById("inputFunctionGroup").options[document.getElementById("inputFunctionGroup").selectedIndex].value;
	//var jobLevel = document.getElementById("inputJobLevel").options[document.getElementById("inputJobLevel").selectedIndex].value;
	//console.log('location: '+location+', functionGrp: '+functionGrp+', jobLevel: '+jobLevel);
	
	if(document.getElementById("ocptnSrchTxtId").value != "" && null != document.getElementById("ocptnSrchTxtId").value) {
		if($("input:radio[name='occupation']").is(":checked")) {
			var occupationData = document.getElementById("ocptnSrchTxtId").value;
			
			//Create a model
				var userProfile = Spine.Model.sub();
				userProfile.configure("/user/auth/updateDemographicalInfo", "occupationData", "emailIds");
				userProfile.extend( Spine.Model.Ajax );
					
				//Populate the model with data to transfer
				var modelPopulation = new userProfile({  
					occupationData: occupationData,		
					emailIds: sessionStorage.getItem("jctEmail")
				});
				modelPopulation.save(); //POST the data
				
				userProfile.bind("ajaxError", function(record, xhr, settings, error) {
					alertify.alert("Unable to connect to the server.");
				});
				
				userProfile.bind("ajaxSuccess", function(record, xhr, settings, error) {
					//alertify.alert("Unable to connect to the server.");
					var jsonStr = JSON.stringify(xhr);
					var obj = jQuery.parseJSON(jsonStr);
					if (obj.statusCode == 200) {
						if (document.getElementById("addWithoutEditsId").checked) {
							addDiagWithoutEdits();
						} else {
							addEdits();
						}
					} else if (obj.statusCode == 500){
						alertify.confirm("<img src='../img/confirm-icon.png'><br /><p>Is your occupation still the following? <br />"+obj.statusDesc+"</p>", function (e) {
						    if (e) {
						    	addDiagWithoutEdits();
						    }
						});
					} else {
						alertify.alert("<img src='../img/failed-icon.png'><br /><p>"+obj.statusDesc+"</p>");
					}
				});
		  } else {
			  alertify.set({ labels: {
				    ok     : "Ok",
				    cancel : "Cancel"
				} });
			  alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please change occupation to update</p>");
			  return false;
		  }
		} else {
			  alertify.set({ labels: {
				    ok     : "Ok",
				    cancel : "Cancel"
				} });
			  alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please select an occupation from the list</p>");
			  return false;
		  }	
}

function populateBSDiagram() {
	var master = Spine.Model.sub();
	master.configure("/user/search/getPrevDiag", "jobReferenceNos", "distinction");
	master.extend( Spine.Model.Ajax );
	
	//Populate the model
	var populateMaster = new master({  
		jobReferenceNos: sessionStorage.getItem("jobReferenceNo"),
		distinction: "BS"
	});
	populateMaster.save(); //POST
	master.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to fetch before sketch diagram.");
	});
	populateMaster.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		document.getElementById('headerDivFinal2').style.display = "none";
		//document.getElementById('bsLinkDiv').style.display = "block";
		//document.getElementById('asLinkDiv').style.display = "none";
		
		
		//document.getElementById('headerDivFinal').innerHTML = "<table width='100%'><tr><td align='left'>Your previous before sketch</td><td align='right'><a href='#' onclick='fetchASDiag()'>After Sketch Diagram</a></td></tr></table>";
		if (obj.selectedDiagram != "NA") {
			document.getElementById('mainPicFinal').innerHTML = "<img src="+obj.selectedDiagram+" class='thumbnailPopup'/>";
			document.getElementById('headerDivFinal').style.display = "block";
		} else {
			document.getElementById('mainPicFinal').innerHTML = "No Diagram to display";
			document.getElementById('headerDivFinal').style.display = "none";
		}
		
	});
}

function fetchASDiag() {
	var master = Spine.Model.sub();
	master.configure("/user/search/getPrevDiag", "jobReferenceNos", "distinction");
	master.extend( Spine.Model.Ajax );
	
	//Populate the model
	var populateMaster = new master({  
		jobReferenceNos: sessionStorage.getItem("jobReferenceNo"),
		distinction: "AS"
	});
	populateMaster.save(); //POST
	master.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to fetch before sketch diagram.");
	});
	populateMaster.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		document.getElementById('headerDivFinal').style.display = "none";
		//document.getElementById('bsLinkDiv').style.display = "none";
		//document.getElementById('asLinkDiv').style.display = "block";
		//document.getElementById('headerDivFinal').innerHTML = "<table width='100%'><tr><td align='left'>Your previous after sketch</td><td align='right'><a href='#' onclick='populateBSDiagram()'>Before Sketch Diagram</a></td></tr></table>";
		if (obj.selectedDiagram != "NA") {
			document.getElementById('mainPicFinal').innerHTML = "<img src="+obj.selectedDiagram+" class='thumbnailPopup'/>";
			document.getElementById('headerDivFinal2').style.display = "block";
		} else {
			document.getElementById('mainPicFinal').innerHTML = "No Diagram to display";
			document.getElementById('headerDivFinal2').style.display = "none";
		}
		
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
		if (sessionStorage.getItem("bsEdit")) {
			sessionStorage.removeItem("bsEdit");
		}
		if((statusCode == 200)){
			if(sessionStorage.getItem("isLogout") == "N"){
				sessionStorage.setItem("pageSequence", 7); // Before sketch edit
				sessionStorage.setItem("isCompleted", obj.isCompleted);
				sessionStorage.setItem("total_json_add_task", obj.bsJson);
				var myJsonObj = JSON.parse(obj.bsJson);
				sessionStorage.setItem("total_count", myJsonObj.length+1);
				sessionStorage.setItem("div_intial_value", obj.initialJson);
				sessionStorage.setItem("bsEdit","resultPageFinal.jsp");
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
					//sessionStorage.clear();
					//window.location = "../signup-page.jsp";
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
		/*if((statusCode == 200)){
			sessionStorage.clear();
			window.location = "../signup-page.jsp";
		} else */
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
	if (!sessionStorage.getItem("infoDispReq")) {
		document.getElementById("infoMsgDispFinal").style.display = "block";
	}
	sessionStorage.setItem("pageSequence", 6);
	if (sessionStorage.getItem('open')) {
		//var evLink = document.createElement('a');
		//evLink.href = generatePDFLink();
		//evLink.target = '_blank';
		//document.body.appendChild(evLink);
		//evLink.click();
		window.open(generatePDFLink(), "", "width=700, height=600");
		sessionStorage.removeItem('open');
	}
	if (sessionStorage.getItem("bsEdit")) {
		sessionStorage.removeItem("bsEdit");
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
		lastActivePage: "/user/view/resultPageFinal.jsp"
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

function searchOccupationSearch() {
	
	$("#occupationDiv").addClass("onet_scroll");
	$("#dataPlottingDiv").addClass("data_plot_div");
	if (document.getElementById("dataPlottingDiv")) {
		document.getElementById("dataPlottingDiv").innerHTML = "";
	}
	var occupationSearchVal = document.getElementById("ocptnSrchTxtId").value.trim();
	if (occupationSearchVal.trim().length == 0) {	
		$("#occupationDiv").removeClass("onet_scroll");
		$("#dataPlottingDiv").removeClass("data_plot_div");
		
		 alertify.set({ labels: {
			    ok     : "Ok",
			    cancel : "Cancel"
			} });
		alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please provide occupation search String");
		//$("#searchOccupation").click(function() {$.fancybox.close({});});
		return false;
	} else {
		//document.getElementById("goDiv").style.display = "none";
		document.getElementById("dataPlottingDiv").innerHTML = "";
		document.getElementById("occupationPlottingDiv").innerHTML = "<br /><br /><img src='../img/Processing.gif' />";			
		/*$("#searchOccupation").click(function() {
			$.fancybox.open({
				href : '#occupationDiv',
				closeClick : false,
				openEffect : 'elastic',
				openSpeed  : 150,
				closeEffect: 'elastic',
				closeSpeed : 150,
				padding    : 5,
				closeClick : true,
				modal      : true,
				showCloseButton : true,
				afterShow : function() {
					$('.fancybox-skin').append('<a title="Close" class="fancybox-item fancybox-close" href="javascript:jQuery.fancybox.close();"></a>');
				}
			});
		});*/
		/*$("#searchOccupation").click(function() {
			document.getElementById(occupationDiv).style.display = "block";
		});*/
	}
	// Search the o* net list
	var stub = Spine.Model.sub();
	stub.configure("/user/auth/searchOccupationList", "searchString");
	stub.extend( Spine.Model.Ajax );
	occupationSearchVal = occupationSearchVal.replace(/,+/g, ' ');
	var val = occupationSearchVal.replace(/ +/g, ',');	
	if (val.indexOf(",") == 0) {
		val = val.substring(1, val.length);
	}
	val = val.replace(/([*()\+\\{}\[\]\|\/?])+/g, '');
	if (val == "") {
		//alertify.alert("Only special character is not valid.");
		var plottingDiv = document.getElementById("dataPlottingDiv");
		plottingDiv.innerHTML = "<div align='center'><br /><br /><img src='../img/not-found.png'><br /><p>The occupation you have entered is not in the system.<br />Please enter search words and select an occupation from the list provided.</p></div>";
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
			plotData(obj.dtoList);
		} else {
			document.getElementById("occupationPlottingDiv").innerHTML = "";
			var plottingDiv = document.getElementById("dataPlottingDiv");
			//plottingDiv.innerHTML = "<div align='center'><br /><br /><img src='../img/not-found.png'><br /><p><b>"+occupationSearchVal+"</b> "+obj.statusMsg+"</p></div>";
			plottingDiv.innerHTML = "<div align='center'><br /><br /><img src='../img/not-found.png'><br /><p>"+obj.statusMsg+"</p></div>";
		}
	});	
}

function plotData(dtoList) {
	document.getElementById("occupationPlottingDiv").innerHTML = "";
	var plottingDiv = document.getElementById("dataPlottingDiv");
	plottingDiv.innerHTML = "";
	var str = "<table width='65%' class='custom_table_onet'><thead><tr class='tab_header'></tr></thead><tbody>";
	
	for (var index = 0; index < dtoList.length; index++) {
		var trColor = "";
		var titleId = "title_" + index;
		var descId = "desc_" + index;
		//str = str + "<tr id='"+titleId+"' bgcolor="+trColor+"><td class='custom_td_onet' ><input type='radio' class='custom_radio_onet' name='occupation' value='"+dtoList[index].title+"' onclick='showDescription(\""+descId+"\", \""+dtoList.length+"\")'>"+dtoList[index].title+"</td></tr><tr id= '"+descId+"' style='display:none;'><td class='onet_desc_custom_cls' >"+dtoList[index].desc+"</td></tr>";
		var showHideId = "show_hide_" + index;
		str = str + "<tr id='"+titleId+"' bgcolor="+trColor+"><td class='custom_td_onet' ><input type='radio' onclick='setOnetDesc()' class='custom_radio_onet' name='occupation' value='"+dtoList[index].title+"'>"+dtoList[index].title+"</td><td id='"+showHideId+"' class='hide_onet' onclick='showDescription(\""+descId+"\", \""+dtoList.length+"\", \""+showHideId+"\")'></td></tr><tr id= '"+descId+"' style='display:none;'><td class='onet_desc_custom_cls' >"+dtoList[index].desc+"</td></tr>";
	}
	str = str + "</tbody></table>";
	plottingDiv.innerHTML = str;
	//document.getElementById("goDiv").style.display = "block";
}

function closeDiv() {
	if (undefined != $("input[name=occupation]:checked").val()) {
		document.getElementById("ocptnSrchTxtId").value = $("input[name=occupation]:checked").val();
		jQuery.fancybox.close();
	} else {
		alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please select an occupation from the list</p>");
		return false;
	}
}

/*function showDescription(descId, length) {
	for (var index = 0; index < length; index++) {
		var allDsecId = "desc_" + index;
		document.getElementById(allDsecId).style.display = "none";
	}
	document.getElementById(descId).style.display = "block";
	
	if (undefined != $("input[name=occupation]:checked").val()) {
		document.getElementById("ocptnSrchTxtId").value = $("input[name=occupation]:checked").val();
		//jQuery.fancybox.close();
	} else {
		alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please select an occupation from the list</p>");
		return false;
	}
	
}*/

function showDescription(descId, length, showHideId) {
	if($('#'+showHideId).attr('class') == "hide_onet") {
		$("#" + showHideId).removeClass("hide_onet");
		$("#" + showHideId).addClass("show_onet");
		document.getElementById(descId).style.display = "block";
	} else {
		$("#" + showHideId).removeClass("show_onet");
		$("#" + showHideId).addClass("hide_onet");
		document.getElementById(descId).style.display = "none";
	}
	
	//if (undefined != $("input[name=occupation]:checked").val()) {
		//document.getElementById("ocptnSrchTxtId").value = $("input[name=occupation]:checked").val();
		//jQuery.fancybox.close();
	//} 
}

function setOnetDesc() {
	if (undefined != $("input[name=occupation]:checked").val()) {
		document.getElementById("ocptnSrchTxtId").value = $("input[name=occupation]:checked").val();
		//jQuery.fancybox.close();
	} else {
		alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please select an occupation from the list</p>");
		return false;
	}
}
function showOldPDF(){
	window.open('showPdf.jsp', '_blank');	
}