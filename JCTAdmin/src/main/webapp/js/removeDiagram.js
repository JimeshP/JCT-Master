$(document).ready(function() {
	fetchAllDiagramCount();
	fetchAllDiagramOnLoad();	
});


/**
 * Function fetches before sketch report
 * @param null
 */
function fetchAllDiagramCount(){
	var bsModel = Spine.Model.sub();
	bsModel.configure("/admin/adminRemoveDiagram/fetchAllDiagramCount", "none");
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
		if(obj.statusCode = 200) {
			sessionStorage.setItem("totalDiagramCount",obj.totalDiagramCount);
		}
	});
}

/**
 * Function searches the diagram which the users have shared 
 * in public database
 */
function searchUserDiagram() {
	
	document.getElementById("totalDiagramCountId").style.display = "none";
	var emailId = document.getElementById("userName").value.trim();
	if (validateEmailId(emailId)) {
		document.getElementById("diagramPlottingDiv").style.display = "block";
		//document.getElementById("diagramPlottingDiv").style.display = "block";
		var plottingDiv = document.getElementById("existingUserDiagramId");
		plottingDiv.innerHTML = "<div align='center' width='100%'><img src='../img/dataLoading.GIF'><br />Searching Diagram</div>";
		var userDiag = Spine.Model.sub();
		userDiag.configure("/admin/adminRemoveDiagram/searchUserDiagram", "emailId");
		userDiag.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new userDiag({
			emailId: emailId
		});
		modelPopulator.save(); //POST
		userDiag.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
		});
		userDiag.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			plotDiagrams(obj.diagramList, obj.statusCode);
		});
	}
}

function validateEmailId(emailId) {
	var emailReg = /^\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$/;
	if (!emailReg.test(emailId)){
		alertify.alert("Please enter valid E-mail ID.");
		return false;
	}
	return true;
}

function plotDiagrams(list, statusCode) {
	document.getElementById("diagramPlottingDiv").style.display = "block";
	var plottingDiv = document.getElementById("existingUserDiagramId");
	if (statusCode == 200) {
		document.getElementById("existing_Diagram_Id").style.display="block";
		var plottingString = "<div class='single_form_item' style='padding:1%'>";	
		for (var index = 0; index < list.length; index++) {
			// for BS diagram
			plottingString = plottingString + "<div class='col-sm-6 imageHolder imageHolder_ipad' title='Before Sketch Diagram' onclick='showLargeImage("+list[index].beforeSketchBaseString+")'><img id='dispId' src="+list[index].beforeSketchBaseString+" height='100%' width='100%'/></div>";
			// for AS Diagram
			plottingString = plottingString + "<div class='col-sm-6 imageHolder imageHolder_ipad' title='After Sketch Diagram' onclick='showLargeImageAS("+list[index].aftereSketchBaseString+")'><img id='dispId2' src="+list[index].aftereSketchBaseString+" height='100%' width='100%'/></div><div class='clearfix'></div>";
			//plottingString = plottingString + "<p class='' id='removeDiagmUserNameId_"+list[index].rowId+"'>"+list[index].userName+"</p>";	
			plottingString = plottingString + "<div class='clearfix'></div></div>";
			
			// for remove button
			plottingString = plottingString + "<div class='single_form_item'><p><input type='button' class='btn_admin btn_admin-primary btn_admin-sm search_btn' value='Remove' onclick='deleteDiagram("+list[index].rowId+")' /></p></div></div>";
		}	
		plottingDiv.innerHTML = plottingString;
	} else if (statusCode == 201) {
		document.getElementById("existing_Diagram_Id").style.display="none";
		plottingDiv.innerHTML = "<div align='center'><img src='../img/no-record.png'><br /><div class='textStyleNoExist' align='center'>User has not shared any diagram.</div></div>";
	} else {
		plottingDiv.innerHTML = "<div align='center'>Server encountered error.</div>";
	}
}

	
function deleteDiagram(rowId) {
	
	alertify.set({ buttonReverse: true });
		alertify.set({ labels: {
		    ok     : "Yes",
		    cancel : "No"
		} });
	alertify.confirm("<p>Are you sure to delete the following before sketch and after diagram from searchable database?</p>", function (e) {
	    if (e) {
	    	// soft delete the row with 2
	    	var userDiag = Spine.Model.sub();
			userDiag.configure("/admin/adminRemoveDiagram/removeDiagram", "emailId", "rowId", "rowIdSet");
			userDiag.extend( Spine.Model.Ajax );
			//Populate the model with data to transfer
			var modelPopulator = new userDiag({  
				emailId: sessionStorage.getItem("jctEmail"),
				rowId: rowId,
				rowIdSet: ""
			});
			modelPopulator.save(); //POST
			userDiag.bind("ajaxError", function(record, xhr, settings, error) {
				alertify.alert("Unable to connect to the server.");
			});
			userDiag.bind("ajaxSuccess", function(record, xhr, settings, error) {
				var jsonStr = JSON.stringify(xhr);
				var obj = jQuery.parseJSON(jsonStr);
				if (obj.statusCode == 200) {
					alertify.alert(obj.statusMessage);
					document.getElementById("existingUserDiagramId").innerHTML = "<br /><br /><div align='center'><img src='../img/right.gif' /><br />"+obj.statusMessage+"</div>";
				} else {
					alertify.alert(obj.statusMessage);
				}
			});
		}
	});
}
function showLargeImage(imageString) {
	if (document.getElementById("diagramDiv") == null) {
		var div = document.createElement('div');
		div.id = 'diagramDiv';
		document.body.appendChild(div);
	}
	$.fancybox.open({
			href : '#diagramDiv',
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
				$('.fancybox-skin').append('<a title="Close" class="fancybox-item fancybox-close" onclick="simulateClose();"></a>');
			}
	});
	if (document.getElementById("diagramDiv") == null) {
		var div = document.createElement('div');
		div.id = 'diagramDiv';
		document.body.appendChild(div);
		div.innerHTML = "<img src="+imageString+">";
	} else {
		document.getElementById("diagramDiv").innerHTML = "<img src="+imageString+">";
	}
}

function simulateClose() {
if (document.getElementById("diagramDiv") == null) {
		var div = document.createElement('div');
		div.id = 'diagramDiv';
		document.body.appendChild(div);
}
if (document.getElementById("diagramDiv2") == null) {
		var div = document.createElement('div');
		div.id = 'diagramDiv2';
		document.body.appendChild(div);
}
jQuery.fancybox.close();
}


function showLargeImageAS(imageString) {
	if (document.getElementById("diagramDiv2") == null) {
		var div = document.createElement('div');
		div.id = 'diagramDiv2';
		document.body.appendChild(div);
	}
	$.fancybox.open({
			href : '#diagramDiv2',
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
				$('.fancybox-skin').append('<a title="Close" class="fancybox-item fancybox-close" onclick="simulateClose();"></a>');
			}
	});
	if (document.getElementById("diagramDiv2") == null) {
		var div = document.createElement('div');
		div.id = 'diagramDiv2';
		document.body.appendChild(div);
		div.innerHTML = "<img src="+imageString+">";
	} else {
		document.getElementById("diagramDiv2").innerHTML = "<img src="+imageString+">";
	}
}


function fetchAllDiagramOnLoad() {
	document.getElementById("diagramPlottingDiv").style.display = "block";
	document.getElementById("diagramPlottingDiv").value = "<img src='../img/dataLoading.GIF'>";
	fetchAllDiagramCount();
	sessionStorage.setItem("checkedValue",0);
	firstResult=0;
	sessionStorage.setItem("firstResult",firstResult);
	fetchAllDiagram();
}

/**
 * Function fetches before sketch report
 * @param null
 */
function fetchAllDiagram(){
	var bsModel = Spine.Model.sub();
	bsModel.configure("/admin/adminRemoveDiagram/populateAllDiagram", "recordIndex");
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
		populateAllDiagram(obj.diagramList, obj.statusCode);
	});
}


function populateAllDiagram(list, statusCode) {
	document.getElementById("diagramPlottingDiv").style.display = "block";
	var plottingDiv = document.getElementById("existingUserDiagramId");		
	var totalCountDiv = document.getElementById("totalDiagramCountId");
	if (statusCode == 200) {
		document.getElementById("existing_Diagram_Id").style.display="block";
		totalCountDiv.innerHTML = "Displaying "+list.length+ " Out of " +sessionStorage.getItem("totalDiagramCount")+ " Diagrams.";
		var plottingString = "";	
		
		plottingString = plottingString + "<div class='remove_diagram_outer_box'>";
		for (var index = 0; index < list.length; index++) {
			createdTs = new Date((list[index].jctCreatedTs)).toDateString();
			formattedDate = dateformat(new Date (createdTs));
			plottingString = plottingString + "<div class='single_form_item remove_diagram_box'>";
			// for Checkbox
			plottingString = plottingString + "<div class='col-sm-1 remove_diagram_chkBox'><input type='checkbox' id='removeDiagmChkBoxId_"+list[index].rowId+"' value="+list[index].rowId+" onclick='getCheckBoxValue(this)'></div>";
			plottingString = plottingString + "<div class='col-sm-5 imageHolderSmall imageHolder_ipad' title='Before Sketch Diagram' onclick='showLargeImage("+list[index].beforeSketchBaseString+")'><img id='dispId' src="+list[index].beforeSketchBaseString+" height='100%' width='100%'/></div>";
			// for AS Diagram
			plottingString = plottingString + "<div class='col-sm-5 imageHolderSmall imageHolder_ipad' title='After Sketch Diagram' onclick='showLargeImageAS("+list[index].aftereSketchBaseString+")'><img id='dispId2' src="+list[index].aftereSketchBaseString+" height='100%' width='100%'/></div>";
			//for User Name
			plottingString = plottingString + "<p id='removeDiagmUserNameId_"+list[index].rowId+"'><b>"+list[index].userName+"<br>"+formattedDate+"</p>";				
			plottingString = plottingString + "</div>";			
		}	
		plottingString = plottingString + "</div>";
		// for remove button
		plottingString = plottingString + "<div class='single_form_item'><p><input type='button' class='next_prev_btn btn_admin btn_admin-primary btn_admin-sm search_btn prev_btn_remove_diagram' id='prevBtnId' value='PREV' onclick='backToPrevious()' /><input type='button' class='next_prev_btn btn_admin btn_admin-primary btn_admin-sm search_btn' value='Remove' onclick='deleteAllCheckedDiagram()' /><input type='button' id='nextBtnId' class='btn_admin btn_admin-primary btn_admin-sm search_btn next_btn_remove_diagram' value='NEXT' onclick='goToNext()' /></p></div></div>";
		plottingDiv.innerHTML = plottingString;
		if(sessionStorage.getItem("firstResult") == 0){
			$('#prevBtnId').attr('disabled','disabled');
		} 
		if ((parseInt(sessionStorage.getItem("firstResult")) + parseInt(list.length)) == (parseInt(sessionStorage.getItem("totalDiagramCount")))) {
			$('#nextBtnId').attr('disabled','disabled');
		}
		setCheckBoxValue();
	} else if (statusCode == 201) {
		document.getElementById("existing_Diagram_Id").style.display="none";
		plottingDiv.innerHTML = "<div align='center'><img src='../img/no-record.png'><br /><div class='textStyleNoExist' align='center'>User has not shared any diagram.</div></div>";
	} else {
		plottingDiv.innerHTML = "<div align='center'>Server encountered error.</div>";
	}
}

/**
 * This function gives the date in format
 * @param inputdate
 * @returns formatedDate
 */
function dateformat(inputdate) {
	var dateStr = inputdate.toDateString();
	var len=dateStr.length;
	var formatedDate = dateStr.substr(4,len);
	var newchar = '-';
	formatedDate = formatedDate.split(' ').join(newchar);
	return formatedDate;
}

function goToNext() {
	var indSecond = 0;
	if(sessionStorage.getItem("firstResult") == 0){
		indSecond = 4;
		sessionStorage.setItem("firstResult",indSecond);
	} else {
		indSecond = parseInt(sessionStorage.getItem("firstResult")) + 4;
		sessionStorage.setItem("firstResult",indSecond);
	}
	fetchAllDiagram();
}

function backToPrevious() {	
	var indSecond = 0;
	if(sessionStorage.getItem("firstResult") == 0){
		indSecond = 4;
		sessionStorage.setItem("firstResult",indSecond);
	} else {
		indSecond = parseInt(sessionStorage.getItem("firstResult")) - 4;
		sessionStorage.setItem("firstResult",indSecond);
	}
	fetchAllDiagram();
}

function getCheckBoxValue(obj) {
	var rowId = obj.value;	
	var checkedValue = "";
	if(obj.checked==true){	
		if(sessionStorage.getItem("checkedValue")) {
			checkedValue = sessionStorage.getItem("checkedValue") +"``"+ rowId;
		} 
	} else {
		checkedValue = sessionStorage.getItem("checkedValue").replace("``"+rowId,"");
	}
	sessionStorage.setItem("checkedValue",checkedValue);
}


function setCheckBoxValue() {
	if(sessionStorage.getItem("checkedValue")) {
	var checkedValue = sessionStorage.getItem("checkedValue");
	var rowId = checkedValue.split("``");
	for(var outerIndex = 0; outerIndex < rowId.length; outerIndex++){			
		if(document.getElementById("removeDiagmChkBoxId_"+rowId[outerIndex])) {
			document.getElementById("removeDiagmChkBoxId_"+rowId[outerIndex]).checked = true;
		}
	}
	
	}
}


function deleteAllCheckedDiagram() {
	var checkedValue = sessionStorage.getItem("checkedValue");
	var rowId = checkedValue.split("``");
	if(rowId.length > 1) {		
		alertify.set({ buttonReverse: true });
			alertify.set({ labels: {
			    ok     : "Yes",
			    cancel : "No"
			} });
		alertify.confirm("<p>Are you sure to delete the following before sketch and after diagram from searchable database?</p>", function (e) {
		    if (e) {
		    	// soft delete the row with 2
		    	var userDiag = Spine.Model.sub();
				userDiag.configure("/admin/adminRemoveDiagram/removeDiagram", "emailId", "rowId", "rowIdSet");
				userDiag.extend( Spine.Model.Ajax );
				//Populate the model with data to transfer
				var modelPopulator = new userDiag({  
					emailId: sessionStorage.getItem("jctEmail"),
					rowId: 0,
					rowIdSet: sessionStorage.getItem("checkedValue")
				});
				modelPopulator.save(); //POST
				userDiag.bind("ajaxError", function(record, xhr, settings, error) {
					alertify.alert("Unable to connect to the server.");
				});
				userDiag.bind("ajaxSuccess", function(record, xhr, settings, error) {
					var jsonStr = JSON.stringify(xhr);
					var obj = jQuery.parseJSON(jsonStr);
					if (obj.statusCode == 200) {
						alertify.alert(obj.statusMessage);
						//fetchAllDiagram();
						fetchAllDiagramOnLoad();
					} else {
						alertify.alert(obj.statusMessage);
					}
				});
			}
		});
	} else {
		alertify.alert("Please select one diagram to delete.");
	}	
}