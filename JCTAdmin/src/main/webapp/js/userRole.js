var mainHeight = $(".main_contain").height();
/**
 * Function calls when the page is loaded 
 * to populate the function group list
 */
$(document).ready(function() {	
	document.getElementById('addUserRoleText').innerHTML = "Add Function Group";
	fetchFunctionGroupList('F');
	/***************** To Disable the paste functionality ***************/
    /*$("#fuctionGrpId").bind("paste",function(e) {
        e.preventDefault();
    });
    $("#jobLevelId").bind("paste",function(e) {
        e.preventDefault();
    });*/
});

/**
 * Function to fetch all the function group
 * from database while landing to the page
 * @param objval
 */
function fetchFunctionGroupList(objVal) {	
	var userProf = Spine.Model.sub();
	userProf.configure("/admin/contentconfig/populateExistingFunctionGroup", "none");
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
			populateFunctionGrpTable(obj.existingFunctionGrpList, "VIEW");						
		} else if(statusCode == 201) {
			document.getElementById("existingUserRoleListId").innerHTML = "<div align='center'>"+obj.statusDesc+"</div>";
		} else if(statusCode == 500) {
			//Show error message
		}
	});
}


/**
 * Function to fetch all the job level
 * from database while landing to the page
 * @param objval
 */
function fetchJobLevelList(objVal) {	
	var userProf = Spine.Model.sub();
	userProf.configure("/admin/contentconfig/populateExistingJobLevel", "none");
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
				populateJobLevelTable(obj.existingJobLevelList, "VIEW");			
		} else if(statusCode == 201) {
			document.getElementById("existingUserRoleListId").innerHTML = "<div align='center'>"+obj.statusDesc+"</div>";
		} else if(statusCode == 500) {
			//Show error message
		}
	});
}

/**
 * Function to show the function group list
 * on table view
 * @param functionGrpListData
 * @param action
 */
function populateFunctionGrpTable(functionGrpListData, action) {
	document.getElementById("existing_data_list").innerHTML = "Existing Function Groups";
	var tableDiv = document.getElementById("existingUserRoleListId");
	if (null != functionGrpListData) {
		var headingString = "<table width='94%' id='functionGrpTableId' border=1 align='center' bordercolor='#78C0D3' class='tablesorter'><thead class='tab_header'><tr><th width='15%'><b>SL. No.</b></th><th><b>Function Group Name</b></th></tr></thead><tbody>";
		var splitRec = functionGrpListData.split("$$$");
		if(splitRec.length <= 2){
			document.getElementById("changeOrderFGBtnId").disabled = true;
		}
		var count = 1;
		var content = "";
		//for(var index=0; index<splitRec.length-1; index++) {
		for(var index = 0; index < splitRec.length-1; index++){			
			var objSplit = splitRec[index].split("~");	
			var rowColor = "";
			if(index%2 == 0){
				rowColor = "#D2EAF0";
			} else {
				rowColor = "#F1F1F1";
			}
			content = content + "<tr bgcolor='"+rowColor+"' class='row_width'><td align='center'>"+count+".</td><td align='center'>"+objSplit[0]+"</td></tr>";
			count = count + 1;
		}
		headingString = headingString + content + "</tbody></table>";
		tableDiv.innerHTML = headingString;
		//new SortableTable(document.getElementById('functionGrpTableId'), 1);
		$("table").tablesorter(); 
	} else {		
		document.getElementById("changeOrderFGBtnId").disabled = true;
		tableDiv.innerHTML = "<div align='center'>No Function Group To Display</div>";
	}	
}


/**
 * Function to show the job level list
 * on table view
 * @param jobLevelListData
 * @param action
 */
function populateJobLevelTable(jobLevelListData, action) {
	document.getElementById("existing_data_list").innerHTML = "Existing Job Levels";
	var tableDiv = document.getElementById("existingUserRoleListId");
	var headingString = "<table width='94%' id='jobLevelTableId' border=1 align='center' bordercolor='#78C0D3' class='tablesorter' ><thead class='tab_header'><tr><th width='15%'><b>SL. No.</b></th><th><b>Job Level Name</b></th></tr></thead><tbody>";
	var splitRec = jobLevelListData.split("$$$");
	var count = 1;
	var content = "";
	for(var index = 0; index < splitRec.length-1; index++){	
		var objSplit = splitRec[index].split("~");	
		var rowColor = "";
		if(index%2 == 0){
			rowColor = "#D2EAF0";
		} else {
			rowColor = "#F1F1F1";
		}
		content = content + "<tr bgcolor='"+rowColor+"' class='row_width'><td align='center'>"+count+".</td><td align='center'>"+objSplit[0]+"</td></tr>";
		count = count + 1;
	}
	headingString = headingString + content + "</tbody></table>";
	tableDiv.innerHTML = headingString;
	//new SortableTable(document.getElementById('jobLevelTableId'), 1);
	$("table").tablesorter(); 
}


/**
 * Function add to check the  
 * selected radio button
 * @param obj
 */
function checkSelectedOption(obj) {		
	var funcGroupDiv = document.getElementById('funcGroupDiv');
	var jobLevelDiv = document.getElementById('jobLevelDiv');
	var functionGrpAddDiv = document.getElementById('functionGrpAddDiv');
	var jobLevelAddDiv = document.getElementById('jobLevelAddDiv');
	var addUserRoleText = document.getElementById('addUserRoleText');	
	
	if(obj.value == 'F'){		
		funcGroupDiv.setAttribute("style", "display:block");	
		jobLevelDiv.setAttribute("style", "display:none");
		functionGrpAddDiv.setAttribute("style", "display:block");	
		jobLevelAddDiv.setAttribute("style", "display:none");	
		addUserRoleText.innerHTML = "Add Function Group";
		fetchFunctionGroupList(obj);		
	} else {		
		funcGroupDiv.setAttribute("style", "display:none");	
		jobLevelDiv.setAttribute("style", "display:block");
		functionGrpAddDiv.setAttribute("style", "display:none");	
		jobLevelAddDiv.setAttribute("style", "display:block");			
		addUserRoleText.innerHTML = "Add Job Level";
		fetchJobLevelList(obj);	
	}
}


/**
 * Function to add the
 * function group data to the database
 * @param null
 */
function saveFunctionGroup() {
	var functionGrp = document.getElementById("fuctionGrpId").value;
	if(validateTextFG(functionGrp)) {
		var newUserProf = Spine.Model.sub();
		newUserProf.configure("/admin/contentconfig/saveFunctionGroup", "functionGrpVal", "createdBy");
		newUserProf.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new newUserProf({  
			functionGrpVal: functionGrp,
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
				alertify.alert("Function Group has been successfully created.");
				populateFunctionGrpTable(obj.existingFunctionGrpList, "ADD");
				document.getElementById('fuctionGrpId').value = "";
				document.getElementById('fuctionGrpId').innerHTML = "";	
			} else if(statusCode == 600) {
				alertify.alert("Function Group already exists.");
			} else if(statusCode == 500) {
				// Error Message
			}
		});
	}
}

/**
 * Function to add the
 * job level data to the database
 * @param null
 */
function saveJobLevel() {
	var jobLevel = document.getElementById("jobLevelId").value;
	if(validateTextJL(jobLevel)) {
		var newUserProf = Spine.Model.sub();
		newUserProf.configure("/admin/contentconfig/saveJobLevel", "jobLevelVal", "createdBy");
		newUserProf.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new newUserProf({  
			jobLevelVal: jobLevel,			
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
				alertify.alert("Job Level has been successfully created.");
				populateJobLevelTable(obj.existingJobLevelList, "ADD");
				document.getElementById('jobLevelId').value = "";	
				document.getElementById('jobLevelId').innerHTML = "";				
			} else if(statusCode == 600) {
				alertify.alert("Job Level already exists.");
			} else if(statusCode == 500) {
				// Error Message
			}
		});
	}
}

/**
 * Function to validate the
 * enter text for function group and job level
 * @param input
 */
function validateTextFG(input){
	var re=/^\s*$/;

	if (re.test(input)){
		alertify.alert("Please enter function group.");
		return false;
	}
	
/*	if (re.test(functionGrpOrder)){
		alertify.alert("Please enter function group order.");
		return false;
	} else if(functionGrpOrder == 0){
		document.getElementById('fuctionGrpOrdertId').value = "";	
		alertify.alert("Order cannot be 0.");
		return false;
	} else {
		var tbl = document.getElementById("functionGrpTableId");
		var numRows = tbl.rows.length;
		var array = new Array();
		for (var i = 1; i < numRows; i++) {
		    var cells = tbl.rows[i].getElementsByTagName('td');
		    array.push(cells[1].innerHTML);
		    if(array.indexOf(functionGrpOrder) > -1){
		    	//document.getElementById('jobLevelOrderId').value = "";	
		    	alertify.alert("Order already exists.");
				return false;
		    }
		}
	}*/
		
	return true;
}

/**
 * Function to validate the
 * enter text for function group and job level
 * @param input
 */
function validateTextJL(input){ 
	var re=/^\s*$/;
	if (re.test(input)){
		alertify.alert("Please enter job level.");
		return false;
	}			
	return true;
}

/**
 * Function call while click on reset button 
 * to reset the value to null
 * @param null
 */
function resetValue() {
	document.getElementById('fuctionGrpId').innerHTML = "";
	document.getElementById('fuctionGrpId').value = "";
	document.getElementById('jobLevelId').innerHTML = "";	
	document.getElementById('jobLevelId').value = "";	
}

/**
 * Function add to allow only numbers
 * as a input except 0
 * @param event
 */
function numberOnly(event) {	
	 var c = event.which || event.keyCode;
	 console.log(c);
	 if ( (c >= 48  && c <= 57) || c == 43) 
	        return true;
	    return false;
}

/**
 * Function to view all the function group
 * in a pop-up window to change the function group order
 * @param null
 */
function changeOrderFG() {
	document.getElementById('headerDescMainId').innerHTML = "Change Function Group Order";	
	var JLSaveBtn = document.getElementById("JLSaveBtnId");	
	JLSaveBtn.setAttribute("style", "display:none");	
	var FGSaveBtn = document.getElementById("FGSaveBtnId");	
	FGSaveBtn.setAttribute("style", "display:block");	
	populateFunctionGroupList();	
}

/**
 * Function to fetch all the function group
 * from database to a popup window
 * @param null
 */
function populateFunctionGroupList() {	
	var userProf = Spine.Model.sub();
	userProf.configure("/admin/contentconfig/populateFunctionGroupByOrder", "none");
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
			populateFGOrderTable(obj.existingFunctionGrpList);						
		} else if(statusCode == 201) {
			document.getElementById("existingUserRoleListId").innerHTML = "<div align='center'>"+obj.statusDesc+"</div>";
		} else if(statusCode == 500) {
			//Show error message
		}
	});
}

/**
 * Function to populate the function group list
 * in a table view
 * @param listData
 */
function populateFGOrderTable(listData) {		
	var tableDiv = document.getElementById("mainAreaListId");
	if (null != listData) {
		var headingString = "<table width='90%' id='table-1' border='1' align='center' class='tablesorter'><thead align='center' class='tab_header'><tr><th><b>Function Group</b></th></tr></thead><tbody>";
		var splitRec = listData.split("$$$");
		var counter = 1;
		for(var outerIndex = 0; outerIndex < splitRec.length-1; outerIndex++){			
			var objSplit = splitRec[outerIndex].split("~");		
				headingString = headingString + "<tr align='center' id="+counter+" class='custom_table_row' onmousedown='highlight()'><td align='center'>"+objSplit[0]+"</td>";	
				headingString = headingString + "</tr>";
				counter = counter + 1;							
		}		
		headingString = headingString + "</tbody></table>";
		tableDiv.innerHTML = headingString;
		
	} 
	 $("#table-1").tableDnD();
	 $("table").tablesorter(); 
}

function highlight() {
	$("tr").mousedown(function(){
		$(this).addClass("selected").siblings().removeClass("selected");
	});
}

/**
 * Function to save the function group
 * by function group order
 * @param null
 */
function saveFuncGrpOrder() {
	var builder = "";
	$('#table-1 tbody td').each(function() {
	    var cellText = $(this).html();  
	    builder = builder + cellText +"~~";	
	});
	var newUserProf = Spine.Model.sub();
	newUserProf.configure("/admin/contentconfig/saveFuncGrpOrder", "builder");
	newUserProf.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new newUserProf({  
		builder: builder		
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
			alertify.alert("Function Group Order has been changed successfully.");
			populateFunctionGroupList();
			fetchFunctionGroupList('F');
		}  else if(statusCode == 500) {
			// Error Message
		}
	});
}

/**
 * Function to view all the job level
 * in a pop-up window to change the job level order
 * @param null
 */
function changeOrderJL() {
	document.getElementById('headerDescMainId').innerHTML = "Change Job Level Order";		
	var JLSaveBtn = document.getElementById("JLSaveBtnId");	
	JLSaveBtn.setAttribute("style", "display:block");	
	var FGSaveBtn = document.getElementById("FGSaveBtnId");	
	FGSaveBtn.setAttribute("style", "display:none");	
	populateJobLevelList();
}

/**
 * Function to fetch all the job level
 * from database to a popup window
 * @param null
 */
function populateJobLevelList() {	
	var userProf = Spine.Model.sub();
	userProf.configure("/admin/contentconfig/populateJobLevelByOrder", "none");
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
			populateJLOrderTable(obj.existingJobLevelList);						
		} else if(statusCode == 201) {
			document.getElementById("existingUserRoleListId").innerHTML = "<div align='center'>"+obj.statusDesc+"</div>";
		} else if(statusCode == 500) {
			//Show error message
		}
	});
}
/**
 * Function to populate the job level list
 * in a table view
 * @param listData
 */
function populateJLOrderTable(listData) {	
	var tableDiv = document.getElementById("mainAreaListId");
	if (null != listData) {
		var headingString = "<table width='90%' id='table-1' border='1' align='center' class='tablesorter'><thead align='center' class='tab_header'><tr><th><b>Job Level</b></th></tr></thead><tbody>";
		var splitRec = listData.split("$$$");
		var counter = 1;
		for(var outerIndex = 0; outerIndex < splitRec.length-1; outerIndex++){			
			var objSplit = splitRec[outerIndex].split("~");		
				headingString = headingString + "<tr align='center' id="+counter+" class='custom_table_row' onmousedown='highlight()'><td align='center'>"+objSplit[0]+"</td>";	
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
 * Function to save the job level
 * by job level order
 * @param null
 */
function saveJobLevelOrder() {
	var builder = "";
	$('#table-1 tbody td').each(function() {
	    var cellText = $(this).html();  
	    builder = builder + cellText +"~~";	
	});
	var newUserProf = Spine.Model.sub();
	newUserProf.configure("/admin/contentconfig/saveJobLevelOrder", "builder");
	newUserProf.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new newUserProf({  
		builder: builder		
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
			alertify.alert("Job Level Order has been successfully changed.");
			populateJobLevelList();
			fetchJobLevelList('J');
		}  else if(statusCode == 500) {
			// Error Message
		}
	});
}