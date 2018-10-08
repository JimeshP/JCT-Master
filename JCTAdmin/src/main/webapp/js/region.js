/**
 * Function calls when the page is loaded 
 * to populate the region list
 */
$(document).ready(function() {
	document.getElementById('addHeaderDiv').innerHTML = "Add Region Name";
	fetchRegionList("VIEW");
	/***************** To Disable the paste functionality ***************/
   /* $("#regionInptId").bind("paste",function(e) {
        e.preventDefault();
    });*/
});

/**
 * Function calls to populate region list
 * @param null
 */
function fetchRegionList(action) {
	var userProf = Spine.Model.sub();
	userProf.configure("/admin/contentconfig/populateExistingRegion", "none");
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
			if( action == "VIEW") {
				populateRegionTable(obj.existingRegionList, "VIEW");
			} else {
				populateRegionOrder(obj.existingRegionList);
			}
			
		} else if(statusCode == 201) {
			document.getElementById("existingRegionListId").innerHTML = "<div align='center'><img src='../img/no-record.png'><br />"+obj.statusDesc+"</div>";
		} else if(statusCode == 500) {
			//Show error message
		}
	});
}

/**
 * Function to save the region data
 * @param null
 */
function saveRegion() {
	var regionName = document.getElementById("regionInptId").value;
	if(validateRegion(regionName)) {
		var newUserProf = Spine.Model.sub();
		newUserProf.configure("/admin/contentconfig/saveRegion", "regionNameVal", "createdBy");
		newUserProf.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new newUserProf({  
			regionNameVal: regionName,
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
				//alertify.set({ delay: 2700 });
				alertify.alert("Region has been successfully created.");
				populateRegionTable(obj.existingRegionList, "ADD");
				document.getElementById('regionInptId').value = "";
			} else if(statusCode == 600) {
				alertify.alert("Region already exists.");
			} else if(statusCode == 500) {
				// Error Message
			}
		});
	}
}

/**
 * Function to validate the text for region name
 * @param regionName
 */
function validateRegion(regionName){
	var re=/^\s*$/;
	if (re.test(regionName)){
		alertify.alert("Please enter Region.");
		return false;
	}
	return true;
}

/**
 * Function to show region list as a table view
 * @param regionListData
 * @param action
 */
function populateRegionTable(regionListData, action) {
	var tableDiv = document.getElementById("existingRegionListId");
	var headingString = "<table width='94%' id='regionTableId' border=1 align='center' bordercolor='#78C0D3' class='tablesorter'><thead class='tab_header'><tr><th width='10%'><b>SL. No.</b></th><th><b>Region Name</b></th><th width='12%'><b>Action</b></th></tr></thead><tbody>";
	var splitRec = regionListData.split("$$$");
	var count = 1;
	var content = "";
	for(var index=0; index<splitRec.length-1; index++) {
		var objSplit = splitRec[index].split("~");
		var rowColor = "";
		if(index%2 == 0){
			rowColor = "#D2EAF0";
		} else {
			rowColor = "#F1F1F1";
		}
		content = content + "<tr bgcolor='"+rowColor+"' class='row_width'><td align='center'>"+count+".</td><td align='center'>"+objSplit[1]+"</td>";
		//content = content + "<td class='table_col_txt_style'>&nbsp;&nbsp;&nbsp;&nbsp;<a class='delete_style' href='#' onClick='deleteRegion(\""+objSplit[0]+"\", \""+objSplit[1]+"\")'><img src = '../img/delete.png' /></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a class='edit_style' href='#' onClick='editRegion(\""+objSplit[0]+"\", \""+objSplit[1]+"\")'><img src = '../img/edit.png' /></a></td>";
		content = content + "<td class='table_col_txt_style' align='center'><a class='edit_style' href='#' onClick='editRegion(\""+objSplit[0]+"\", \""+objSplit[1]+"\")'><img src = '../img/edit.png' /></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a class='delete_style' href='#' onClick='deleteRegion(\""+objSplit[0]+"\", \""+objSplit[1]+"\")'><img src = '../img/delete.png' /></a></td>";
		content = content + "</tr>";
		count = count + 1;
	}
	headingString = headingString + content + "</tbody></table>";
	tableDiv.innerHTML = headingString;
	//new SortableTable(document.getElementById('regionTableId'), 1);
	$("table").tablesorter();
}

/**
 * Function to edit the
 * region data
 * @param tablePkId
 * @param regionDesc
 */
function editRegion(tablePkId, regionDesc){
	document.getElementById('regionInptId').value = regionDesc;
	document.getElementById("updateId").innerHTML = "Update Region";
	document.getElementById('addHeaderDiv').innerHTML = "Update Region Name";
	document.getElementById("addRegionBtnId").value = "Update";
	document.getElementById("addRegionBtnId").setAttribute("onclick", "updateRegion("+tablePkId+",'"+regionDesc+"')");
	var updateSection = document.getElementById("RegionEditDiv");
	updateSection.className = updateSection.className + " otherclass";
	
	document.getElementById("chgOrderBtnId").disabled = true;
}

/**
 * Function to update the
 * region data
 * @param tablePkId
 * @param prevVal
 */
function updateRegion(tablePkId, prevVal) {
	var regionDesc = document.getElementById("regionInptId").value;
	if(validateRegion(regionDesc)) {
		var newUserProf = Spine.Model.sub();
		newUserProf.configure("/admin/contentconfig/updateRegion", "regionDescVal", "tablePkId");
		newUserProf.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new newUserProf({  
			regionDescVal: regionDesc,
			tablePkId: tablePkId
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
				alertify.alert("Region has been edited successfully.");
				populateRegionTable(obj.existingRegionList, "UPDATE");
				
				document.getElementById('regionInptId').value = "";				
				document.getElementById("updateId").innerHTML = "Region";
				document.getElementById('addHeaderDiv').innerHTML = "Add Region Name";
				document.getElementById("addRegionBtnId").value = "Add";
				document.getElementById("addRegionBtnId").setAttribute("onclick", "saveRegion()");
				var updateSection = document.getElementById("RegionEditDiv");
				updateSection.className = updateSection.className - " otherclass";
				
				document.getElementById("chgOrderBtnId").disabled = false;
				
			} else if(statusCode == 600) {
				alertify.alert("Region already exists.");
			} else if(statusCode == 500) {
				populateRegionTable(obj.existingRegionList, "UPDATE");
				alertify.alert(obj.statusDesc);
			}
		});
	}
}

/**
 * Function to delete the
 * region data
 * @param tablePkId
 * @param regionDesc
 */
function deleteRegion(tablePkId, regionDesc){
	alertify.set({ buttonReverse: true });
	alertify.confirm("Are you sure you want to delete the region?", function (e) {
		if (e) {	    	
	    	var newUserProf = Spine.Model.sub();
	    	newUserProf.configure("/admin/contentconfig/deleteRegion", "tablePkId", "regionDesc");
	    	newUserProf.extend( Spine.Model.Ajax );
	    	//Populate the model with data to transfer
	    	var modelPopulator = new newUserProf({  
	    		tablePkId: tablePkId,
	    		regionDesc: regionDesc
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
					alertify.alert("Region has been deleted successfully.");
					populateRegionTable(obj.existingRegionList, "DELETE");
	    		} else if(statusCode == 500) {
	    			// Error Message
	    		}
	    	});
	    
		}
	});
		
}


/**
 * Function call while click on reset button 
 * to reset the value to null
 * @param null
 */
function resetValue() {
	document.getElementById('regionInptId').value = "";
	var updateSection = document.getElementById("RegionEditDiv");
	updateSection.className = updateSection.className - " otherclass";
	document.getElementById("addRegionBtnId").value = "Add";
	document.getElementById("updateId").innerHTML = "Region";
	document.getElementById('addHeaderDiv').innerHTML = "Add Region Name";
	document.getElementById("addRegionBtnId").setAttribute("onclick", "saveRegion()");
	document.getElementById("chgOrderBtnId").disabled = false;
	
	
}

/**
 * Function to view all the region
 * in a pop-up window to change the region order
 * @param null
 */
function changeOrder() {
	document.getElementById('headerDescMainId').innerHTML = "Change Region Order";	
	fetchRegionList("ORDER");
	
}

/**
 * Function to populate the region list
 * in a table view
 * @param listData
 */
function populateRegionOrder(listData) {	
	var tableDiv = document.getElementById("mainAreaListId");
	if (null != listData) {
		var headingString = "<table width='90%' id='table-1' border='1' align='center' class='tablesorter'><thead align='center' class='tab_header'><tr><th><b>Region</b></th></tr></thead><tbody>";
		var splitRec = listData.split("$$$");
		var counter = 1;
		for(var outerIndex = 0; outerIndex < splitRec.length-1; outerIndex++){			
			var objSplit = splitRec[outerIndex].split("~");		
				headingString = headingString + "<tr align='center' id="+counter+" class='custom_table_row' onmousedown='highlight()'><td align='center'>"+objSplit[1]+"</td>";	
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
 * Function to save the region
 * by region order
 * @param null
 */
function saveRegionOrder() {
	var builder = "";
	$('#table-1 tbody td').each(function() {
	    var cellText = $(this).html();  
	    builder = builder + cellText +"~~";	
	});
	var newUserProf = Spine.Model.sub();
	newUserProf.configure("/admin/contentconfig/saveRegionOrder", "builder");
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
			alertify.alert("Region Order has been changed successfully.");
			fetchRegionList("ORDER");
			fetchRegionList("VIEW");
		}  else if(statusCode == 500) {
			// Error Message
		}
	});
}