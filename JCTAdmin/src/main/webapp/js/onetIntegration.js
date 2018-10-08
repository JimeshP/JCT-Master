/**
 * Function calls when the page is loaded 
 * to populate the ONet Occupation data list
 */
$(document).ready(function() {
if((navigator.userAgent.match(/iPad/i))) {
  	document.getElementById("filename").disabled = true;
  	document.getElementById("onetDataSaveBtnId").disabled = true;
}
	fetchOnetOccupationList();
});

/**
 * Function calls to ONet Occupation list
 * @param null
 */
function fetchOnetOccupationList() {
	var userProf = Spine.Model.sub();
	userProf.configure("/admin/contentconfig/populateOnetOccupationList", "none");
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
			populateOnetOccupationTable(obj.onetOccupationList);			
		} else if(statusCode == 201) {
			
			document.getElementById("onetOcuupationListId").innerHTML = "<br /><br /><br /><div align='center'><img src='../img/no-record.png'><br /><div class='textStyleNoExist'>"+obj.statusDesc+"</div></div>";
		} else if(statusCode == 500) {
			//Show error message
		}
	});
}



/**
 * Function to show ONet Occupation list as a table view
 * @param onetOccupationData
 */
function populateOnetOccupationTable(onetOccupationData) {
	document.getElementById("existing_list_Id").innerHTML = "<table width='100%'><tr><td width='80%'>Existing O*Net Data</td><td >Export To: <a id='generateXL' href='#' onclick='downloadOnetData()'><img src='/admin/img/Excel-icon.png' /></td></tr></table>";
	var tableDiv = document.getElementById("onetOcuupationListId");
	var headingString = "<table width='94%' id='onetTableId' border=1 align='center' bordercolor='#78C0D3' class='tablesorter'><thead class='tab_header'><tr><th width='4%'><b>SL. No.</b></th><th width='10%'><b>Occupation Code</b></th><th><b>Occupation Title</b></th><th><b>Occupation Description</b></th></tr></thead><tbody>";
	var splitRec = onetOccupationData.split("$$$");
	var count = 1;
	for(var index=0; index<splitRec.length-1; index++) {
		var objSplit = splitRec[index].split("~");
		var rowColor = "";
		if(index%2 == 0){
			rowColor = "#D2EAF0";
		} else {
			rowColor = "#F1F1F1";
		}
		headingString = headingString + "<tr align='center' bgcolor='"+rowColor+"' class='row_width'><td>&nbsp; "+count+".</td>";
		headingString = headingString + "<td align='center'>"+objSplit[1]+"</td>";
		headingString = headingString + "<td align='center'>&nbsp;&nbsp; "+objSplit[2]+"</td>";	
		headingString = headingString + "<td align='center'>&nbsp;&nbsp; "+objSplit[3]+"</td>";	
		headingString = headingString + "</tr>";		
		count = count + 1;
	}
	headingString = headingString + "</tbody></table>";
	tableDiv.innerHTML = headingString;
	//new SortableTable(document.getElementById('onetTableId'), 1);
	$("table").tablesorter();
}


/**
 * Function to check the extension of the uploaded file
 * @param null
 */
function validateFile() {
	var parts = document.getElementById('filename').value.split('.');
	if(null == document.getElementById('filename').value || document.getElementById('filename').value == "") {
		alertify.alert('Please upload XLS file.');
	    return false;
	} else if (parts[parts.length-1] != 'xls' && parts[parts.length-1] != 'xlsx') {
		alertify.alert('Please upload XLS / XLSX file only.');
	    return false;
	}
	return true;
}

function downloadOnetData() {
	var link = document.getElementById('generateXL');
	var reportLink = "../../admin/reports/downloadOnetExcel/";
	reportLink = reportLink + "/existing_onet_data.xls";
	link.setAttribute("href", reportLink);
}