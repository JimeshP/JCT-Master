var tableContent = "";

$(document).ready(function() {
	populateHeaderReport();
});

/**
 * Function populates the overview report
 */
function populateHeaderReport(){
	var obj = jQuery.parseJSON(sessionStorage.getItem("jsonStr"));
	tableContent = "";
	var hdrDiv = document.getElementById('mainHeaderDivCont');
	var tableDiv = document.getElementById("miscHeaderContentDiv");
	var headerDispTable = obj.tableHeaderString.split("&&&");
	var headingString = "<table width='100%' border='1' bordercolor='#78C0D3' id='myTable'><thead class='tab_header'><tr><th>"+headerDispTable[0]+"</th><th>"+headerDispTable[6]+"</th><th>"+headerDispTable[1]+"</th><th>"+headerDispTable[2]+"</th><th>"+headerDispTable[3]+"</th><td align='center'><b>"+headerDispTable[4]+"</b></td><td align='center'><b>"+headerDispTable[5]+"</b></td></tr></thead><tbody>";
	var dispTable = obj.reportList.split("$$$");
	var size = parseInt(dispTable.length-1);
	hdrDiv.innerHTML = "<table width='100%'><tr><td><b>Total Records Fetched: </b>"+size+"</td><td align='right'><b>Export To: </b><a id='generateXL' href='#' onclick='generateMiscExcel()'><img src='../img/Excel-icon.png' /></td></tr></table>";
	for(var index=0; index<dispTable.length-1; index++){
		var content = "<tr bgcolor='#D2EAF0'>";
		var headerNonRepetingData = dispTable[index].split("#");
		content = content + "<td align='center'>"+parseInt(index+1)+"</td>";
		content = content + "<td align='center'>"+headerNonRepetingData[5]+"</td>";
		content = content + "<td align='center'>"+headerNonRepetingData[0]+"</td>";
		content = content + "<td align='center'>"+headerNonRepetingData[2]+"</td>";
		content = content + "<td align='center'>"+headerNonRepetingData[3]+"</td>";
		content = content + "<td align='center'> <a href='#' id='myDoc' data-transition='pop' onclick=expandView(\""+headerNonRepetingData[1]+"\",'"+headerNonRepetingData[0]+"')><u>"+headerNonRepetingData[4]+"</u></a></td>";
		content = content + "<td align='center'>"+headerNonRepetingData[6]+"</td>";
		content = content + "</tr>";
		tableContent = tableContent + content;
	}
	if(dispTable.length-1 > 0){
		document.getElementById("miscHeaderContentDiv").style.display = "block";
		document.getElementById("mainHeaderDivCont").style.display = "block";
		document.getElementById("noDataDiv").style.display = "none";
	}else{
		//tableDiv.innerHTML = "<div align='center'> <br /><br /><br /><img src='../img/no-record.png'><br />No records</div>";
		document.getElementById("noDataDiv").style.display = "block";
		document.getElementById("miscHeaderContentDiv").style.display = "none";
		document.getElementById("mainHeaderDivCont").style.display = "none";
	}
	var mainContent = headingString + tableContent+"</tbody></table>";
	tableDiv.innerHTML = mainContent;
	new SortableTable(document.getElementById('myTable'), 1);
}
/**
 * Function fetches the detailed report for the job ref nos and email id selected
 * @param jrNo
 * @param emailId
 */
function expandView(jrNo, emailId){
	document.getElementById('mainHeaderDivCont').style.display="none";
	document.getElementById('miscHeaderContentDiv').style.display="none";
	document.getElementById('detailedDiv').style.display="block";
	var bsModel = Spine.Model.sub();
	bsModel.configure("/admin/reports/populateMiscDetailedReport", "jrNos");
	bsModel.extend( Spine.Model.Ajax );
	//Populate the model
	var populateModel = new bsModel({  
		jrNos: jrNo
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
		if(obj.statusCode == 200){
			populateDetailedView(jsonStr, emailId);
		}
	});
}
/**
 * Function populates the detailed div
 * @param jsonStr
 * @param emailId
 */
function populateDetailedView(jsonStr, emailId){
	var tableContents = "";
	var tableDtlHdr = document.getElementById("detailedHdr");
	tableDtlHdr.innerHTML = "<table width='100%'><tr><td><b>Displaying records: </b>"+emailId+"</td><td align='right'><input type='button' class='go_back_btn' value='Go Back' onclick='goBack()' /></td></tr></table>";
	var tableDiv = document.getElementById("detailedDivContent");
	var headingString = "<table border='1' bordercolor='#78C0D3' id='myTable2' class='table'><thead class='tab_header'><tr><th>Sl No.</th><th>Start date/time</th><th>End date/time</th></tr></thead><tbody>";
	var obj = jQuery.parseJSON(jsonStr);
	var dispTable = obj.reportList.split("$$$");
	
	for(var index=0; index<dispTable.length-1; index++){
		var content = "<tr bgcolor='#D2EAF0'>";
		var headerNonRepetingData = dispTable[index].split("#");
		content = content + "<td align='center'>"+parseInt(index+1)+"</td>";
		content = content + "<td align='center'>"+headerNonRepetingData[0]+"</td>";
		if (headerNonRepetingData[0] == headerNonRepetingData[1]) {
			content = content + "<td align='center'>Currently Logged In</td>";
		} else {
			content = content + "<td align='center'>"+headerNonRepetingData[1]+"</td>";
		}
		//content = content + "<td align='center'>"+headerNonRepetingData[1]+"</td>";
		content = content + "</tr>";
		tableContents = tableContents + content;
	}
	var mainContent = headingString + tableContents+"</tbody></table>";
	tableDiv.innerHTML = mainContent;
	new SortableTable(document.getElementById('myTable2'), 1);
}
/**
 * Function is used to go back to the overview div from the detailed div
 */
function goBack(){
	document.getElementById('detailedDiv').style.display="none";
	document.getElementById('miscHeaderContentDiv').style.display="block";
	document.getElementById('mainHeaderDivCont').style.display="block";	
}
/**
 * Function is used to generate and download the excel report
 */
function generateMiscExcel(){
	var link = document.getElementById('generateXL');
	var reportLink = "../../admin/reports/generateExcelMisc/";
	//var functionGrp = document.getElementById("inputFunctionGroup").options[document.getElementById("inputFunctionGroup").selectedIndex].value;
	//var jobLevel = document.getElementById("inputJobLevel").options[document.getElementById("inputJobLevel").selectedIndex].value;
	//if(functionGrp == ""){
	//	functionGrp = "JCTDUMMYVARJCTFN";
	//}
	//if(functionGrp == ""){
	//	functionGrp = "JCTDUMMYVARJCTJL";
	//}
	
	//reportLink = reportLink + functionGrp + "/" + jobLevel + ".xls";
	reportLink = reportLink+"report.xls";
	
	link.setAttribute("href", reportLink);
	//link.submit();
}