var tableContent = "";

$(document).ready(function() {
	//generateSurveyReportAll();
});

/**
 * Function called to populate all the survey reports
 */
/*function generateSurveyReportAll(){
		var userTypeInt = 5;		
		var facIndMdl = Spine.Model.sub();
		facIndMdl.configure("/admin/reports/generateSurveyReport", "adminEmailId", "usrType");
		facIndMdl.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new facIndMdl({  
			adminEmailId: sessionStorage.getItem("jctEmail"),
			usrType: userTypeInt
		});
		modelPopulator.save(); //POST
		
		facIndMdl.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
		});
		facIndMdl.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			var statusCode = obj.statusCode;
			if(statusCode == 200) {
				sessionStorage.setItem("jsonStr",jsonStr);
				populateReport();
			}else {
				alertify.alert("Some thing went wrong.");			 
			}
		});
}*/
/**
 * Function populates the overview report
 */
/*function populateReport(){
	var obj = jQuery.parseJSON(sessionStorage.getItem("jsonStr"));
	var tableDiv = document.getElementById("surveyRptDiv");
	tableContent = "";
	var header1 = "";
	var counter=1;
	var headingString = "<table class='etable' width='100%' border='1' bordercolor='#78C0D3' id='myTable'><thead class='tab_header'><tr><th>SL. No.</th><th>Name</th><th>User Type</th><th>Created On</th><th>Expiration Date</th><td align='center'><b>Survey Question</b></td><td align='center'><b>Answer</b></td></tr></thead><tbody>";
	var gXLDiv = document.getElementById('genearteXLDiv');	
	var dispTable = obj.reportList.split("@@@");
	var size = parseInt(dispTable.length-1);
	if(size > 0){
	document.getElementById("beforeSketchResultDiv").style.display="block";
	gXLDiv.innerHTML = "<table width='100%'><tr><td align='right'><b>Export To</b><a id='generateXLRep' href='#' onclick='generateXLReport()'>&nbsp;<img src='../img/Excel-icon.png' /></td></tr></table>";
	}else{
		document.getElementById("beforeSketchResultDiv").style.display="none";
		document.getElementById("genearteXLDiv").style.display="none";
		tableDiv.innerHTML="<div align='center'> <br /><br /><br /><img src='../img/no-record.png'><br /><div class='textStyleNoExist'>No Records found</div></div>";
	}
	if( dispTable.length-1 > 0 ){
		for(var index=0; index<dispTable.length; index++){
			var rowSpanCalculator = dispTable[index].split("$$$");
			var rowSpanLength=rowSpanCalculator.length-1;
			for(var i=0;i<rowSpanCalculator.length-1;i++){
				var content = "<tr bgcolor='#D2EAF0'>";
				var hderNonRepetingData = rowSpanCalculator[i].split("#");
				var facInd=hderNonRepetingData[1];
				if(facInd==1){
					hderNonRepetingData[1]="Individual User";
				}else{
					hderNonRepetingData[1]="Facilitator";
				}
									
				if(header1!=hderNonRepetingData[0]){
					header1 = hderNonRepetingData[0];
					content = content + "<td align='center' rowspan='"+rowSpanLength+"'>"+counter+"</td>";
					content = content + "<td align='center' rowspan='"+rowSpanLength+"'>"+hderNonRepetingData[0]+"</td>";
					content = content + "<td align='center' rowspan='"+rowSpanLength+"'>"+hderNonRepetingData[1]+"</td>";
					content = content + "<td align='center' rowspan='"+rowSpanLength+"'>"+dateFormatForReports(hderNonRepetingData[2])+"</td>";
					content = content + "<td align='center' rowspan='"+rowSpanLength+"'>"+dateFormatForReports(hderNonRepetingData[3])+"</td>";
					content = content + "<td align='center'>"+hderNonRepetingData[4]+"</td>";
					content = content + "<td align='center'>"+hderNonRepetingData[5]+"</td>";
					content = content + "</tr>";
					counter = counter+1;
				}
				else{
					content = content + "<td align='center'>"+hderNonRepetingData[4]+"</td>";
					content = content + "<td align='center'>"+hderNonRepetingData[5]+"</td>";
					content = content + "</tr>";	
				}
				tableContent = tableContent + content;
			}	
		}
		var mainContent = headingString + tableContent+"</tbody></table>";
		tableDiv.innerHTML = mainContent;
	}
	
}*/


/**
 *  Function called to generate specific survey report for facilitator or individual
 */
/*function generateReport() {
	var usrType = document.getElementById("inptUserType").options[document.getElementById("inptUserType").selectedIndex].value;
	if (usrType.trim().length == 0) { // Nothing has been selected
		alertify.alert("Please select an user type to generate report.");
		return false;
	} else {
		var userTypeInt = 1;
		if (usrType == "F") {
			userTypeInt = 3;
		}
		var facIndMdl = Spine.Model.sub();
		facIndMdl.configure("/admin/reports/generateSurveyReport", "adminEmailId", "usrType");
		facIndMdl.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new facIndMdl({  
			adminEmailId: sessionStorage.getItem("jctEmail"),
			usrType: userTypeInt
		});
		modelPopulator.save(); //POST
		
		facIndMdl.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
		});
		facIndMdl.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			var statusCode = obj.statusCode;
			if(statusCode == 200) {
				sessionStorage.setItem("jsonStr",jsonStr);
				populateReport();
			}else {
				alertify.alert("Some thing went wrong.");			 
			}
		});
	}
}*/


/**
 * Function call to generate the excel report 
 */
function generateXLReport(e){
	var link = document.getElementById('generateXL');
	var reportLink = "../../admin/reports/generateSurveyExcel/";
	var userType = document.getElementById("inptUserType").options[document.getElementById("inptUserType").selectedIndex].value;
	if(userType == ""){
		alertify.alert("Please select user type.");
		e.preventDefault();
		e.stopPropagation();
		//return false;
	}
	reportLink = reportLink + userType + "/survey_report.xls";	
	link.setAttribute("href", reportLink);
}
