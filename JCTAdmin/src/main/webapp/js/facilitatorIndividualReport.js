var userTypeInt = 0;
$(document).ready(function(){
	$("#existing_data_div").hide();
});

/**
 * Function is called when search button is clicked.
 */
function generateReport() {	
	var usrType = document.getElementById("inptUserType").options[document.getElementById("inptUserType").selectedIndex].value;
	if (usrType.trim().length == 0) { // Nothing has been selected
		$("#existing_data_div").hide();
		alertify.alert("Please select an user type to generate report.");
		return false;
	} else {
		userTypeInt = usrType == "F" ? 3 : (usrType == "I" ? 1 : 0);
		var facIndMdl = Spine.Model.sub();
		facIndMdl.configure("/admin/facIndRprt/generateFacIndUIReport", "usrType");
		facIndMdl.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new facIndMdl({  
			//adminEmailId: sessionStorage.getItem("jctEmail"),
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
				plotReport(userTypeInt);
			}  else {
				//Show error message
				alertify.alert("Some thing went wrong.");
			}
		});
	}
}

/**
 * Function plots report on ui
 */
function plotReport(userTypeInt) {	
	$("#existing_data_div").show();
	$("#existing_list_Id").show();	
	var obj = jQuery.parseJSON(sessionStorage.getItem("jsonStr"));
	var dispTable = obj.reportList.split("$$$");
	var userType = userTypeInt == 1 ? "Indivudual" : ( userTypeInt == 3 ? "Facilitator" : "All");
	content = "<div style='float:left;font-weight:bold;width:45%;color:#616161;'>Report fetched for "+userType+" user: "+(dispTable.length-1)+"</div>"+
			"<div style='float:right;font-weight:bold;width:20%;color:#616161;margin-right: -10%;'>Export to: <a id='generateExcel' href='#' onclick='generateFaciIndiExcel()'><img src='../img/Excel-icon.png'></a></div>";
	document.getElementById("existing_list_Id").innerHTML = content;
	
	if ( !$(".form_area_top").is(":visible") ) {		//	when came back from individual's details area
		$(".form_area_top").show();	
		$("#existing_list_Id").show();
	}	
	if( dispTable.length-1 > 0 ){						//		when at-least 1 row fetched from database
		//		******   Creating Table Header
		tableContent = "<br><table width='100%' border='1' bordercolor='#78C0D3' id='myTable' class='tablesorter'><thead class='tab_header'><tr><th>Sl. no</th><th>User name</th><th>Profile name</th><th>Group name</th><th width='22%'>Created By E-Commerce/Admin</th></thead><tbody>";
		//		******	Creating Table Body
		for(var index=0; index<dispTable.length-1; index++){
			var td = dispTable[index].split("#");		// 0 - USER NAME  1 - PROFILE NAME  2 - GROUP NAME  3 - CREATED BY  4 - CREATED ON  5 - EXPIRATION DATE			
			console.log(td);
			var content = "<tr bgcolor='#D2EAF0' style='color:#616161;font-weight:bold;'>";
			content = content + "<td align='center'>"+(index+1)+"</td>";	//	Sl. no
			content = content + "<td align='center'><a class='link' onclick = 'expandView(\""+td[0]+"\",\""+td[1]+"\",\""+td[2]+"\",\""+td[3]+"\",\""+td[4]+"\",\""+td[5]+"\")'><b>"+td[0]+"</b></a></td>";		//	USER NAME
			content = content + "<td align='center'>"+td[1]+"</td>";		//	PROFILE NAME
			content = content + "<td align='center'>"+td[2]+"</td>";		//	GROUP NAME			
			content = content + "<td align='center'>"+td[3]+"</td>";		//	CREATED BY
			content = content + "</tr>";	
			tableContent = tableContent + content;							//		Adding the current row to the table
		}
		tableContent = tableContent + "</table>";
	} else {											//		No Row selected, because user has not shared any diagram.
		$("#existing_list_Id").hide();
		tableContent = "<div align='center'><br /><br /><br /><img src='../img/no-record.png'><br><div class='textStyleNoExist'>No result found</div><br/></div>";
	}
	document.getElementById("existingUserListId").innerHTML = tableContent;
	$("table").tablesorter(); 
}

function expandView(td0,td1,td2,td3,td4,td5){
	
	$(".form_area_top").hide();	
	$("#existing_list_Id").hide();
	
	tableContent = "<div id='userDetailsHeader' style='padding: 3%;'><div style='float: left;width:50%;font-weight:bold;color:#616161;'>User Details: "+td0+" </div><div style='float:left;width:50%;'><input id='backButton' type='button' class='go_back_btn' value='Go Back' onclick='plotReport()'></div></div>";
	tableContent = tableContent + "<table width='100%' border='1' bordercolor='#78C0D3' id='myTable'><thead class='tab_header'><tr><th>User name</th><th>Profile name</th><th>Group name</th><th width='16%'>Created By E-Commerce/Admin</th><th>Created on</th><th>Expiration Date</th></thead><tbody>";
	var content = "<tr bgcolor='#D2EAF0' style='color:#616161;font-weight:bold;'>";
	content = content + "<td align='center'>"+td0+"</td>";						//	USER NAME
	content = content + "<td align='center'>"+td1+"</td>";						//	PROFILE NAME
	content = content + "<td align='center'>"+td2+"</td>";						//	GROUP NAME
	content = content + "<td align='center'>"+td3+"</td>";						//	CREATED BY
	content = content + "<td align='center'>"+td4.substr(0, 10)+"</td>";		//	CREATED ON
	content = content + "<td align='center'>"+td5.substr(0, 10)+"</td>";		//	EXPIRATION DATE
	content = content + "</tr>";	
	tableContent = tableContent + content;										//		Adding row to the table
	tableContent = tableContent + "</table>";
	document.getElementById("existingUserListId").innerHTML = tableContent;	
}
/**
 * Function is to export user table to excel sheet. 
 */
function generateFaciIndiExcel(){
	var link = document.getElementById('generateExcel');
	var reportLink = "../../admin/facIndRprt/generateFaciIndiExcel/";
	reportLink = reportLink+ userTypeInt +"/report.xls";	//	userTypeInt is a global variable
	link.setAttribute("href", reportLink);
}