/**
 * Function is to export payment details to excel sheet. 
 */
function generateExcel(e) {
	var userType = document.getElementById("inptUserType").options[document.getElementById("inptUserType").selectedIndex].value;
	if (userType == "") {
		alertify.alert("Please select an user type.");
		e.preventDefault();
		e.stopPropagation();
		//return false;
	}
	var link = document.getElementById('generateXL');
	var reportLink = "../../admin/reports/generatePaymentExcel/";
	reportLink = reportLink+userType+"/report.xls";	//	userTypeInt is a global variable
	link.setAttribute("href", reportLink);
}