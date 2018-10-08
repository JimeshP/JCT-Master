/**
 * Function call to generate the excel report 
 * param null
 */
function generateXLReport() {
	//check which all are checked
	//order bs, as, rq, ap
	var checkedPreference = "";
	if (document.getElementById('bs').checked) {
		checkedPreference = checkedPreference + "C~";
	} else {
		checkedPreference = checkedPreference + "N~";
	}
	if (document.getElementById('as').checked) {
		checkedPreference = checkedPreference + "C~";
	} else {
		checkedPreference = checkedPreference + "N~";
	}
	if (document.getElementById('rq').checked) {
		checkedPreference = checkedPreference + "C~";
	} else {
		checkedPreference = checkedPreference + "N~";
	}
	if (document.getElementById('ap').checked) {
		checkedPreference = checkedPreference + "C~";
	} else {
		checkedPreference = checkedPreference + "N~";
	}
	if (document.getElementById('sr').checked) {
		checkedPreference = checkedPreference + "C";
	} else {
		checkedPreference = checkedPreference + "N";
	}
	var dateString = getDateString();
	var link = document.getElementById('generateXL');
	var reportLink = "../../admin/allReports/generateExcelAllReport/";
	reportLink = reportLink + checkedPreference + "/Merged_Report - "+dateString+".xlsx";
	link.setAttribute("href", reportLink);
}
/**
 * function returns the date string in MMM-dd-yyyy format
 */
function getDateString() {
	var today = new Date();
	var monthInt = today.getMonth();
	var monthVal = "";
	if (monthInt == 0) {
		monthVal = "Jan";
	} else if (monthInt == 1) {
		monthVal = "Feb";
	} else if (monthInt == 2) {
		monthVal = "Mar";
	} else if (monthInt == 3) {
		monthVal = "Apr";
	} else if (monthInt == 4) {
		monthVal = "May";
	} else if (monthInt == 5) {
		monthVal = "Jun";
	} else if (monthInt == 6) {
		monthVal = "Jly";
	} else if (monthInt == 7) {
		monthVal = "Aug";
	} else if (monthInt == 8) {
		monthVal = "Sep";
	} else if (monthInt == 9) {
		monthVal = "Oct";
	} else if (monthInt == 10) {
		monthVal = "Nov";
	} else if (monthInt == 11) {
		monthVal = "Dec";
	}
	var fullYear = today.getFullYear();
	var todayDate = today.getDate();
	return monthVal+"-"+todayDate+"-"+fullYear;
}