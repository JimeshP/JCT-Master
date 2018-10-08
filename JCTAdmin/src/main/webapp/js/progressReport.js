var tableContent = "";
/**
 * Function calls when the page is loaded 
 * to populate the drop-down for Function Group 
 * and job level and to populate the report
 */
$(document).ready(function() {
	populateAllProgressReports();
	fetchUserGroupForFacilitator();
});

/**
 * Function fetches user progress reports
 * @param null
 */
window.setInterval(function(){
	populateUserProgressReports();
}, 60000);
 
function populateUserProgressReports(){
	var userGroup = document.getElementById("inputUserGroupInpt").options[document.getElementById("inputUserGroupInpt").selectedIndex].value;
	var userGroupName = document.getElementById("inputUserGroupInpt").options[document.getElementById("inputUserGroupInpt").selectedIndex].text;
	var str = userGroup.indexOf("!");
	var userGroupVal = userGroup.substring(0,str);
	if(userGroup == ""){
	//	document.getElementById("newProgressResult").style.display = 'block';
		populateAllProgressReports();
	}
	else{
		/*alertify.alert("No Records present.");
		document.getElementById("newProgressResult").innerHTML = "<div align='center' class='noRecordDiv'> <br /><br /><br /><img src='../img/no-record.png'><br />No Users Avalaible</div>";*/
		var bsModel = Spine.Model.sub();
		bsModel.configure("/admin/facIndRprt/populateProgressReports", "jctEmail", "userGroupName");
		bsModel.extend( Spine.Model.Ajax );
		
		var populateModel = new bsModel({  
			jctEmail: sessionStorage.getItem("jctEmail"),
			userGroupName: userGroupName
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
				sessionStorage.setItem("jsonStr",jsonStr); //contains whole population string
				populateProgressReport();
			//	window.location = "/admin/view/progressReport.jsp";
				
			}
		});
	}
}


function populateAllProgressReports(){
	var bsModel = Spine.Model.sub();
	bsModel.configure("/admin/facIndRprt/populateAllProgressReports", "jctEmail");
	bsModel.extend( Spine.Model.Ajax );
	
	var populateModel = new bsModel({  
		jctEmail: sessionStorage.getItem("jctEmail")
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
			sessionStorage.setItem("jsonStr",jsonStr); //contains whole population string
			populateProgressReport();
		//	window.location = "/admin/view/progressReport.jsp";
			
		}
	});
}


/**
 * Function populates the report 
 * for all job level and function group
 * @param null
 */
function populateProgressReport(){
	var obj = jQuery.parseJSON(sessionStorage.getItem("jsonStr"));
	tableContent = "";
	var tableDiv = document.getElementById("newProgressResult");
	var headingString = "<table width='94%' id='myTable' border=1 align='center' bordercolor='#78C0D3' class='etable tablesorter'><thead class='tab_header'><tr><th width='4%' class='remove_arrow_th'>#</th><th>Name (Last, First)</th><th>User Group Name</th><th>Email</th><th>Progress</th><th>Last Login Date</th><th>Total Time Spent (HH:MM)</th><th>Expiration Date</th></tr></thead><tbody>";
	var dispTable = obj.reportList.split("$$$");
	var size = parseInt(dispTable.length-1);
	var pageName = "";
	if(dispTable.length-1 > 0){
		document.getElementById("progressReportDiv").style.display="block";
		for(var index=0; index<dispTable.length-1; index++){
		var content = "<tr bgcolor='#D2EAF0'>";
		var headerNonRepetingData = dispTable[index].split("#");

		content = content + "<td align='center'>"+parseInt(index+1)+"</td>";
	
		content = content + "<td align='center'>"+headerNonRepetingData[6]+", "+headerNonRepetingData[7]+"</td>";
		
		content = content + "<td align='center'>"+headerNonRepetingData[1]+"</td>";
		content = content + "<td align='center'>"+headerNonRepetingData[0]+"</td>";
		//content = content + "<td align='center'>"+headerNonRepetingData[3]+"</td>"; 
		//alert(headerNonRepetingData[2]);
		
		var spaceLogin = headerNonRepetingData[4].indexOf(" ");
		var strLogin = headerNonRepetingData[4].substring(0,spaceLogin);
		var replaceStrLogin =strLogin.replace(/-/g, ' ');
		var lastLogin = new Date(replaceStrLogin);
		var lastLoginTime = lastLogin.getTime();
		var lastLoginDate = dateformat(new Date(lastLoginTime));
		
		if(headerNonRepetingData[8] != 'N/A'){
		if(headerNonRepetingData[8].search('landing-page.jsp') != -1){
				pageName = "Landing Page";
		}
		else if(headerNonRepetingData[8].search('beforeSketch.jsp') != -1){
			pageName = "1. Before Sketch - Tasks";
		}
		else if(headerNonRepetingData[8].search('Questionaire.jsp') != -1){
			pageName = "2. Before Sketch - Reflection Questions";
		} 
		else if(headerNonRepetingData[8].search('afterSketch1.jsp') != -1){
			pageName = "3. After Diagram - Mapping Yourself";
		} 
		else if(headerNonRepetingData[8].search('afterSketch2.jsp') != -1){
			pageName = "4. After Diagram - Tasks";
		} 
		else if(headerNonRepetingData[8].search('actionPlan.jsp') != -1){
			pageName = "5. After Diagram - Action Plan";
		} 
		else if(headerNonRepetingData[8].search('finalPage.jsp') != -1){
			pageName = "6. Final Report";
		} else if(headerNonRepetingData[8].search('resultPage.jsp') != -1){
			pageName = "6. Final Report";
		} else if(headerNonRepetingData[8].search('resultPageFinal.jsp') != -1){
			pageName = "6. Final Report";
		}else{
			pageName = headerNonRepetingData[8];
		}
		content = content + "<td align='center'>"+pageName+"</td>";
		}
		else{
			content = content + "<td align='center'>"+headerNonRepetingData[8]+"</td>";
		}
		
		if(headerNonRepetingData[4]=='N/A'){
			content = content + "<td align='center'>"+headerNonRepetingData[4]+"</td>"; 
		}else{
		content = content + "<td align='center'>"+dateFormatForReports(headerNonRepetingData[4])+"</td>"; 
		}
		//console.log(headerNonRepetingData[5]);
		//content = content + "<td align='center'>"+headerNonRepetingData[5]+":"+headerNonRepetingData[3]+"</td>";
		var spt = headerNonRepetingData[5].split(":");
		if (spt[1] != undefined) {
		content = content + "<td align='center'>"+spt[0]+" : "+spt[1]+"</td>";
		} else {
		content = content + "<td align='center'>"+spt[0]+" : 00</td>";
		}
		/*if(headerNonRepetingData[5]=='N/A'){
			content = content + "<td align='center'>"+headerNonRepetingData[5]+":"+headerNonRepetingData[3]+"</td>";
			} else{			
				var sp = headerNonRepetingData[5].split(":");				
				content = content + "<td align='center'>"+sp[0]+" : "+sp[1]+"</td>";
			}*/
		content = content + "<td align='center'>"+dateFormatForReports(headerNonRepetingData[2])+"</td>";
				
		tableContent = tableContent + content;
		}
		var mainContent = headingString + tableContent+"</tbody></table>";
		tableDiv.innerHTML = mainContent;
		//$("table").tablesorter(); 
		$(function(){
			  $("table").tablesorter({
			    headers: {
			      0: { sorter: false }      // disable first column			     
			    }
			  });
			});

	}
	else{
		document.getElementById("progressReportDiv").style.display="none";
		var mainContentNoRecord = "<div align='center' class='noRecordDiv'> <br /><br /><br /><img src='../img/no-record.png'><br /><div class='textStyleNoExist'>No Record Found.</div></div>";
		tableDiv.innerHTML = mainContentNoRecord;
	}
	
	

	/*$("table").tablesorter();*/
}


function fetchUserGroupForFacilitator() {
	var userGrp = Spine.Model.sub();
	userGrp.configure("/admin/manageuserFacilitator/populateUserGroupDroDwnFacilitator", "customerId");
	userGrp.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userGrp({  
		customerId: sessionStorage.getItem("customerId")
	});
	modelPopulator.save(); //POST
	userGrp.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	userGrp.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if(statusCode == 200) {
			populateDropDown(obj.userGroupMap);
		}  else {
			//Show error message
			alertify.alert("Some thing went wrong.");
		}
	});
}

/**
 * Function to populate the dropdown for user group
 */
function populateDropDown(userGroupMap) {
	var userGroupSelect = document.getElementById("inputUserGroupInpt");
	for (var key in userGroupMap) {				
		var option = document.createElement("option");
		option.text = userGroupMap[key];
	    option.value = key+"!"+userGroupMap[key];
	    option.className = "form-control-general";
	    try {
	    	userGroupSelect.add(option, null); //Standard 
	    }catch(error) {
	    	//regionSelect.add(option); // IE only
	    }
	}
}
/**
 * Function to calculate the total time
 * @param secs
 */
function secondsToTime(secs) {
    var hours = Math.floor(secs / (60 * 60));
    if (hours < 10) {
    	hours = "0"+hours;
    }
    var divisor_for_minutes = secs % (60 * 60);
    var minutes = Math.floor(divisor_for_minutes / 60);
    if (minutes < 10) {
    	minutes = "0"+minutes;
    }
    var divisor_for_seconds = divisor_for_minutes % 60;
    var seconds = Math.ceil(divisor_for_seconds);
    if (seconds < 10) {
    	seconds = "0"+seconds;
    }
    return "<b>"+hours+" : "+minutes+ "</b>";
}


/**
 * Function to calculate the total time
 * @param secs
 */
function calculateTotalTime(hour, mint) {	
	if(hour.length < 2) {
		hour = "0"+hour;
	} 
	if(mint.length < 2) {
		mint = "0"+mint;
	} 
	totalTime = hour+" : "+mint;
	return totalTime;
}
