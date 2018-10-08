var tableContent = "";
/**
 * Function calls when the page is loaded 
 * to populate the drop-down for Function Group 
 * and job level and to populate the report
 */
$(document).ready(function() {
	//populateDropdowns();
	populateReport();
	if (sessionStorage.getItem("occupation")) {
		document.getElementById("ocptnSrchTxtId").value = sessionStorage.getItem("occupation"); 
	}
});

/**
 * Function populates the drop-down for Function Group and job level
 * @param null
 */
function populateDropdowns(){
	var option = null;
	var obj = jQuery.parseJSON(sessionStorage.getItem("jsonStr"));
	var functionGrpSelect = document.getElementById("inputFunctionGroup");
	for (var key in obj.functionGroup) {
		//Create a option element dynamically
		option = document.createElement("option");
		option.text = obj.functionGroup[key];
	    option.value = obj.functionGroup[key];
	    option.className = "form-control-general";
	    if(sessionStorage.getItem("fnLSelected")) {
	    	if(option.value == sessionStorage.getItem("fnLSelected")) {
		    	option.setAttribute("selected", "selected");
		    }
	    }
	    try {
	    	functionGrpSelect.add(option, null); //Standard 
	    }catch(error) {
	    	//functionGrpSelect.add(option); // IE only
	    }
	}
	
	var jobLevelSelect = document.getElementById("inputJobLevel");
	for (var keys in obj.jobLevel) {
		//Create a option element dynamically
		option = document.createElement("option");
		option.text = obj.jobLevel[keys];
	    option.value = obj.jobLevel[keys];
	    option.className = "form-control-general";
	    if(sessionStorage.getItem("jLSelected")){
	    	if(option.value == sessionStorage.getItem("jLSelected")){
		    	option.setAttribute("selected", "selected");
		    }
	    }
	    
	    try {
	    	jobLevelSelect.add(option, null); //Standard 
	    }catch(error) {
	    	//jobLevelSelect.add(option); // IE only
	    }
	}
	sessionStorage.setItem('bsrptdropdownpopulated','Y');
}


/**
 * Function populates the report for all function group and job level
 * @param null
 * 
 */
function populateReport(){
	var obj = jQuery.parseJSON(sessionStorage.getItem("jsonStr"));
	var header1 = "";
	var memory = "";
	tableContent = "";
	var tableDiv = document.getElementById("actionPlanRptDiv");
	var hdrDiv = document.getElementById('actionPlanRptHdrDiv');
	var gXLDiv = document.getElementById('genearteXLDiv');	
	var headerDispTable = obj.tableHeaderString.split("&&&");
	var headingString = "<table width='100%' class='etable'><thead class='tab_header'><tr><th width='14%'>"+headerDispTable[13]+"</th><th width='14%'>"+headerDispTable[0]+"</th><th width='14%'>"+headerDispTable[1]+"</th><th width='13%'>"+headerDispTable[3]+"</th><th width='13%'>"+headerDispTable[4]+"</th><th width='34%'>"+headerDispTable[5]+"</th></tr></thead><tbody>";	
	var dispTable = obj.reportList.split("$^$");
	var indFirst = parseInt(sessionStorage.getItem("firstResultAP")) + 1;
	var indSecond = 0;
	if(sessionStorage.getItem("firstResultAP") == 0){
		indSecond = 20;
	} else {
		var total = obj.totalCount;
		indSecond = parseInt(sessionStorage.getItem("firstResultAP")) + 20;
		if(indSecond >= total){
			//calculate how much more
			indSecond = parseInt(sessionStorage.getItem("firstResultAP")) + parseInt(dispTable.length-1);
		} 
	}
	if(indSecond > parseInt(obj.totalCount)){
		indSecond = obj.totalCount;
	}
	
	if(dispTable.length-1 > 0){
		hdrDiv.innerHTML = "<table width='100%'><tr><td><b>"+headerDispTable[6]+" </b>"+indFirst+" <b> "+headerDispTable[7]+" </b>"+indSecond+" <b>"+headerDispTable[8]+" </b>"+obj.totalCount+"</td></tr></table>";
		gXLDiv.innerHTML = "<table width='100%'><tr><td align='right'><b>"+headerDispTable[9]+" </b><a id='generateXL' href='#' onclick='generateReport()'><img src='../img/Excel-icon.png' /></td></tr></table>";
	}
	for(var index=0; index<dispTable.length-1; index++){
		var header0 = "";
		var header2 = "";
		var header3 = "";
		var header4 = "";
		var counter = 1;
		var headerNonRepetingData = dispTable[index].split("#");
		var rowSpanCalculator = headerNonRepetingData[5].split("~").length-1;
		var taskTimePopulator = headerNonRepetingData[5].split("~");	
		var energyCounter = (headerNonRepetingData[5].split("!").length-1)/2;
		//alert(headerNonRepetingData);
		var textareaBgColor = "";
		for(var taskCounter=0; taskCounter<rowSpanCalculator; taskCounter++){				
			var headerId = index;
			if(headerId % 2 == 0)
			{
				var content = "<tr id = '"+headerId+"' bgcolor='#D2EAF0'>";
				textareaBgColor = "#D2EAF0";
			} else
			{
				var content = "<tr id = '"+headerId+"' bgcolor='#F1F1F1'>";
				textareaBgColor = "#F1F1F1";
			}
			//var content = "<tr id = '"+headerId+"' bgcolor='#D2EAF0'>";
			if(header0 != headerNonRepetingData[4]){
				header0 = headerNonRepetingData[4];
				content = content + "<td rowspan='"+rowSpanCalculator+"'>"+headerNonRepetingData[4]+"</td>";
			} 
			if(header1 != headerNonRepetingData[0]){
				header1 = headerNonRepetingData[0];
				content = content + "<td rowspan='"+rowSpanCalculator+"'>"+headerNonRepetingData[0]+"</td>";
			} 
			if(header2 != headerNonRepetingData[1]){
				header2 = headerNonRepetingData[1];
				content = content + "<td rowspan='"+rowSpanCalculator+"'>"+headerNonRepetingData[1]+"</td>";
				//content = content + "<td><textarea rows='2' cols='38' readonly style='background-color: "+textareaBgColor+"; border: none; text-align:center;'>"+headerNonRepetingData[1]+"</textarea></td>";
			} 
			if(header3 != headerNonRepetingData[2]){
				header3 = headerNonRepetingData[2];
				content = content + "<td rowspan='"+rowSpanCalculator+"'>"+headerNonRepetingData[2]+"</td>";
			}	
			if(header4 != headerNonRepetingData[3]){
				header4 = headerNonRepetingData[3];
				content = content + "<td rowspan='"+rowSpanCalculator+"'>"+headerNonRepetingData[3]+"</td>";
			}	
			/*if(memory != taskTimePopulator[taskCounter].split('!')[0]){
				counter++;
				memory = taskTimePopulator[taskCounter].split('!')[0];
				content = content + "<td rowspan='"+energyCounter+"'>"+taskTimePopulator[taskCounter].split('!')[0]+"</td>";				
			} */
			if(memory != taskTimePopulator[taskCounter].split('!')[0]){
				counter++;
				memory = taskTimePopulator[taskCounter].split('!')[0];
				content = content + "<td>"+taskTimePopulator[taskCounter].split('!')[0]+"</td>";				
			} else {				
				content = content + "<td>---</td>";
			}	
			
			if(taskTimePopulator[taskCounter].split('!')[1] == "" || null == taskTimePopulator[taskCounter].split('!')[1]){
				content = content + "<td>---</td>";
			} else {
				content = content + "<td><textarea rows='2' cols='105' readonly style='background-color: "+textareaBgColor+"; border: none;'>"+taskTimePopulator[taskCounter].split('!')[1]+"</textarea></td>";
			}
			content = content + "</tr>";
			tableContent = tableContent + content;
			
		}
	}
	if(dispTable.length-1 > 0){
		document.getElementById("actionPlanResultDiv").style.display="block";
		var mainContent = headingString + tableContent+"</tbody></table>";
		//Both disabled
		if((obj.previousDisabled == 1) && (obj.nextDisabled == 1)) {
			mainContent = mainContent + "<br /> <div align='center'> <span><< &nbsp; "+headerDispTable[10]+"</span> &nbsp;&nbsp;&nbsp;&nbsp; <span>"+headerDispTable[11]+" &nbsp; >></span>";
		} else if((obj.previousDisabled == 1) && (obj.nextDisabled == 0)) {
			mainContent = mainContent + "<br /> <div align='center'> <span><< &nbsp; "+headerDispTable[10]+"</span> &nbsp;&nbsp;&nbsp;&nbsp; <a href='#' onClick='next()'>"+headerDispTable[11]+" &nbsp; >></a>";
		} else if((obj.previousDisabled == 0) && (obj.nextDisabled == 1)) {
			mainContent = mainContent + "<br /> <div align='center'> <a href='#' onClick='previous()'><< &nbsp; "+headerDispTable[10]+"</a> &nbsp;&nbsp;&nbsp;&nbsp; <span>"+headerDispTable[11]+" &nbsp; >></span>";
		} else {
			mainContent = mainContent + "<br /> <div align='center'> <a href='#' onClick='"+headerDispTable[10]+"()'><< &nbsp; Previous</a> &nbsp;&nbsp;&nbsp;&nbsp; <a href='#' onClick='next()'>"+headerDispTable[11]+" &nbsp; >></a>";
		}
		
	} else {
		mainContent = "<div align='center'> <br /><br /><br /><img src='../img/no-record.png'><br /><div class='textStyleNoExist'>"+headerDispTable[12]+"</div></div>";
		document.getElementById("actionPlanResultDiv").style.display="none";
	}
	tableDiv.innerHTML = mainContent;
}

/**
 * Function populates the report 
 * for the selected job level and function group
 * @param dist
 */
function populateActionPlanSelectedReport(dist){
	//var functionGrp = document.getElementById("inputFunctionGroup").options[document.getElementById("inputFunctionGroup").selectedIndex].value;
	//var jobLevel = document.getElementById("inputJobLevel").options[document.getElementById("inputJobLevel").selectedIndex].value;
	if(dist == 'newid'){
		if(sessionStorage.getItem("firstResultAP")){
			sessionStorage.setItem("firstResultAP", 0);
		}
	}
	/*if(jobLevel == "All"){
		jobLevel= "";
	}*/
	//sessionStorage.setItem("fnLSelected", functionGrp);
	//sessionStorage.setItem("jLSelected", jobLevel);
	var occupationVal = "";
	if (document.getElementById("ocptnSrchTxtId")) {
		if (document.getElementById("ocptnSrchTxtId").value.trim().length != 0) {
			occupationVal = document.getElementById("ocptnSrchTxtId").value;
		} 
	}
	var bsModel = Spine.Model.sub();
	bsModel.configure("/admin/reports/populateActionPlanSelectedReport", "occupationVal", "recordIndex");
	bsModel.extend( Spine.Model.Ajax );
	
	//Populate the model
	var populateModel = new bsModel({  
		occupationVal: occupationVal,
		recordIndex: sessionStorage.getItem("firstResultAP")
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
			sessionStorage.setItem("occupation",occupationVal);
			sessionStorage.setItem("jsonStr",jsonStr); //contains whole population string
			window.location = "reportsActionPlan.jsp";
		}
	});
}

/**
 * Function call while clicking on next button 
 * @param null
 */
function next(){
	if(sessionStorage.getItem("firstResultAP")){
		var count = parseInt(sessionStorage.getItem("firstResultAP")) + 20;
		sessionStorage.setItem("firstResultAP",count);
		populateActionPlanSelectedReport('existing');
	}
}

/**
 * Function call while clicking on previous button 
 * @param null
 */
function previous(){
	if(sessionStorage.getItem("firstResultAP")){
		var count = parseInt(sessionStorage.getItem("firstResultAP")) - 20;
		sessionStorage.setItem("firstResultAP",count);
		populateActionPlanSelectedReport('existing');
	}
}

/**
 * Function call to generate the excel report
 * @param null
 */
function generateReport(){
	var link = document.getElementById('generateXL');
	var reportLink = "../../admin/reports/generateExcelActionPlan/";
	//var functionGrp = document.getElementById("inputFunctionGroup").options[document.getElementById("inputFunctionGroup").selectedIndex].value;
	//var jobLevel = document.getElementById("inputJobLevel").options[document.getElementById("inputJobLevel").selectedIndex].value;
	/*if(functionGrp == ""){
		functionGrp = "JCTDUMMYVARJCTFN";
	}
	if(jobLevel == ""){
		jobLevel = "JCTDUMMYVARJCTJL";
	}*/
	var occupation = "";
	if (sessionStorage.getItem("occupation")) {
		occupation = sessionStorage.getItem("occupation");
	} else {
		sessionStorage.setItem("occupation", "NA");
		occupation = "NA";
	}
	reportLink = reportLink + occupation + "/report.xls";	
	link.setAttribute("href", reportLink);
}

function searchOccupationSearch() {
	var occupationSearchVal = document.getElementById("ocptnSrchTxtId").value.trim();
	if (occupationSearchVal.trim().length == 0) {		
		alertify.alert("<p>Please provide occupation search String");
		$("#searchOccupation").click(function() {$.fancybox.close({});});
		return false;
	} else {
		document.getElementById("occupationPlottingDiv").innerHTML = "<br /><br /><img src='../img/Processing.gif' />";			
		$("#searchOccupation").click(function() {
			$.fancybox.open({
				href : '#occupationDiv',
				closeClick : false,
				openEffect : 'elastic',
				openSpeed  : 150,
				closeEffect: 'elastic',
				closeSpeed : 150,
				padding    : 5,
				closeClick : true,
				modal      : true,
				showCloseButton : true,
				afterShow : function() {
					$('.fancybox-skin').append('<a title="Close" class="fancybox-item fancybox-close" href="javascript:jQuery.fancybox.close();"></a>');
				}
			});
		});
	}
	
	// Search the o* net list
	var stub = Spine.Model.sub();
	stub.configure("/admin/authAdmin/searchOccupationList", "searchString");
	stub.extend( Spine.Model.Ajax );
		
	//Populate the model with data to transfer
	var modelPopulation = new stub({ 
		searchString: occupationSearchVal
	});
	modelPopulation.save(); //POST the data
	
	stub.bind("ajaxError", function(record, xhr, settings, error) {
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
		alertify.alert("Unable to connect to the server.");
	});
	
	stub.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if (statusCode == "200") {
			plotData(obj.dtoList);
		} else {
			document.getElementById("occupationPlottingDiv").innerHTML = "";
			var plottingDiv = document.getElementById("dataPlottingDiv");
			plottingDiv.innerHTML = "<div align='center'><br /><br /><img src='../img/not-found.png'><br /><p><b>"+occupationSearchVal+"</b> "+obj.statusMsg+"</p></div>";
		}
	});	
}

function plotData(dtoList) {
	document.getElementById("occupationPlottingDiv").innerHTML = "";
	var plottingDiv = document.getElementById("dataPlottingDiv");
	plottingDiv.innerHTML = "";
	var str = "<table width='65%' class='custom_table_onet'><thead><tr class='tab_header'></tr></thead><tbody>";
	
	for (var index = 0; index < dtoList.length; index++) {
		var trColor = "";
		 var titleId = "title_" + index;
		 var descId = "desc_" + index;
		str = str + "<tr id='"+titleId+"' bgcolor="+trColor+"><td class='custom_td_onet' ><input type='radio' class='custom_radio_onet' name='occupation' value='"+dtoList[index].code+"' onclick='showDescription(\""+descId+"\", \""+dtoList.length+"\")'>"+dtoList[index].title+"</td></tr><tr id= '"+descId+"' style='display:none;'><td class='onet_desc_custom_cls' >"+dtoList[index].desc+"</td></tr>";
	}
	str = str + "</tbody></table>";
	plottingDiv.innerHTML = str;
	document.getElementById("goDiv").style.display = "block";
}


function validateOnetData(){
    var textInput = document.getElementById("ocptnSrchTxtId").value;
    //textInput = textInput.replace(/[^A-Za-z]/g, "");
    textInput = textInput.replace(/%/g, "");
    document.getElementById("ocptnSrchTxtId").value = textInput;
}

function showDescription(descId, length) {
	for (var index = 0; index < length; index++) {
		var allDsecId = "desc_" + index;
		document.getElementById(allDsecId).style.display = "none";
	}
	document.getElementById(descId).style.display = "block";
	
}

function closeDiv() {
	if (undefined != $("input[name=occupation]:checked").val()) {
		document.getElementById("ocptnSrchTxtId").value = $("input[name=occupation]:checked").val();
		//populateBeforeSketchSelectedReport('newid');
		sessionStorage.setItem("occupation", document.getElementById("ocptnSrchTxtId").value);
		jQuery.fancybox.close();
	} else {
		alertify.alert("<p>Please select an occupation from the list</p>");
		return false;
	}
}