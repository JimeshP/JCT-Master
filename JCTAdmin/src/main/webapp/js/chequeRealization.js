$(document).ready(function() {
	fetchExistingChequeUsers();
});
/**
 * Function to allow only numeric value
 */
$("#chequeNumInpt").keypress(function (e) {   
    if (e.which != 8 && e.which != 0 && (e.which < 48 || e.which > 57)) {       
       return false;
   }
});

/**
 * function to reset all numeric fields if any non-numeric or special char entered 
 * */
$("#chequeNumInpt").blur(function(){	
	var facilitatorId = $(this).attr('id');
	var val = $("#"+facilitatorId).val().trim();	
	if( isNaN(val)){
		$("#"+facilitatorId).val("");			
	}	
});

/**
 * Method fetches users having check payment based on userName or check number or both
 * 
 */
function fetchChequeDetails() {
	var chequeNum = document.getElementById("chequeNumInpt").value;
	var userName = document.getElementById("userNameInpt").value;
	var fetchType = "";
	if(chequeNum=="" && userName==""){
		alertify.alert("Please enter User Name or Check Number");
		return false;
	} else if (chequeNum != "" && userName == ""){		
		fetchType = "BY_CHECK";		
	} else if(chequeNum == "" && userName != ""){
		fetchType = "BY_USERNAME";		
	} else {	//	 using check no & user Name
		fetchType = "BY_ALL";		
	}
	
	var newChequeDetails = Spine.Model.sub();
	newChequeDetails.configure("/admin/manageuser/fetchExistingChequeUsers", "fetchType", "chequeNum","userName");
	newChequeDetails.extend( Spine.Model.Ajax );
	var modelPopulator = new newChequeDetails({
		fetchType: fetchType,
		chequeNum: chequeNum,
		userName: userName
	});
	modelPopulator.save();
	newChequeDetails.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	newChequeDetails.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if(statusCode==375){
			alertify.alert("No result found corresponding to the given input");
			document.getElementById("chequeDetails").innerHTML = "<div align='center'><br /><br /><br /><img src='../img/no-record.png'><br /><div class='textStyleNoExist'>No Existing Users having cheque payment</div><br/></div>";
			document.getElementById("existing_list_Id").style.display = "none";
			document.getElementById("chequeNumInpt").value = "";
			document.getElementById("userNameInpt").value = "";
		} else if(statusCode == 200) {
			//populateChequeDetails(obj.existingUsers);
			populateChequeDetails(obj.chkDetailsString);
		} else{
			alertify.alert("Something went wrong.");
		}
	});
  }
/**
 * Method populates the check realization details
 * @param existingUsers 
 */
function populateChequeDetails(existingUsers){
	if(existingUsers.length>0){
	document.getElementById("existing_list_Id").style.display = "block";
	tableContent = "";

	var tableDiv = document.getElementById("chequeDetails");
	var headingString = "<table width='100%' border='1' bordercolor='#78C0D3' id='myTable' class='tablesorter'><thead class='tab_header'>" +
			"<tr><th>SL. No.</th><th>User Name</th><th>Check Number</th><th>Date Of Payment</th><th>Honored / Dishonored</th></tr></thead><tbody>";
	var counter = 1;
	tableContent = "";
	var outerStr = existingUsers.split("#");
	for(var i = 0; i< outerStr.length-1 ; i++ ){
		var singleRow = outerStr[i].split("$$");
		var commonStrObj = singleRow[0].split("~");
		var chkNo = commonStrObj[0];
		var paymentDate = commonStrObj[1];
		var transId = commonStrObj[2];	
		
		var trColor = "";
		if(i % 2 == 0) {
			trColor = "#D2EAF0";
		} else {
			trColor = "#F1F1F1";
		}
		
		tableContent += "<tr bgcolor='"+trColor+"'><td align='center'>"+counter+"</td><td align='left'><ol style='list-style-type: square;margin-top: 1%;'>";
		var restStrObj = singleRow[1].split("!");
		for(var j = 0; j < restStrObj.length-1; j++ ){			
			tableContent += "<li>" + restStrObj[j].split("~")[0] + "</li>";		//	each email id
		}
		tableContent += "</ol></td><td align='center'>"+chkNo+"</td><td align='center'>"+paymentDate+"</td>" +
				"<td align='center'><a href='#' onclick='honoredCheck(\""+transId+"\")'><b>Honored</b></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" +
					 "<a href='#' onclick='dishonoredCheck(\""+transId+"\")'><b>Dishonored</b></a></td>";		
		tableContent += "</tr>";		
		counter++;	
	}
	var mainContent = headingString + tableContent+"</tbody></table>";
	tableDiv.innerHTML = mainContent;
	$("table").tablesorter();
	}
	else{
		document.getElementById("chequeDetails").innerHTML = "<div align='center'><br /><br /><br /><img src='../img/no-record.png'><br /><div class='textStyleNoExist'>No Record Found.</div><br/></div>";
		document.getElementById("existing_list_Id").style.display = "none";
	}
}

/**
 * Method fetches all existing users having cheque payment
 */
function fetchExistingChequeUsers(){
	var chequeProf = Spine.Model.sub();
	chequeProf.configure("/admin/manageuser/fetchExistingChequeUsers", "none", "fetchType");
	chequeProf.extend( Spine.Model.Ajax );
	var modelPopulator = new chequeProf({  
		none: "",
		fetchType: ""
	});
	modelPopulator.save();
	chequeProf.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	chequeProf.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;		
		if(statusCode == 200) {
			populateChequeDetails(obj.chkDetailsString);
		}
	});
}
/**
 * function to honored check
 * @param tranId
 **/
function honoredCheck(tranId) {
	alertify.set({ buttonReverse: true });
	alertify.confirm("Are you sure you want to honored the check ?", function (e) {
		if (e) {
			// Blur the screen
			$(".loader_bg").fadeIn();
			$(".loader").fadeIn();
			
			var userDelete = Spine.Model.sub();
			userDelete.configure("/admin/manageuser/honoredCheck", "tranId", "createdBy");
			userDelete.extend( Spine.Model.Ajax );
			//Populate the model with data to transfer
			var modelPopulator = new userDelete({
				tranId: tranId,
				createdBy: sessionStorage.getItem("jctEmail")		
			});
			modelPopulator.save(); //POST
			userDelete.bind("ajaxError", function(record, xhr, settings, error) {
				alertify.alert("Unable to connect to the server.");
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			});
			userDelete.bind("ajaxSuccess", function(record, xhr, settings, error) {
				var jsonStr = JSON.stringify(xhr);
				var obj = jQuery.parseJSON(jsonStr);
				var statusCode = obj.statusCode;
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
				if (statusCode == 200) {					
					alertify.alert("Check Realization is done successfully.");
					fetchExistingChequeUsers();
				} 
			});
		}
	});
}
/**
 * function to dishonored check
 * @param tranId
 **/
function dishonoredCheck(tranId){
	alertify.set({ buttonReverse: true });
	alertify.confirm("Are you sure you want to dishonored the check ?", function (e) {
		if (e) {
			// Blur the screen
			$(".loader_bg").fadeIn();
			$(".loader").fadeIn();			
			var userDelete = Spine.Model.sub();
			userDelete.configure("/admin/manageuser/dishonoredCheck", "tranId","createdBy");
			userDelete.extend( Spine.Model.Ajax );
			//Populate the model with data to transfer
			var modelPopulator = new userDelete({  
				tranId: tranId,
				createdBy: sessionStorage.getItem("jctEmail")
			});
			modelPopulator.save(); //POST
			userDelete.bind("ajaxError", function(record, xhr, settings, error) {
				alertify.alert("Unable to connect to the server.");
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			});
			userDelete.bind("ajaxSuccess", function(record, xhr, settings, error) {
				var jsonStr = JSON.stringify(xhr);
				var obj = jQuery.parseJSON(jsonStr);
				var statusCode = obj.statusCode;
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
				if (statusCode == 200) {					
					alertify.alert("Check Realization is done successfully.");
					fetchExistingChequeUsers();
				} 
			});
		}
	});
}