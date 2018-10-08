var emailIdCSVString = "";
var totalUser = 0;
/**
 * Function calls when the page is loaded 
 * to populate the existing user
 */
$(document).ready(function() {
	sessionStorage.removeItem("previous_user_email");	
	disableAllField();
	var userType = sessionStorage.getItem("userType");
	setTimeout(function(){		
		if(userType == 1){
			document.getElementById("selectUserType").selectedIndex = '1';
		} else if (userType == 2){
			document.getElementById("selectUserType").selectedIndex = '2';
		} 
	}, 200);
	changeUserType(userType);
});


/**
 * Function to allow only numeric value
 */
$("#facilitatorId, #noOfUser1, #chequeno, #noOfUser2, #chequeno").keypress(function (e) {  	   
    if (e.which != 8 && e.which != 0 && (e.which < 48 || e.which > 57)) {       
       return false;
   }
});
/**
 * Function to allow only numeric and . for Amount fields
 */
$("#totalAmount, #cashAmount").keypress(function (e) {  	   
    if (e.which != 8 && e.which != 0 && (e.which < 48 || e.which > 57) && e.which != 46) {
       return false;
   }
});
/**
 * function to reset all numeric fields if any non-numeric or special char entered 
 * */
$("#facilitatorId, #totalAmount, #cashAmount, #noOfUser1, #chequeno, #noOfUser2, #chequeno").blur(function(){	
	var fieldId = $(this).attr('id');
	var val = $("#"+fieldId).val().trim();	
	if( isNaN(val) || val == 0){
		$("#"+fieldId).val("");			
	}	
});
 
/**
 * function to toogle payment mode input fields 
 **/
function tooglePaymentMode() {
	var paymentMode = document.getElementById("paymentMode").options[document.getElementById("paymentMode").selectedIndex].value;
	$(".inputRightSide").hide();	//	all the payment related details input
	if (paymentMode == "CHK") {
		document.getElementById("cashAmount").value = "";
		document.getElementById("paymentDtCash").value = "";
		document.getElementById("nosuser1").style.display = "block";
		document.getElementById("paymentDt2").style.display = "block";
		document.getElementById("chkNo2").style.display = "block";
		$("#cashAmountDiv").hide();

	} else if(paymentMode == "CMP") {
		document.getElementById("cashAmount").value = "";
		document.getElementById("paymentDtCash").value = "";
		document.getElementById("totalAmount").value = "";
		document.getElementById("chequeno").value = "";
		document.getElementById("datepicker").value = "";
		document.getElementById("paymentDt1").style.display = "block";		
	} else if (paymentMode == "CASH"){	
		document.getElementById("totalAmount").value = "";
		document.getElementById("chequeno").value = "";
		document.getElementById("datepicker").value = "";
		$("#cashAmountDiv").show().val('');
		$("#paymentDtCashDiv").show().val('');
	} else {
		document.getElementById("cashAmount").value = "";
		document.getElementById("paymentDtCash").value = "";
		document.getElementById("totalAmount").value = "";
		document.getElementById("chequeno").value = "";
		document.getElementById("datepicker").value = "";
		document.getElementById("paymentDt1").style.display = "block";
	}
}

/**
 * function to  save the renewed / subscribed facilitator's details
 **/
function subscribeOrRenewUser() {
	var usrType = document.getElementById("selectUserType").options[document.getElementById("selectUserType").selectedIndex].value;
	if(usrType == 0) {	//		when no user type selected
		alertify.alert("Please select an user type");		
	} else {		
		if (usrType == "2") { //FACILITATOR
			var facilitatorId = $("#facilitatorId").val().trim();
			var paymentMode = document.getElementById("paymentMode").options[document.getElementById("paymentMode").selectedIndex].value;
			var inputRenewSub = document.getElementById("inputRenewSub").options[document.getElementById("inputRenewSub").selectedIndex].value;
			var noOfUser = $("#noOfUserFac").val().trim();
			if ( facilitatorId == ("" || 0 ) ){
				alertify.alert("Please enter Facilitator ID");
				return false;
			} /*else if( isNaN(facilitatorId) ){
				alertify.alert("Facilitator Id should contain numeric value only");
				return false;
			}*/
			if ( inputRenewSub == "" ){
				alertify.alert("Please select Subscribe / Renew");
				return false;
			}
			if( noOfUser == "" || isNaN(noOfUser)){
				alertify.alert("Please enter number of users the admin has paid for");
				return false;		
			}	
			if( paymentMode == "" ){
				alertify.alert("Please select payment mode");
				return false;
			} else {
				if ( paymentMode == "CHK" ){
					var allFieldValidation = validateAllMandatoryFields ("CHK");
					if (allFieldValidation == true) {
						$(".loader_bg").fadeIn();
						$(".loader").fadeIn();
						saveFacilitatorChkPayment(); // Check Payment				
					}				
				} else if( paymentMode == "CMP" ){	//	FREE	
						$(".loader_bg").fadeIn();
						$(".loader").fadeIn();
						saveFacilitatorCmpPayment(); // Free				
				} else if ( paymentMode == "CASH" ){			
					var allFieldValidation = validateAllMandatoryFields ("CASH");
					if (allFieldValidation == true) {
						$(".loader_bg").fadeIn();
						$(".loader").fadeIn();
						saveFacilitatorCashPayment(); // Cash Payment				
					}	
				}		
			}
		} else { //GENERAL USER
			var paymentMode = document.getElementById("paymentMode").options[document.getElementById("paymentMode").selectedIndex].value;
			if (paymentMode.trim().length > 0) {
				if (paymentMode == "CHK") {					
					var allFieldValidation = validateAllMandatoryFieldsGeneralUser ("CHK");
					if (allFieldValidation == true) {						
						if (document.getElementById("manual").checked) {			 // Manual
							//renewGeneralUserCSVChkPayment("manual"); // Check Payment
							validateGeneralUserPayment("manual");
						} else {													 // CSV
							//renewGeneralUserCSVChkPayment("csv");
							validateGeneralUserPayment("csv");
						}
					}
				} else if (paymentMode == "CMP") {
					var allFieldValidation = validateAllMandatoryFieldsGeneralUser ("CMP");
					if (allFieldValidation == true) {
					/*	$(".loader_bg").fadeIn();
						$(".loader").fadeIn();*/
						/*if (document.getElementById("manual").checked) {			 // Manual
							//renewGeneralUserManualCmpPayment(prepareEmailIdString()); // Complementary Payment
							renewGeneralUserCmpPayment("manual");
						} else {													 // CSV
							renewGeneralUserCmpPayment("csv");
						}*/
						if (document.getElementById("manual").checked) {			 // Manual
							//renewGeneralUserCSVChkPayment("manual"); // Check Payment
							validateGeneralUserPayment("manual");
						} else {													 // CSV
							//renewGeneralUserCSVChkPayment("csv");
							validateGeneralUserPayment("csv");
						}
					}
				} else if (paymentMode == "CASH") {
					var allFieldValidation = validateAllMandatoryFieldsGeneralUser ("CASH");
					if (allFieldValidation == true) {
						/*$(".loader_bg").fadeIn();
						$(".loader").fadeIn();
						if (document.getElementById("manual").checked) {			 // Manual
							renewGeneralUserManualCashPayment(prepareEmailIdString()); // Complementary Payment
						} else {													 // CSV
							renewGeneralUserCSVCashPayment();
						}	*/
						if (document.getElementById("manual").checked) {			 // Manual
							//renewGeneralUserCSVChkPayment("manual"); // Check Payment
							validateGeneralUserPayment("manual");
						} else {													 // CSV
							//renewGeneralUserCSVChkPayment("csv");
							validateGeneralUserPayment("csv");
						}
					}
				}
			} else {
				alertify.alert("Please select a payment mode");
			}
		}
	
	}		
}

function bankDetailsValidation(bankDetails){			
	var pattern = /^[A-Za-z0-9 \.\,]+$/;						//  special char aren't allowed
	if(!pattern.test(bankDetails)){
	  alertify.alert('Bank Details should not contains any special charecter.');
	  return false;
	}
}
function chequeNumberValidation(chequeNos){					
	if( chequeNos.length < 6){
		alertify.alert('Length of Check Number should be atleast 6.');
		return false;
	}
}
/**
 * function to to validate all mandatory input fields
 * @param paymentMode
 **/
function validateAllMandatoryFields (paymentMode) {
	//var numberOfUsers = document.getElementById("noOfUser1").value.trim();
	if (paymentMode == "CHK") {
		var paymentDate = document.getElementById("datepicker").value.trim();		
		var chequeNos = document.getElementById("chequeno").value.trim();	
		var chkAmount = document.getElementById("totalAmount").value.trim();		
		if((chkAmount == ( 0 || '') )){
			alertify.alert("Please enter check amount");
			return false;
		}
		if (chequeNos.length == 0) {
			alertify.alert("Please enter check details");
			return false;
		} else {
			if( chequeNumberValidation(chequeNos) == false )
				return false;		
		}
		if (paymentDate.length == 0) {
			alertify.alert("Please select a payment date");
			return false;
		} /*else if (numberOfUsers.length == 0) {
			alertify.alert("Please enter number of users the admin has paid for");
			return false;
		} */	
		return true;
	} else if (paymentMode == "CASH") {
		var cashAmount = $("#cashAmount").val().trim();
		var paymentdate = $("#paymentDtCash").val().trim();
		if(cashAmount == ( 0 || '')){
			alertify.alert("Please enter cash amount");
			return false;
		} else if (paymentdate.length == 0) {
			alertify.alert("Please select a payment date");
			return false;			
		} /*else if (numberOfUsers.length == 0) {
			alertify.alert("Please enter number of users the admin has paid for");
			return false;
		} 	*/
		return true;
	}
}
/**
 * Function saves facilitator through CHECK payment
 */
function saveFacilitatorChkPayment() {
	var facilitatorID = $("#facilitatorId").val().trim();
	var paymentDate = document.getElementById("datepicker").value;
	var numberOfUsers = document.getElementById("noOfUserFac").value;
	
	var chequeNos = document.getElementById("chequeno").value;
	var totalAmountToBePaid = document.getElementById("totalAmount").value;
	var inputRenewSub = document.getElementById("inputRenewSub").options[document.getElementById("inputRenewSub").selectedIndex].value;
	var model = Spine.Model.sub();
	model.configure("/admin/manageuser/renewSubscribe", "createdBy","facilitatorID","paymentDate",
			"numberOfUsers", "chequeNos", "totalAmountToBePaid", "paymentType","inputRenewSub");
	model.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new model({  		
		createdBy: sessionStorage.getItem("jctEmail"),
		facilitatorID:facilitatorID,
		//bankDetails: bankDetails,		
		paymentDate: paymentDate,
		numberOfUsers: numberOfUsers,
		chequeNos: chequeNos,
		totalAmountToBePaid: totalAmountToBePaid,
		paymentType: "CHECK",
		inputRenewSub: inputRenewSub
	});
	modelPopulator.save(); //POST
	model.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	});
	model.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		alertify.alert(obj.message);
		$(".loader_bg").fadeOut();
	    $(".loader").hide();	
	    if (obj.statusCode != 123) {		//	success
	    	 if (obj.statusCode == 200) {
	    		 clearAllFacilitatorData();
    		    fetchRecentActivity(facilitatorID,numberOfUsers,"CHECK");
	 	    } else if (obj.statusCode == 201) {	// no such user
	    		$("#updatedFaciTable").hide();	    		
	    	}
	    	
	    }
	});
}

/**
 * Function saves facilitator through CASH payment
 */
function saveFacilitatorCashPayment() {
	var facilitatorID = $("#facilitatorId").val().trim();
	//var bankDetails = document.getElementById("bankDetailsVal").value;	
	var paymentDate = document.getElementById("paymentDtCash").value;
	var numberOfUsers = document.getElementById("noOfUserFac").value;
	var totalAmountToBePaid = document.getElementById("cashAmount").value;
	var inputRenewSub = document.getElementById("inputRenewSub").options[document.getElementById("inputRenewSub").selectedIndex].value;
	
	var model = Spine.Model.sub();
	model.configure("/admin/manageuser/renewSubscribe", "createdBy","facilitatorID","paymentDate",
			"numberOfUsers", "totalAmountToBePaid", "paymentType","inputRenewSub");
	model.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new model({  		
		createdBy: sessionStorage.getItem("jctEmail"),
		facilitatorID:facilitatorID,		
		paymentDate: paymentDate,
		numberOfUsers: numberOfUsers,
		totalAmountToBePaid: totalAmountToBePaid,
		paymentType: "CASH",
		inputRenewSub: inputRenewSub
	});
	modelPopulator.save(); //POST
	model.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	});
	model.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		alertify.alert(obj.message);
		$(".loader_bg").fadeOut();
	    $(".loader").hide();	    
	    if (obj.statusCode != 123) {    	//	success
	    	 if (obj.statusCode == 200) {
	    		 clearAllFacilitatorData();
			    fetchRecentActivity(facilitatorID,numberOfUsers,"CASH");
	    	 } else if (obj.statusCode == 201) {	// no such user
	    		 $("#updatedFaciTable").hide();
		    }
	    }
	});
}
/**
 * Function saves facilitator through FREE payment
 */
function saveFacilitatorCmpPayment() {	
	//var numberOfUsers = document.getElementById("noOfUser1").value;
	var numberOfUsers = $("#noOfUserFac").val();
	var facilitatorID = document.getElementById("facilitatorId").value;
	var inputRenewSub = document.getElementById("inputRenewSub").options[document.getElementById("inputRenewSub").selectedIndex].value;
		
	var model = Spine.Model.sub();
	model.configure("/admin/manageuser/renewSubscribe", "paymentDate","facilitatorID", "numberOfUsers",
			"chequeNos", "totalAmountToBePaid","paymentType","createdBy","inputRenewSub");	
	model.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new model({  
		facilitatorID: facilitatorID,
		paymentDate: "",
		numberOfUsers: numberOfUsers,
		chequeNos: "",
		totalAmountToBePaid: 0,
		createdBy: sessionStorage.getItem("jctEmail"),
		paymentType: "FREE",
		inputRenewSub: inputRenewSub
	});
	
	modelPopulator.save(); //POST	
	model.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	});
	model.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		alertify.alert(obj.message);
		$(".loader_bg").fadeOut();
	    $(".loader").hide();	 
	    if (obj.statusCode != 123) {		//	success
	    	if (obj.statusCode == 200) {
	    		clearAllFacilitatorData();
			    fetchRecentActivity(facilitatorID,numberOfUsers,"FREE", inputRenewSub);
	    	} else if (obj.statusCode == 201) {	// no such user
	    		$("#updatedFaciTable").hide();    		
	    	}
	    }
	});
}
/**
 * function to fetch facilitator email Id using facilitatorId
 * and call plotUpdatedData() to popolate data table 
 * @param facilitatorID
 * @param numberOfUsers
 * @param paymentType
 * */
function fetchRecentActivity(facilitatorID,numberOfUsers,paymentType, inputRenewSub) {
	// get facilitator email id 
	var model = Spine.Model.sub();
	model.configure("/admin/manageuser/getFacilitatorEmail", "facilitatorID");
	model.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new model({
		facilitatorID: facilitatorID
	});
	
	modelPopulator.save(); //POST	
	model.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	});
	model.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		if(obj.statusCode == "200") {	//	success			
			$(".loader_bg").fadeOut();
			$(".loader").hide();
			plotUpdatedDataTable(facilitatorID,numberOfUsers,paymentType,obj.email, inputRenewSub);
		}	   
	});
	
}
/**
 * function to populate updated facilitator details table
 * @param facilitatorID
 * @param numberOfUsers
 * @param paymentType
 * @param facilitatorEmail
 **/
function plotUpdatedDataTable(facilitatorID,numberOfUsers,paymentType,facilitatorEmail, inputRenewSub) {
	document.getElementById("existinglistId").innerHTML = "Subscribed/Renewed Details";
	var counter = 1;	
	paymentType = paymentType == "FREE" ? "Free" : (paymentType == "CHECK" ? "Check" : (paymentType == "CASH" ? "Cash" : "e-commerce"));
	if(inputRenewSub == "AD"){
		inputRenewSub = "(Purchased for Subscription)";
	} else {
		inputRenewSub = "(Purchased for Renewal)";
	}
	table = "<table width='94%' border='1' bordercolor='#78C0D3' id='myTable' align='center' class='tablesorter'><thead width='94%' class='tab_header'><tr><th>SL. No.</th><th>Facilitator Id</th><th>Facilitator Name</th><th>Subscribed/Renewed User</th><th width='22%'>Payment Mode</th></tr></thead>";
	table += "<tbody width='94%'><tr bgcolor='#D2EAF0' align='center'><td>"+counter+"</td><td>"+facilitatorID+"</td><td>"+facilitatorEmail+"</td><td>"+numberOfUsers+" "+inputRenewSub+"</td><td>"+paymentType+"</td></tr></tbody></table>";
	
	document.getElementById("updatedFaciTable").style.display="block";
	document.getElementById("mainPopulationDiv").innerHTML = table;
	$("#updatedFaciTable").show();
}
/**
 * function to reset all input fields
 * */
function clearAllFacilitatorData(){	
	document.getElementById("facilitatorId").value = "";
	document.getElementById("noOfUserFac").value = "";
	document.getElementById("totalAmount").value = "";
	document.getElementById("chequeno").innerHTML = "";
	document.getElementById("datepicker").innerHTML = "";
	document.getElementById("inputRenewSub").selectedIndex = "0";
	document.getElementById("paymentMode").selectedIndex = "0";
	document.getElementById("paymentMode").selectedIndex = "0";
	$("#cashAmount, #paymentDtCash").val('');
	$('.inputRightSide').hide();
	$('.inputRightSide input').val('');
	document.getElementById("paymentDt1").style.display = "block";
}

function clearAllIndividualData(){	
	
	document.getElementById("inputNoOFUser").selectedIndex = "0";
	document.getElementById("manualEntryArea1").style.display = "none";
	document.getElementById("paymentMode").selectedIndex = "0";
	document.getElementById("totalAmount").value = "";    
	document.getElementById("chequeno").innerHTML = "";
	document.getElementById("datepicker").innerHTML = "";
	$("#cashAmount, #paymentDtCash").val('');
	$('#manual').prop('checked',true);
	
	$('.inputRightSide').hide();
	$('.inputRightSide input').val('');
	document.getElementById("paymentDt1").style.display = "block";	
}

/**
 * function to add datePicker functionality to selected input 
 **/
$(document).ready(function () {                                               
	 $.datepicker.setDefaults( 
	    {showOn: 'button', buttonImageOnly: true, buttonImage: '../img/datepicker-icon.png',dateFormat: "mm/dd/yy"});  
	 $('#datepicker, #paymentDtCash').datepicker();
	});
$(function() {
	$(".datepickerPastDate").datepicker("option", "minDate", 0);
 });


function changeUserType(usrType) {
	document.getElementById("updatedFaciTable").style.display="none";	
	//var usrType = document.getElementById("selectUserType").options[document.getElementById("selectUserType").selectedIndex].value;
	if(usrType == "0") { //NOTHING SELECTED
		//document.getElementById("paymentMode").setAttribute("disabled", "true");
		//document.getElementById("inputRenewSub").setAttribute("disabled", "true");	
	} else if (usrType == "2") { //FACILITATOR
		document.getElementById("paymentMode").removeAttribute("disabled", "true");
		document.getElementById("inputRenewSub").removeAttribute("disabled", "true");		
		document.getElementById("renewLabelId").innerHTML  = "Subscribe / Renew:";
		document.getElementById("inputRenewSub").selectedIndex = "0";
		document.getElementById("inputRenewSub").removeAttribute("disabled", "true");
		document.getElementById("radioButtonDivIndividual").style.display = "none";
		document.getElementById("noUserIndividual").style.display = "none";
		document.getElementById("noUserIndividual1").style.display = "block";		
		document.getElementById("csvFileArea1").style.display = "none";
		$("#manualEntryArea1").hide();			
		document.getElementById("subscribtionDivFacilitator").style.display = "block";
		document.getElementById("noOfUserFacilitator").style.display = "block";		
	} else {				//INDIVIDUAL				
		document.getElementById("paymentMode").removeAttribute("disabled", "true");
		document.getElementById("inputRenewSub").removeAttribute("disabled", "true");		
		document.getElementById("renewLabelId").innerHTML  = "Renew:";
		document.getElementById("inputRenewSub").selectedIndex = "1";
		document.getElementById("inputRenewSub").setAttribute("disabled", "true");
		document.getElementById("radioButtonDivIndividual").style.display = "block";
		document.getElementById("subscribtionDivFacilitator").style.display = "none";
		document.getElementById("noOfUserFacilitator").style.display = "none";		
		document.getElementById("noUserIndividual").style.display = "block";
		document.getElementById("noUserIndividual1").style.display = "none";
		document.getElementById("inputNoOFUser").selectedIndex = "0";
		$('#manual').prop('checked',true);
	}
}


function changeRenewSubscribe(){
	
}
function toogleSelectionDiv(selection) {
	if (selection == "M") {					
		document.getElementById("inputNoOFUser").selectedIndex = "0";
		document.getElementById("noUserIndividual").style.display = "block";
		document.getElementById("noUserIndividual1").style.display = "none";
		document.getElementById("csvFileArea1").style.display = "none";				
	} else {		
		document.getElementById("noUserIndividual").style.display = "none";
		document.getElementById("noUserIndividual1").style.display = "block";
		document.getElementById("csvFileArea1").style.display = "block";
		$("#manualEntryArea1").hide();
	}
}

/**
 * Function to populate the manual entry fields
 * while user click on mnaual entry
 */
 function populateManualEntryFields() {
	$("#manualEntryArea1").show();

	divCount = $('#manualEntryArea1').children('.single_form_item:visible').length;
	var jsonObj = [];
	var counter = 0;
	for (var i = 1; i <= divCount; i += 1) {
	var unitJ = {};	
	var value = document.getElementById('manualEmailId_'+i).value;
	var expDate = document.getElementById('expdateId_'+i).value;
	unitJ["ElementValue"+i] = value;
	unitJ["ExpDate"+i] = expDate;
	jsonObj[counter++] = unitJ;
	}
	console.log(jsonObj);
	sessionStorage.setItem("previous_user_email",JSON.stringify(jsonObj));		
	var manualEntryNumber = document.getElementById("inputNoOFUser").options[document.getElementById("inputNoOFUser").selectedIndex].value;
	
	var plotPlement = document.getElementById("manualEntryArea1");
	document.getElementById("manualEntryArea1").innerHTML = "";
	document.getElementById("manualEntryArea1").style.display = "block";
	var string = "";
	var idCountIndex = 1;
	if (manualEntryNumber == "") {
		alertify.alert("Please select numbers.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	} else {
		for (var index = 0; index < manualEntryNumber; index++) {
			var textFieldId = "manualEmailId_"+idCountIndex;
			var expDateId = "expdateId_"+idCountIndex;			
			
			string = string + "<div class='single_form_item'><div class='col-sm-3 '><label for='inputQtnNo' class='col-sm-12 control-label text_label'>User Name "+idCountIndex+":</label></div>" +
			"<div class='col-sm-3'><input type='text' class='form-control-general' maxlength='50' value='' placeholder='Enter Email Address' id='"+textFieldId+"'></div>" +
			"<div class='col-sm-3 '><label for='inputQtnNo' class='col-sm-12 control-label text_label'>Expiry Duration:</label></div>" +
			"<div class='col-sm-3'><select class='form-control-general font_ipad_two' id='"+expDateId+"'>" +
				"<option class='form-control-general' value=''>Select Expiry Duration</option>" +
				"<option class='form-control-general' value='1'>1 Month</option>" +
				"<option class='form-control-general' value='2'>2 Month</option>" +
				"<option class='form-control-general' value='3'>3 Month</option>" +
				"<option class='form-control-general' value='4'>4 Month</option>" +
				"<option class='form-control-general' value='5'>5 Month</option>" +
				"<option class='form-control-general' value='6'>6 Month</option>" +
				"<option class='form-control-general' value='7'>7 Month</option>" +
				"<option class='form-control-general' value='8'>8 Month</option>" +
				"<option class='form-control-general' value='9'>9 Month</option>" +
				"<option class='form-control-general' value='10'>10 Month</option>" +
				"<option class='form-control-general' value='11'>11 Month</option>" +
				"<option class='form-control-general' value='12'>12 Month</option>" +
				"</select></div>" +
			"<div class='clearfix'></div></div>";				
			idCountIndex = idCountIndex + 1;
		}
		plotPlement.innerHTML = string;
	}
	
	if(sessionStorage.getItem("previous_user_email") != null){
		var jsonObj = JSON.parse(sessionStorage.getItem("previous_user_email"));
		var count = jsonObj.length;
			for (var j = 0; j<count; j++){	
			var value = jsonObj[j]["ElementValue"+(j+1)];
			var expDate = jsonObj[j]["ExpDate"+(j+1)];
			var div = document.getElementById('manualEmailId_'+(j+1));
			    if (div) {
				document.getElementById('manualEmailId_'+(j+1)).value = value;
				document.getElementById('expdateId_'+(j+1)).value = expDate;
				}			
			}		
		}
}
 
function disableAllField() {
	document.getElementById("selectUserType").selectedIndex = "0";
	document.getElementById("paymentMode").setAttribute("disabled", "true");
	document.getElementById("inputRenewSub").setAttribute("disabled", "true");
	
}

function validateAllMandatoryFieldsGeneralUser (paymentMode) {		
	if (paymentMode == "CHK") {
		var paymentDate = document.getElementById("datepicker").value.trim();
		var chequeNos = document.getElementById("chequeno").value.trim();	
		var totalAmount = document.getElementById("totalAmount").value.trim();	
	
		if ((totalAmount.length == 0) || (totalAmount == '') || isNaN(totalAmount)){
			alertify.alert("Please enter check amount.");
			return false;			
		}
		if (chequeNos.length == 0) {					//		check number
			alertify.alert("Please enter check number");
			return false;
		} else {			
			if( chequeNumberValidation(chequeNos) == false )
				return false;			
		}
		if (paymentDate.length == 0) {					//		payment date
			alertify.alert("Please select a payment date");
			return false;
		}
		
		if (document.getElementById("manual").checked) {
			var nosOfUser = document.getElementById("inputNoOFUser").options[document.getElementById("inputNoOFUser").selectedIndex].value;
			if (nosOfUser.trim().length == 0) {
				alertify.alert("Please select number of users");
				return false;
			}
		}
		
		return true;
	} else if (paymentMode == "CMP") {	//	Free
		var usrType = document.getElementById("selectUserType").options[document.getElementById("selectUserType").selectedIndex].value;
		if (usrType.trim().length == 0) {
				alertify.alert("Please select an user type");
				return false;
		}
		if (document.getElementById("manual").checked) {
			var nosOfUser = document.getElementById("inputNoOFUser").options[document.getElementById("inputNoOFUser").selectedIndex].value;
			if (nosOfUser.trim().length == 0) {
				alertify.alert("Please select number of users");
				return false;
			}
		}
		return true;
	} else if (paymentMode == "CASH") {	//	Free
		var usrType = document.getElementById("selectUserType").options[document.getElementById("selectUserType").selectedIndex].value;		
		var totalAmount = document.getElementById("cashAmount").value.trim();
		var paymentDate = document.getElementById("paymentDtCash").value.trim();
		if (usrType.trim().length == 0) {
				alertify.alert("Please select an user type");
				return false;
		} else if ((totalAmount.length == 0 )||( totalAmount == 0)){
			alertify.alert("Please enter cash amount");
			return false;
		} else if ((paymentDate.length == 0)||( paymentDate == 0)){
			alertify.alert("Please enter payment date");
			return false;
		}
		if (document.getElementById("manual").checked) {
			var nosOfUser = document.getElementById("inputNoOFUser").options[document.getElementById("inputNoOFUser").selectedIndex].value;
			if (nosOfUser.trim().length == 0) {
				alertify.alert("Please select number of users");
				return false;
			}
		}
		return true;
	}
}

/**
 * Function prepares the email id string seperated by ~
 * @returns {String}
 */
function prepareEmailIdString() {
	emailIdCSVString = "";
	var numberOfUsers = document.getElementById("inputNoOFUser").options[document.getElementById("inputNoOFUser").selectedIndex].value;
	var id = 1;
	for (var index=0; index<numberOfUsers; index++) {
		var emailString = document.getElementById("manualEmailId_"+id).value;
		//var expDuration = document.getElementById("expdateId_"+id).value;
		var expDuration = document.getElementById("expdateId_"+id).options[document.getElementById("expdateId_"+id).selectedIndex].value;
		if (validateEmailId(emailString)) {
			if (emailString.length > 5) {
				//if(id != numberOfUsers) {
					emailIdCSVString = emailIdCSVString + emailString + "#`#" +expDuration + "~";
				//} else {
				//	emailIdCSVString = emailIdCSVString + emailString + "#`#" +expDuration;
				//}
				id = id + 1;
			}
		} else {
			//alert('Please provide a valid email address');
		    return false;
		}
		
	}
	
	return emailIdCSVString;
}



/**
 *  Function validates the email id.
 * @param emailId
 */
function validateEmailId(emailId) {	
	var filter = /^([a-zA-Z0-9_\.\-])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})+$/;
    if (!filter.test(emailId)) {
    	//alert('Please provide a valid email address');
    	return false;
    }
    return true;
}
/**
 * Function saves payment details for general user 
 * creation along with the mail ids
 */
/*function renewGeneralUserManualChkPayment(emailIdString) {
	if(validateEmailAndExpiryDuratn()) {
		var paymentDate = document.getElementById("datepicker").value;
		var numberOfUsers = document.getElementById("inputNoOFUser").options[document.getElementById("inputNoOFUser").selectedIndex].value;
		var chequeNos = document.getElementById("chequeno").value;
		var totalAmountToBePaid = document.getElementById("totalAmount").value;
				
		var model = Spine.Model.sub();
		model.configure("/admin/manageuser/renewGeneralUserPaymentManual", "emailIdString",
				"paymentDate", "numberOfUsers", "chequeNos", "totalAmountToBePaid",
				"createdBy", "paymentType");
		model.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new model({  
			emailIdString: emailIdString,
			paymentDate: paymentDate,
			numberOfUsers: numberOfUsers,
			chequeNos: chequeNos,
			totalAmountToBePaid: totalAmountToBePaid,
			createdBy: sessionStorage.getItem("jctEmail"),
			paymentType: "CHECK"
		});
		modelPopulator.save(); //POST
		model.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		});
		model.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			displayMessages(obj);
		});
	} else {
		//alertify.alert('Please provide a valid email address.');
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	}
}*/


/*function renewGeneralUserCSVChkPayment() {
	//FOR CSV
	if (emailIdCSVString == "") {
		alertify.alert("Please select the CSV file.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	} else if (emailIdCSVString.trim().length < 5){
		alertify.alert("CSV file is empty.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	} else {
		// First Validate CSV entry
		var model = Spine.Model.sub();
		model.configure("/admin/manageuser/validateCSVEntry", "emailIdString");
		model.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new model({  			
			emailIdString: emailIdCSVString
		});
		modelPopulator.save(); //POST
		model.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		});
		model.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
		    var statusCode = obj.statusCode;
		    if (statusCode == 215) { // All entries are correct
		    	populateSuccessValidation(obj);
		    } else {
		    	populateFailureValidation(obj);
		    }
		});
	}
}*/

function validateGeneralUserPayment(selectionType) {
	if(selectionType == "csv") {  //validation for CSV
		var dispTable = emailIdCSVString.split("~");
		for(var index = 0; index < dispTable.length-1; index++) {
			var nonRepeting = dispTable[index];
			var nonRepeting1 = nonRepeting.split("#`#");
			if ((nonRepeting1[1] <= 0 || nonRepeting1[1] > 12) && nonRepeting1[1] != "") {
				validateEntryForCSVAndManual(emailIdCSVString,selectionType);
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			    return false;
			} else if (isNaN(nonRepeting1[1]) || nonRepeting1[1] == "") {
				alertify.alert("Expiry Duration should be 1-12");
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			    return false;
			}	
		}		
		if (emailIdCSVString == "") {
			alertify.alert("Please select the CSV file.");
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		} else if (emailIdCSVString.trim().length < 5){
			alertify.alert("CSV file is empty.");
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		} else {
			validateEntryForCSVAndManual(emailIdCSVString, selectionType);
		}
	} else {		//validation for Manual		
		if(validateEmailAndExpiryDuratn()){
			//var emailIdManualString = usrType == "1"?  ""   : prepareEmailIdString();
			var emailIdManualString;
			if(usrType == 1){
				var emailId = $("#userEmail").val().trim();
				if(emailId.length < 1){
					alertify.alert("Please provide email id");
					return false;
				}
				emailIdManualString = emailId + "#`#" +expDuration + "~";	//	email#`#duration~				
			} else {
				emailIdManualString = prepareEmailIdString();			
			}
			validateEntryForCSVAndManual(emailIdManualString, selectionType);
		} 		
	}
}


function validateEntryForCSVAndManual(emailIdCSVString, selectionType) {
	// First Validate CSV entry
	$(".loader_bg").fadeIn();
	$(".loader").fadeIn();
	
	var model = Spine.Model.sub();
	model.configure("/admin/manageuser/validateCSVEntry", "emailIdString", "selectionType");
	model.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new model({  			
		emailIdString: emailIdCSVString,
		selectionType: selectionType
	});
	modelPopulator.save(); //POST
	model.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	});
	model.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
	    var statusCode = obj.statusCode;		   
	    if (statusCode == 215) { // All entries are correct
	    	populateSuccessValidation(obj);
	    } else {
	    	populateFailureValidation(obj);
	    }
	});
}

/*function renewGeneralUserCSVCashPayment() {
	//FOR CSV
	if (emailIdCSVString == "") {
		alertify.alert("Please select the CSV file.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	} else if (emailIdCSVString.trim().length < 5){
		alertify.alert("CSV file is empty.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	} else {
		// First Validate CSV entry
		var model = Spine.Model.sub();
		model.configure("/admin/manageuser/validateCSVEntry", "emailIdString");
		model.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new model({  			
			emailIdString: emailIdCSVString
		});
		modelPopulator.save(); //POST
		model.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		});
		model.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
		    var statusCode = obj.statusCode;
		    if (statusCode == 215) { // All entries are correct
		    	// Prepare message
		    	$('#myModal').modal('show');
		    	document.getElementById("message").innerHTML = obj.statusDesc;
		    	
		    	var msg = "";
		    	msg = msg + "<br /><div align='center' style='width:300px; height:100px; overflow: auto;'>";
		    	msg = msg + "<table width='100%'><tr><th>User(s):</th><th>Expiration Date(s):</th></tr>";
		    	for (var index=0; index<obj.validEmailIds.length; index++) {
					msg = msg + "<tr><td>"+obj.validEmailIds[index]+"</td><td> TEST </td>";
				}
		    	msg = msg + "</table></div>";
		    	document.getElementById("tableData").innerHTML = msg;
		    	$(".loader_bg").fadeOut();
			    $(".loader").hide();
		    } else {
		    	
		    }
		    if (statusCode == 215) { // All entries are correct
		    	populateSuccessValidation(obj);
		    } else {
		    	populateFailureValidation(obj);
		    }
		});
	}
}*/


function populateSuccessValidation(obj) {
	// Prepare message
	$('#myModal').modal('show');
	document.getElementById("message").innerHTML = obj.statusDesc;
	
	var msg = "";
	msg = msg + "<div align='center' class='error_msg_model'><table style='border: 1px solid black; width: 90%;'>";
	
	if (obj.validEmailIds.length > 0) {
		msg = msg + "<tr><td><table width='100%'><tr class='set-backgrnd-color'><td><b>User(s):</b></td><td><b>New Expiration Date(s):</b></td></tr>";
		for (var index=0; index<obj.validEmailIds.length; index++) {
			var res = obj.validEmailIds[index].split("#`~`#");
			var expDt1 = res[1].toString().substring(0,10);
			var expDt2 = res[1].toString().substring(24,res[1].toString().length);
			var finalDate = expDt1+" "+expDt2;
			var expiryDate = "";
			var userExpryDate = "";
			expiryDate = new Date(finalDate).toDateString();
			userExpryDate = dateformat(new Date (expiryDate));
			msg = msg + "<tr><td>"+res[0]+"</td><td> "+userExpryDate+" </td>";
		}
		msg = msg + "</table></td></tr>";
	}
	msg = msg + "</table></div>";
	
	document.getElementById("tableData").innerHTML = msg;
	$(".loader_bg").fadeOut();
    $(".loader").hide();
}

function populateFailureValidation(obj) {
	// Prepare message
	$('#myModal1').modal('show');
	document.getElementById("message1").innerHTML = obj.statusDesc;
	
	var msg = "";
	msg = msg + "<div align='center' class='error_msg_model'><div style='border: 1px solid black; width:auto;height:auto;padding-bottom: 3%;'><table style='width: 100%;margin-top: -3%;'><tr>&nbsp;</tr>";
	if(obj.validEmailIds.length > 0) {
    	msg = msg + "<tr><td><table width='100%'><tr align='center'><td align='center'><b>Valid User(s):</b></th></tr>";
    	for (var index=0; index<obj.validEmailIds.length; index++) {
			msg = msg + "<tr><td align='center'>"+obj.validEmailIds[index]+"</td></tr>";
		}
    	msg = msg + "</table></td></tr>";
	}
	if (obj.invalidEmailIds.length > 0) {
		msg = msg + "<tr><td>&nbsp;</td></tr>";
    	msg = msg + "<tr><td><table width='100%'><tr align='center'><td align='center' colspan='2'><b>Invalid User(s):</b></td></tr>";
    	for (var index=0; index<obj.invalidEmailIds.length; index++) {
    		var res = obj.invalidEmailIds[index].split("#`~`#");
    		msg = msg + "<tr><td align='right' width='50%'>"+res[0]+"&nbsp;&nbsp;&nbsp;&nbsp;</td><td align='left'>"+res[1]+"</td></tr>";
		}
    	msg = msg + "</table></td></tr><tr>&nbsp;</tr>";
	}
	msg = msg + "</table></div></div>";
	document.getElementById("tableData1").innerHTML = msg;
	$(".loader_bg").fadeOut();
    $(".loader").hide();
}



function renewCSVAfterValidation() {
	$('#myModal').modal('hide');
	$(".loader_bg").fadeIn();
	$(".loader").fadeIn();
	var paymentMode = document.getElementById("paymentMode").options[document.getElementById("paymentMode").selectedIndex].value;
	if (paymentMode == "CHK") {
		var paymentDate = document.getElementById("datepicker").value;
		var chequeNos = document.getElementById("chequeno").value;
		var totalAmountToBePaid = document.getElementById("totalAmount").value;
		var model = Spine.Model.sub();
		model.configure("/admin/manageuser/renewGeneralUserCSV", "emailIdString",
				"paymentDate", "numberOfUsers", "chequeNos", "totalAmountToBePaid",
				"createdBy", "paymentType");
		model.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new model({  
			emailIdString: emailIdCSVString.substring(0, emailIdCSVString.length-1),
			paymentDate: paymentDate,
			numberOfUsers: totalUser-1,
			chequeNos: chequeNos,
			totalAmountToBePaid: totalAmountToBePaid,
			createdBy: sessionStorage.getItem("jctEmail"),
			paymentType: "CHECK"
		});
		modelPopulator.save(); //POST
		model.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		});
		model.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			if(obj.statusCode == 200) {
				//alertify.alert("User(s) has been renewed successfully.");
				alertify.alert("Once the check realizatoion is done User(s) will be renewed successfully.");
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			    //emailIdCSVString = "";
			} else {
				alertify.alert("Unable to renew users.");
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			}
		});
	} else if (paymentMode == "CASH") {
		var totalAmount = document.getElementById("cashAmount").value.trim();
		var paymentDate = document.getElementById("paymentDtCash").value.trim();
		var model = Spine.Model.sub();
		model.configure("/admin/manageuser/renewGeneralUserCSV", "emailIdString",
				"paymentDate", "numberOfUsers", "chequeNos", "totalAmountToBePaid",
				"createdBy", "paymentType");
		model.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new model({  
			emailIdString: emailIdCSVString.substring(0, emailIdCSVString.length-1),
			paymentDate: paymentDate,
			numberOfUsers: totalUser-1,
			chequeNos: "",
			totalAmountToBePaid: totalAmount,
			createdBy: sessionStorage.getItem("jctEmail"),
			paymentType: "CASH"
		});
		modelPopulator.save(); //POST
		model.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		});
		model.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			if(obj.statusCode == 200) {
				alertify.alert("User(s) has been renewed successfully.");
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			   // emailIdCSVString = "";
			} else {
				alertify.alert("Unable to renew users.");
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			}
		});
	} else {
		var model = Spine.Model.sub();
		model.configure("/admin/manageuser/renewGeneralUserCSV", "emailIdString",
				"paymentDate", "numberOfUsers", "chequeNos", "totalAmountToBePaid",
				"createdBy", "paymentType");
		model.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new model({  			
			emailIdString: emailIdCSVString.substring(0, emailIdCSVString.length-1),
			paymentDate: "",
			numberOfUsers: totalUser-1,
			chequeNos: "",
			totalAmountToBePaid: 0,
			createdBy: sessionStorage.getItem("jctEmail"),
			paymentType: "FREE"
		});
		modelPopulator.save(); //POST
		model.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		});
		model.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			if(obj.statusCode == 200) {
				alertify.alert("User(s) has been renewed successfully.");
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			   // emailIdCSVString = "";
			} else {
				alertify.alert("Unable to renew users.");
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			}
		});
	
	}
	clearAllIndividualData();
	if(paymentMode != "CHK") {
		fetchUsersForRenew(paymentMode);	
	}	
}

function closeMyModel() {
	$('#myModal').modal('hide');
}


function closeMyModel1() {
	$('#myModal1').modal('hide');
}

/*function renewGeneralUserManualCmpPayment(emailIdString) {
	//if (emailIdString != "") {
	if(validateEmailAndExpiryDuratn()) {		
		var numberOfUsers = document.getElementById("inputNoOFUser").options[document.getElementById("inputNoOFUser").selectedIndex].value;
		var model = Spine.Model.sub();
		model.configure("/admin/manageuser/renewGeneralUserPaymentManual", "emailIdString",
				"paymentDate", "numberOfUsers", "chequeNos", "totalAmountToBePaid",
				"createdBy", "paymentType");
		model.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new model({  			
			emailIdString: emailIdString,
			paymentDate: "",
			numberOfUsers: numberOfUsers,
			chequeNos: "",
			totalAmountToBePaid: 0,
			createdBy: sessionStorage.getItem("jctEmail"),
			paymentType: "FREE"
		});
		modelPopulator.save(); //POST
		model.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		});
		model.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
		    displayMessages(obj);
		});
	} else {
		$(".loader_bg").fadeOut();
		$(".loader").hide();
	}
	//else {
	//	alertify.alert('Please provide a valid email address.');
		//$(".loader_bg").fadeOut();
	   // $(".loader").hide();
	//}
}*/


function validateEmailAndExpiryDuratn() {
	var numberOfUsers = document.getElementById("inputNoOFUser").options[document.getElementById("inputNoOFUser").selectedIndex].value;	
	for (var index=1; index<=numberOfUsers; index++) {
		var emailString = document.getElementById("manualEmailId_"+index).value.trim();
		//var expDuration = document.getElementById("expdateId_"+index).value.trim();
		var expDuration = document.getElementById("expdateId_"+index).options[document.getElementById("expdateId_"+index).selectedIndex].value;
		var filter = /^([a-zA-Z0-9_\.\-])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})+$/;
		if(null == emailString || emailString ==""){
			alertify.alert('Please provide a valid email address.');
			return false;
		} else if(null == expDuration || expDuration == "" || (expDuration.trim().length == 0)) {
			alertify.alert('Please select expiry duration.');
			return false;
		} else if (!filter.test(emailString)) {
			alertify.alert('Please provide a valid email address.');
	    	return false;
	    }/* else if(expDuration <= 0 || expDuration > 12) {
	    	alertify.alert('Please provide expiry duration 1-12.');
			return false;
	    }*/
	}
		
	return true;
}

/*function renewGeneralUserManualCashPayment(emailIdString){
	if(validateEmailAndExpiryDuratn()) {
		var numberOfUsers = document.getElementById("inputNoOFUser").options[document.getElementById("inputNoOFUser").selectedIndex].value;
		var totalAmount = document.getElementById("cashAmount").value.trim();
		var paymentDate = document.getElementById("paymentDtCash").value.trim();
		var model = Spine.Model.sub();
		model.configure("/admin/manageuser/renewGeneralUserPaymentManual", "emailIdString",
				"paymentDate", "numberOfUsers", "chequeNos", "totalAmountToBePaid",
				"createdBy", "paymentType");
		model.extend( Spine.Model.Ajax );
		//Populate the model with data to transfer
		var modelPopulator = new model({  
			emailIdString: emailIdString,
			paymentDate: paymentDate,
			numberOfUsers: numberOfUsers,
			chequeNos: "",
			totalAmountToBePaid: totalAmount,
			createdBy: sessionStorage.getItem("jctEmail"),
			paymentType: "CASH"
		});
		modelPopulator.save(); //POST
		model.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		});
		model.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			displayMessages(obj);
		});
	} else {
		//alertify.alert('Please provide a valid email address.');
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	}
}*/

function displayMessages(obj) {
	var statusCode = obj.statusCode;
	if(statusCode == 212) {
		var invalidEmailIdStr = "";
		for (var index=0; index<obj.invalidEmailIds.length; index++) {
			invalidEmailIdStr = invalidEmailIdStr + obj.invalidEmailIds[index] + " <br> ";
		}
		alertify.alert("Following " +obj.invalidEmailIds.length+ " user(s) are invalid. Please fix the email id(s) and renew again.<br><br>"+"<b>"+invalidEmailIdStr+"</b>");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	} else if(statusCode == 601) {
		var nonExistEmailIdStr = "";
		for (var index=0; index<obj.nonExistingEmailIds.length; index++) {
			nonExistEmailIdStr = nonExistEmailIdStr + obj.nonExistingEmailIds[index] + "<br> ";
		}
		alertify.alert("Following " +obj.nonExistingEmailIds.length+ " user(s) does not exist. Please fix the email id(s) and renew again.<br><br>"+"<b>"+nonExistEmailIdStr+"</b>");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	} else if(statusCode == 2) {
		var inactiveEmailIdStr = "";
		for (var index=0; index<obj.inactiveEmailIds.length; index++) {
			inactiveEmailIdStr = inactiveEmailIdStr + obj.inactiveEmailIds[index] + " <br> ";
		}
		alertify.alert("Following " +obj.inactiveEmailIds.length+ " user(s) are not active. Please fix the email id(s) and renew again.<br>"+"<b>"+inactiveEmailIdStr+"</b>");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	} else if(statusCode == 200) {
		alertify.alert("User(s) has been renewed successfully.");
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	}
}

function isNumberKey(evt)
{
   var charCode = (evt.which) ? evt.which : event.keyCode;
   if (charCode > 31 && (charCode < 48 || charCode > 57))
      return false;

   return true;
}



$("#filename").change(function(e) {
	emailIdCSVString = "";
	var ext = $("input#filename").val().split(".").pop().toLowerCase();
	var emailIds = "";	
	if($.inArray(ext, ["csv"]) == -1 && ext.length > 0) {
		alertify.alert('Upload only CSV file.');
		$("input#filename").value = null;
		document.getElementById('filename').value="";
		return false;
	}
	if (e.target.files != undefined) {
		var reader = new FileReader();
		reader.onload = function(e) {
		var csvval=e.target.result.split("\r");
		totalUser = 0;
		for(var index = 1; index<csvval.length; index++) {	
			var csvvalue = csvval[index].split(",");
			if (csvvalue[0].trim() != "")
				emailIds = emailIds + csvvalue[0]+ "#`#"+ csvvalue[1] + "~";
			totalUser = totalUser + 1;
		}
		
	emailIdCSVString = JSON.stringify(emailIds.replace(/\n/g, "").replace(/\r/g, ""));
	};
	reader.readAsText(e.target.files.item(0));
	document.getElementById("filename").value = "";
}

return false;
});


function downloadCSVFmt() {
	var link = document.getElementById('downloadCSVFmt');
	link.setAttribute("href", "../../admin/manageuser/downloadIndividualRenewCSVFile/Individual_Renew.csv");
}

/**
 * Function to populate the existing users
 * under the logged in facilitator 
 */
function fetchUsersForRenew(paymentMode) {	
	var userGrp = Spine.Model.sub();
	userGrp.configure("/admin/manageuserFacilitator/populateRenewedUsersByFacilitator", "emailId", "customerId", "emailIdString");
	userGrp.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new userGrp({  
		emailId: sessionStorage.getItem("jctEmail"),
		customerId : sessionStorage.getItem("customerId"),
		//emailIdString: emailIdCSVString.substring(0, emailIdCSVString.length-1)
		emailIdString: emailIdCSVString
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
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		    emailIdCSVString = "";
			populateTable(obj.existingUsers, paymentMode);
		}  else {
			//Show error message
			alertify.alert("Some thing went wrong.");
		}
	});
}

/**
 * Function to show the existing users in table view
 * @param existingUsers
 */
function populateTable(existingUsers, paymentMode) {
	if(existingUsers.length > 0) {
		var createdBy = "";
		document.getElementById("existinglistId").innerHTML = "Renewed User Details";
		document.getElementById("updatedFaciTable").style.display="block";
		var userList = document.getElementById("mainPopulationDiv");
		if(paymentMode == "CHK") {
			paymentMode = "Check Payment";
		} else if (paymentMode == "CASH") {
			paymentMode = "Cash Payment";
		} else {
			paymentMode = "Free";
		}
		var tableStr = "<table width='94%' border='1' bordercolor='#78C0D3' id='myTable' align='center'><thead class='tab_header'><tr><th>SL. No.</th><th>Created By</th><th>User Name</th><th>Payment Mode</th><th>Expiration Date</th></tr></thead><tbody>";
		var counter = 1;
		for (var index=0; index<existingUsers.length; index++) {
			var trColor = "";
			if(index % 2 == 0) {
				trColor = "#D2EAF0";
			} else {
				trColor = "#F1F1F1";
			}
			var expiryDate = new Date(existingUsers[index].jctAccountExpirationDate).toDateString();	
			//var d = new Date();
			//var currentDate = new Date(d).toLocaleDateString();		
			//var currentDate = new Date(d).toDateString();		
			//var date1 = new Date(expiryDate);
			//var date2 = new Date(currentDate);
			//var timeDiff = Math.abs(date2.getTime() - date1.getTime());
//			var diffDays = Math.ceil(timeDiff / (1000 * 3600 * 24)); 
			var userExpryDate = dateformat(new Date (expiryDate));	
			
			var createdByData = existingUsers[index].customerId;
			var result = createdByData.substr(0, 2);
			
			if(result == "99") {
				createdBy = "Facilitator";				
			} else if (result == "98"){
				createdBy = "Admin";		
			} else {
				createdBy = "Ecommerce";		
			}
			tableStr = tableStr + "<tr class='user_list_row_width' bgcolor='"+trColor+"'><td align='center'>"+counter+".</td><td align='center'>"+createdBy+"</td><td align='center'>"+existingUsers[index].email+"</td><td align='center'>"+paymentMode+"</td><td align='center'>"+userExpryDate+"</td>";			
			counter = counter + 1;
		}
		tableStr = tableStr + "</tbody></table>";
		userList.innerHTML = tableStr;
		//new SortableTable(document.getElementById('myTable'), 1);
	} else {
		document.getElementById("updatedFaciTable").style.display="none";
		document.getElementById("existingUsersTableId").innerHTML = "<div align='center'><br /><br /><br /><img src='../img/no-record.png'><br /><div class='textStyleNoExist'>No Existing Users For Renewal</div></div>";
	}
	
}