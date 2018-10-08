$(document).ready(function() {
	//document.getElementById("customerIdReadonly").innerHTML = sessionStorage.getItem("customerId");
	document.getElementById("customerIdReadonly").innerHTML = sessionStorage.getItem("jctEmail");
	document.getElementById("custEmail").value = sessionStorage.getItem("jctEmail");
	document.getElementById("custId").value = sessionStorage.getItem("customerId");
	
});

/**
 * Function to allow only numeric value
 */
$("#nos_of_user").keypress(function (e) {   
    if (e.which != 8 && e.which != 0 && (e.which < 48 || e.which > 57)) {       
       return false;
   }
});

function buyMoreSubscription () {
	if(validateNoOfUser()) {	
		document.getElementById("hiddenSubsNos").value = document.getElementById("nos_of_user").value;
		// fetch the shopify url details
		var model = Spine.Model.sub();
		model.configure("/admin/authAdmin/getShopifyLink", "none");
		model.extend( Spine.Model.Ajax );
		
		//Populate the model
		var modelPopulator = new model({  
			none: ""
		});
		//POST
		modelPopulator.save();
		
		model.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server. Please wait for some time and try again.");
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		});
		
		model.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			// populate the action and variant properties
			document.getElementById("ajaxform").action = "https://"+obj.shopifyFormUrl;
			document.getElementById("prdVarient").value = obj.productVarient;
			document.getElementById('ajaxform').submit();
		});
	}
}

function validateNoOfUser() {
	var userNo =  document.getElementById("nos_of_user").value;
	if(userNo == "") {	
		alertify.alert("Please enter no of users.");
		return false;
	} else if(userNo == 0) {
		alertify.alert("No of user cannot be 0.");
		return false;
	}
	return true;
}