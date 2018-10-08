$(function(){	
	if (sessionStorage.getItem("expDate")) {
		$("#expDate").text(sessionStorage.getItem("expDate"));
	}	
});

function submitForm() {
	// fetch the shopify url details
	var model = Spine.Model.sub();
	model.configure("/user/auth/getShopifyLink", "none");
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
		document.getElementById("custEmail").value = sessionStorage.getItem("jctEmail");
		document.getElementById('ajaxform').submit();
	});
}