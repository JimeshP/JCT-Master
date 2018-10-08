/**
 * Function add to disable browser back button
 * while the page is loaded
 * @param null
 */
window.location.hash="";
window.location.hash="Again-No-back-button";//again because google chrome don't insert first hash into history
window.onhashchange=function(){window.location.hash="";};

/**
 * Function add to allow only alphabets and numbers.
 * @param evt
 */
function alphaNumericOnly(event) {	
	 var c = event.which || event.keyCode;
	 //alert(c);
	  if((c > 32 && c < 48) || (c > 57 && c < 65) || (c > 90 && c < 97) ||
	     (c > 122 && c !== 127))
	    return false;
	  
}

/**
 * This function logs out the admin
 * @param null
 */
function logout(){
	$(".loader_bg").fadeIn();
	$(".loader").fadeIn();
	var model = Spine.Model.sub();
	model.configure("/admin/authAdmin/logout", "none");
	model.extend( Spine.Model.Ajax );
	
	//Populate the model with data to transfer
	var modelPopulator = new model( {  
		none: ""
	});
	modelPopulator.save();
	sessionStorage.clear();
	window.location = "/admin/signup-page.jsp";
}

$(document).ready(function() {
	if(sessionStorage.getItem("userRole") != 2) {	//	admin then no T & C
		$("#tcHolder").show();
	}
	
	$(".first-level").click(function () {
		//alert("gdcjdgdh");
		$(".dropdown-submenu1 ul").show();
		$(".dropdown-submenu1 ul").addClass("menu_open");
		
	});
	$("ul.dropdown-menu").mouseout(function () { 
		$(".menu_open").hide();		
	});	
	if(document.getElementById("interfaceTitle")){
		document.getElementById("interfaceTitle").innerHTML = sessionStorage.getItem("interfaceTitleStr");		
	}	
});

/**
 * This function gives the date in format
 * @param inputdate
 */
function dateformat(inputdate) {
	var dateStr = inputdate.toDateString();
	var len=dateStr.length;
	var newDateStr=dateStr.substr(4,len);
	var newchar = '-';
	newDateStr = newDateStr.split(' ').join(newchar);
	return newDateStr;
}

/**/
function dateFormatForReports(inputDate){
	var spaceExp = inputDate.indexOf(" ");
	var strExp = inputDate.substring(0,spaceExp);
	var monthNum = strExp.substring(5,(strExp.length-3));
	var monthName = "";
	if(monthNum == '01'){
		monthName = 'Jan';
	}else if(monthNum == '02'){
		monthName = 'Feb';
	}else if(monthNum == '03'){
		monthName = 'Mar';
	}else if(monthNum == '04'){
		monthName = 'Apr';
	}else if(monthNum == '05'){
		monthName = 'May';
	}else if(monthNum == '06'){
		monthName = 'Jun';
	}else if(monthNum == '07'){
		monthName = 'Jul';
	}else if(monthNum == '08'){
		monthName = 'Aug';
	}else if(monthNum == '09'){
		monthName = 'Sep';
	}else if(monthNum == '10'){
		monthName = 'Oct';
	}else if(monthNum == '11'){
		monthName = 'Nov';
	}else if(monthNum == '12'){
		monthName = 'Dec';
	}
	var year = strExp.substring(0,4);
	var dateVal = strExp.substring(8,strExp.length);
	var dateFinal = monthName + "-" + dateVal + "-" + year;
	return dateFinal;
}

/**
 * Function added for amount validation
 **/
$( function() {
    $('.decimalPlace').blur(function(e){
    	$(".popover").css("display", "none");
    	var amount = $(this).val();
    	var len = amount.length;
    	var amountZero = amount.substring(0,1);
    	if(!($(this).val()=="")){
    		
    	 	if((!($(this).val().indexOf(".")!=-1))){
    	    	  if(len > 4){
    	    		  alertify.alert("Amount format should be 9999.99");
    	    		  this.value = "";
    	    	 }
    	    	  else if(len <= 4 && (!(amountZero == 0))){
    	    		  this.value = this.value + ".00";
    	    	  }
    	    	  else if(amountZero == 0){
    	    		  alertify.alert("Amount format should be 9999.99");
    	    		  this.value = "";
    	    	  }
    	       }
    	       else if (($(this).val().indexOf(".")!=-1)){
    	    	   if($(this).val().split(".")[0].length > 4){	//	only 4 digits are allowed before decimal place    	    		   
    	    		   alertify.alert("Amount format should be 9999.99");
    	    		   this.value = "";    	    		   
    	    	   } else {
    	    		   if($(this).val().split(".")[1].length == 0){
       	    			this.value = this.value + '00';
       	              } else if($(this).val().split(".")[1].length == 1){
	       	       		this.value = this.value + '0';
	       	           } else if($(this).val().split(".")[1].length == 2){
	       	    		   this.value =  this.value;
	       	    	   } else if($(this).val().split(".")[1].length > 2){
	       	    		   var dot = $(this).val().indexOf(".");
	       	    		   var decimalPlace = amount.substring(0,dot+3);
	       	    		   this.value =  decimalPlace;
	       	    	   }
    	    	   }
    	       }
    	}
   
         return this; //for chaining
    });
});

/**
 * Method to fetch terms & conditions for logged in facilitator
 * */
function fetchTnstructions(){
	var userProfile = parseInt(sessionStorage.getItem("profileId"));
	var tnC = Spine.Model.sub();
	tnC.configure("/admin/authAdmin/fetchTnC", "userProfile", "userType");
	tnC.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new tnC({
		userProfile: userProfile,
		userType: 3
	});
	modelPopulator.save(); //POST
	tnC.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	tnC.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if(statusCode == 200) {
			$('#tnCModal').modal('show');
			document.getElementById("terms_condtn_div_id").innerHTML = obj.jctTcDescription;			
		}  else if (statusCode == 100){		//	no TnC found
			alertify.alert(obj.statusDesc);
		} else {
			//Show error message
			alertify.alert("Some thing went wrong while fetching Terms & Conditions.");
		}
	});
}