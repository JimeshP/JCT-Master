/*$(document).ready(function() {
var isMobile = {
    Android: function() {
        return navigator.userAgent.match(/Android/i);
        alert( "Android!" );
    },
    BlackBerry: function() {
        return navigator.userAgent.match(/BlackBerry/i);
        alert( "BlackBerry!" );
    },
    iOS: function() {
        return navigator.userAgent.match(/iPhone|iPad|iPod/i);
        alert( "IOS!" );
    },
    Opera: function() {
        return navigator.userAgent.match(/Opera Mini/i);
        alert( "Opera Mini!" );
    },
    Windows: function() {
        return navigator.userAgent.match(/IEMobile/i);
        alert( "IEMobile!" );
    },
    any: function() {
        return (isMobile.Android() || isMobile.BlackBerry() || isMobile.iOS() || isMobile.Opera() || isMobile.Windows());
    }
};


if(navigator.userAgent.match(/Android/i) ){
    alert( "Android!" );
    //navigator.userAgent.match(/iPhone|iPad|iPod/i);
    //alert( "IOS!" );
    //alert( "Mobile!" );
} 
else if(navigator.userAgent.match(/iPhone|iPad|iPod/i)){
	alert( " IOS ");
}
else {
	alert( " Not mobile ");
}

});*/
/*#293*/
$(document).ready(function(){
	adjustLogoImage();
	if(navigator.userAgent.match(/Android/i) ){
	//alert("android");
      $(".border").css({ height: "auto" });
    }
	
	if (sessionStorage.getItem("accountExp") && !(sessionStorage.getItem("msgDel"))) {
		if (sessionStorage.getItem("accountExp") != "NA") {
			alertify.alert("<img src='../img/alert-icon.png'><br /><p>"+sessionStorage.getItem("accountExp")+"</p>");
			sessionStorage.setItem("msgDel","Y");
		}
	}
	
	var src = "Tool";
	if (sessionStorage.getItem("src")) {
		src = sessionStorage.getItem("src");
	}
	if (src == "Tool") {
		if (sessionStorage.getItem("accountExp") != "NA") {
			if (document.getElementById("notificationIconDiv")) {
				document.getElementById("notificationIconDiv").style.display = "block";
			}
			/*if (document.getElementById("notificationTextDiv")) {
				document.getElementById("notificationTextDiv").style.display = "block";
			}*/			
		} else {
			if (document.getElementById("notificationIconDiv")) {
				document.getElementById("notificationIconDiv").style.display = "none";
			}
			/*if (document.getElementById("notificationTextDiv")) {
				document.getElementById("notificationTextDiv").style.display = "none";
			}*/
		}
		
	} else {
		if (document.getElementById("notificationIconDiv")) {
			document.getElementById("notificationIconDiv").style.display = "none";
		}
		if (document.getElementById("notificationTextDiv")) {
			document.getElementById("notificationTextDiv").style.display = "none";
		}
	}
});

function updateTime() {
	if (null != sessionStorage.getItem("rowIdentity")) {
		var beforeSketchModel = Spine.Model.sub();
		beforeSketchModel.configure("/user/auth/silentTimeUpdate", "rowId");
		beforeSketchModel.extend( Spine.Model.Ajax );
		var modelPopulator = new beforeSketchModel({  
			rowId: sessionStorage.getItem("rowIdentity")
		});
		modelPopulator.save();
	}
}

/**
 * Function navigates to specified page
 * @param page
 */
function navigateToPage(page) {
	window.location = page;
}

/**
 * Added for Progress Bar
 * function call when click on progress 
 */
$(document).ready(function() {	
	var pageSeq = sessionStorage.getItem("pageSequence"); // 1, 2, 3, 4 or 5
	if(pageSeq == 1) {
		$(".bs_step1").removeClass("done").addClass('activelist');
		$(".as_step1").removeClass("activelist");
	} 
	else if (pageSeq == 2){
		$(".bs_step1").removeClass("activelist").addClass('done');
		$(".bs_step2").removeClass("done").addClass('activelist');
		$(".as_step1").removeClass("activelist");
	}
	else if (pageSeq == 3){
		$(".bs_step1").removeClass("activelist").addClass('done');
		$(".bs_step2").removeClass("activelist").addClass('done');
		$(".as_step1").removeClass("done").addClass('activelist');
	}
	else if (pageSeq == 4){
		$(".bs_step1").removeClass("activelist").addClass('done');
		$(".bs_step2").removeClass("activelist").addClass('done');
		$(".as_step1").removeClass("activelist").addClass('done');
		$(".as_step2").removeClass("done").addClass('activelist');
	}
	else if (pageSeq == 5){
		$(".bs_step1").removeClass("activelist").addClass('done');
		$(".bs_step2").removeClass("activelist").addClass('done');
		$(".as_step1").removeClass("activelist").addClass('done');
		$(".as_step2").removeClass("activelist").addClass('done');
		$(".as_step3").removeClass("done").addClass('activelist');
	}
	
	//$(".form-horizontal").addClass('col-md-11');
	
  	if((navigator.userAgent.match(/iPhone/i)))
    {
  		//alert("xxxxxxxxx");
        $(".b4_sketch").addClass('col-xs-10');
    }
  	else{
  		//alert("yyyyyy");
        $(".form-horizontal").addClass('col-md-11');
    }
	
  	$(".progressCol").addClass('col-xs-1');
	$(".progressCol").addClass('col-md-1');	
    $(".progressBox").css({ right: "-300px"});
    $(".progressFlip").css({ right: "0", opacity: "1" }); 
    
    $(".progress-hide").click(function (){ 
    if(!(navigator.userAgent.match(/iPhone|iPad|iPod/i)) && !(navigator.userAgent.match(/Android/i)) ){
    	$(".progressCol").removeClass('up');
    	if(sessionStorage.getItem("pageSequence") == 1) {
    		relativeToAbsolute();
    	  }	    	
    } 
	   
  	  if($(".answer_no").hasClass("col-md-2")){
		  $(".answer_no").removeClass("col-md-2");
		  $(".answer_no").addClass("col-md-1");
		  $(".answer_input").removeClass("col-md-10");
		  $(".answer_input").addClass("col-md-11");
	  } 
    	
    $(".progressBox").animate({
    display: "none", 
    right: "-300px",
    opacity: "0",
    float: "left"
  }, 300, function(){
	  $(".progressCol").removeClass('up');
	  $(".progressCol").removeClass('col-md-3');
	  $(".progressCol").addClass('col-md-1');
	  $(".progressCol").removeClass('col-xs-6');
	  $(".progressCol").addClass('col-xs-1');
  });
 if((navigator.userAgent.match(/iPad|iPod/i)))
    {
        $(".form-horizontal"); // no action for ipad
    }
else{
	setTimeout(function(){
		 $(".form-horizontal").removeClass("col-md-9").addClass('col-md-11');
		  $(".form-horizontal").removeClass("col-md-6");
		
	},1200);
}
    
    $(".progressFlip").animate({
    	right: "0",
    	opacity: "1"
    	}, 1200); 
    $(".progressFlip").css({ 'z-index':'99'}); 
    /********** Issue# 150 ***********/
    if($(".answer_no").hasClass("col-md-2")){
        $(".answer_no").removeClass("col-md-2");
        $(".answer_no").addClass("col-md-1");
        $(".answer_input").removeClass("col-md-10");
        $(".answer_input").addClass("col-md-11");
    }
    if(!(navigator.userAgent.match(/iPhone|iPad|iPod/i)) && !(navigator.userAgent.match(/Android/i)) ){
    $(".border").css({ height: "auto" });  
   } 
    
	if(!navigator.userAgent.match(/Android/i) || !navigator.userAgent.match(/iPad|iPhone/i)){
	      $(".border").removeClass("height_eql");
	    }; 
 
    
    
    });
});

$(document).ready(function() {
  $(".progressFlip").click(function () { 	 
	  if(sessionStorage.getItem("pageSequence") == 1) {
		  setTaskOnPullOut();
	  }	 
	  $(".diagramCol").removeClass("col-xs-6").addClass('col-xs-1');
	  $(".diagramCol").removeClass("col-md-6").addClass('col-md-1');  	  
	  $(".outer_div_item .commonDrag").css("left", "");
	  $(".outer_div_item .commonDrag").css("top", "");
	  
	  if($(".answer_no").hasClass("col-md-1")){
		  $(".answer_no").removeClass("col-md-1");
		  $(".answer_no").addClass("col-md-2");
		  $(".answer_input").removeClass("col-md-11");
		  $(".answer_input").addClass("col-md-10");		  
	  }
	  
	  if(!(navigator.userAgent.match(/iPhone|iPad|iPod/i)) && !(navigator.userAgent.match(/Android/i)) ){
		   /* $(".oval_draggable").animate({ 
		    	left: "16px" 
		    		}, 100);  */
		    
		    /*$("#drawShape .ellipse_main img").animate({ 
		    	width: "110%" 
		    		}, 100);	*/
		   
	  } 
	  
	  $(".progressFlip").css({ 'z-index':'139'});
	  $(".progressBox").css({ 'z-index':'141'});
	  $(".progressFlip").animate({
		  right: "-100px",
		  opacity: "0"	  
		  }, 100);
	  
    $(".progressBox").animate({
    display: "block", 
    right: "0px",
    opacity: "1",
    float: "left"
  }, 1000);
    $(".diagramBox").animate({
        display: "none", 
        right: "-300px",
        opacity: "0",
        float: "left"
      }, 500);
    $(".diagramFlip").animate({
      	right: "0",
      	opacity: "1"
      	}, 1500); 
$(".progressCol").removeClass("col-md-1").addClass('col-md-3');
$(".progressCol").removeClass("col-xs-1").addClass('col-xs-6');
if(navigator.userAgent.match(/iPad/i)){
	//alert("Ipad");
	$(".form-horizontal"); // no action for ipad
     }
else{
   $(".form-horizontal").removeClass("col-md-11").addClass('col-md-9');
   $(".form-horizontal").removeClass("col-md-6");
}	 

$(".progressCol").addClass('up');  
   
  	/********** Issue# 150 **********/
	if($(".answer_no").hasClass("col-md-1")){
	    $(".answer_no").removeClass("col-md-1");
	    $(".answer_no").addClass("col-md-2");
	    $(".answer_input").removeClass("col-md-11");
	    $(".answer_input").addClass("col-md-10");              
	}
	
	
	if(!navigator.userAgent.match(/Android/i) && !navigator.userAgent.match(/iPad/i)){
      $(".border").addClass("height_eql");
    }; 
   // setTimeout(setOuterDivWhileDelete("increasePageWrap"), 1000);
    if(sessionStorage.getItem("pageSequence") == 1) {
    	setTimeout(function(){setOuterDivWhileDelete("increasePageWrap")}, 1000);  
    	setTimeout(function(){relativeToAbsolute()},1500);
	  }	
  });
});
  
  /**
   * Added for diagram Bar
   * function call while click on diagram
   */
  $(document).ready(function() { 
	 var img = sessionStorage.getItem("snapShotURLS");
	 if(document.getElementById("bsDiag")){
	 var div = document.getElementById("bsDiag");
	 div.innerHTML = "<img class='thumbnail' src=" +img+ " title='User Pic' height='150' width='100%' />";
	 
  	$(".form-horizontal").addClass('col-md-11');
  	
  	if(navigator.userAgent.match(/iPhone/i))
    {
        $(".form-horizontal").addClass('col-md-10');
    }
  	
  	$(".diagramCol").removeClass('col-md-6');
  	
  	$(".diagramCol").addClass('col-xs-1');
  	$(".diagramCol").addClass('col-md-1');	
      $(".diagramBox").css({ right: "-300px", opacity: "0"});
      $(".diagramFlip").css({ right: "0", opacity: "1" });   
      
      $(".diagram-hide").click(function () {  
    	  
      	  if($(".answer_no").hasClass("col-md-2")){
    		  $(".answer_no").removeClass("col-md-2");
    		  $(".answer_no").addClass("col-md-1");
    		  $(".answer_input").removeClass("col-md-10");
    		  $(".answer_input").addClass("col-md-11");
    	  }   
    	  
    	  
      $(".diagramBox").animate({
      display: "none", 
      right: "-300px",
      opacity: "0",
      float: "left"
    }, 500, function(){
  	  $(".diagramCol").removeClass('up');
    }  );
      $(".form-horizontal").removeClass('col-md-6').addClass('col-md-11');
      $(".form-horizontal").removeClass('col-md-9').addClass('col-md-11');
      $(".diagramFlip").animate({
      	right: "0",
      	opacity: "1"
      	}, 1500); 
      $(".diagramFlip").css({ 'z-index':'99'});  
           
      /********** Issue# 150 **********/
      if($(".answer_no").hasClass("col-md-2")){
          $(".answer_no").removeClass("col-md-2");
          $(".answer_no").addClass("col-md-1");
          $(".answer_input").removeClass("col-md-10");
          $(".answer_input").addClass("col-md-11");
      }
    	  $(".diagramCol").removeClass("col-xs-6").addClass('col-xs-1');

    	  $(".diagramCol").removeClass("col-md-6").addClass('col-md-1');  
   
      });
  }	
	
	 $(".diagramBox").hover(function() {
		 $(".diagramCol").removeClass("col-md-6").addClass('col-md-8');
  			}, function() {
  			 $(".diagramCol").removeClass("col-md-8").addClass('col-md-6');
  		});	 	
  });

    $(".diagramFlip").click(function () { 
    	
  	  $(".progressCol").removeClass('col-md-3');
	  $(".progressCol").addClass('col-md-1');
	  $(".progressCol").removeClass('col-xs-6');
	  $(".progressCol").addClass('col-xs-1');
    	
    	
  	  if($(".answer_no").hasClass("col-md-1")){
		  $(".answer_no").removeClass("col-md-1");
		  $(".answer_no").addClass("col-md-2");
		  $(".answer_input").removeClass("col-md-11");
		  $(".answer_input").addClass("col-md-10");		  
	  }
    	
    	
    	
	  $(".diagramFlip").animate({
	  right: "-100px",
	  opacity: "0"
	  }, 500);
    	  
      $(".diagramBox").animate({
      display: "block", 
      right: "0px",
      opacity: "1",
      float: "left"    
    }, 1500);     
      $(".progressBox").animate({
    	    display: "none", 
    	    right: "-300px",
    	    opacity: "0",
    	    float: "left"
    	  }, 500);
      $(".progressFlip").animate({
      	right: "0",
      	opacity: "1"
      	}, 500);

  $(".diagramCol").removeClass("col-xs-1").addClass('col-xs-6');
  $(".diagramCol").removeClass("col-md-1").addClass('col-md-6');  
  
  if(navigator.userAgent.match(/iPad/i)){
		//alert("Ipad");
		$(".form-horizontal"); // no action for ipad
	     }
	else{
		$(".form-horizontal").removeClass("col-md-11").addClass('col-md-6');
	}
  //$(".form-horizontal").removeClass("col-md-11").addClass('col-md-6');
  
  $(".diagramCol").addClass('up');
 
  /******** Issue# 150 *********/
  if($(".answer_no").hasClass("col-md-1")){
      $(".answer_no").removeClass("col-md-1");
      $(".answer_no").addClass("col-md-2");
      $(".answer_input").removeClass("col-md-11");
      $(".answer_input").addClass("col-md-10");             
    }
});  
  
/**
 * Function call while click on loguot button
 * from every page
 */ 
$("#logMeOut").click(function () {
	var src = "Tool";
	if (sessionStorage.getItem("src")) {
		src = sessionStorage.getItem("src");
	}
	// Enable the logout feature only if the user has logged in from the tool
	if (src == "Tool") {
		if(!sessionStorage.getItem('done')) {
			$(".loader_bg").fadeIn();
			$(".loader").fadeIn();
			sessionStorage.setItem("isLogout","Y");
			//Check the page sequence and store data related
			var page = sessionStorage.getItem("pageSequence");
			if(page == 1){			//Before Sketch
				goToNextPage();
			} else if(page == 2){	//Question and Answer
				saveQuestionnaireData('N');
			} else if(page == 3){	//After sketch page 1
				goToNextPage();
			} else if(page == 4){	//After sketch page 2
				goToNextPageAftSktch();
			} else if(page == 5){	//Action Plan
				sessionStorage.setItem("pdfRequired", "N");
				saveActionPlan();
			}  else if(page == 6){	//Result Page
				submitDecisionForLogout();
			} else if(page == 8){	//Final Page
				logoutFinalPage();
			}
			sessionStorage.setItem('done','D');
		}
		sessionStorage.removeItem('notificationIconColor');
	}
});

function logoutFinalPage() {
	var LogoutVO = Spine.Model.sub();
	LogoutVO.configure("/user/auth/logout", "jobReferenceString", "rowId", "lastActivePage");
	LogoutVO.extend( Spine.Model.Ajax );

	//Populate the model with data to transder
	var logoutModel = new LogoutVO({  
		jobReferenceString: sessionStorage.getItem("jobReferenceNo"),
		rowId: sessionStorage.getItem("rowIdentity"),
		lastActivePage: "/user/view/finalPage.jsp"
	});
	logoutModel.save(); //POST
	
	LogoutVO.bind("ajaxError", function(record, xhr, settings, error) {
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
		alertify.alert("Unable to connect to the server.");
	});
	
	LogoutVO.bind("ajaxSuccess", function(record, xhr, settings, error) {
		//var jsonStr = JSON.stringify(xhr);
		sessionStorage.clear();
		window.location = "../signup-page.jsp";
	});
}

function logoutGeneral(){
	//var lastActivePage = "/user/view/afterSketch1.jsp";
	var lastActivePage = "/user/view/landing-page.jsp";
	if (sessionStorage.getItem("pageSequence") <= 1) {
		lastActivePage = "/user/view/beforeSketch.jsp";
	} else if (sessionStorage.getItem("pageSequence") == 3) {
		lastActivePage = "/user/view/afterSketch1.jsp";
	}
	//alert(lastActivePage);
	//Create a model
	var LogoutVO = Spine.Model.sub();
	LogoutVO.configure("/user/auth/logout", "jobReferenceString", "rowId", "lastActivePage");
	LogoutVO.extend( Spine.Model.Ajax );
	
	//Populate the model with data to transder
	var logoutModel = new LogoutVO({  
		jobReferenceString: sessionStorage.getItem("jobReferenceNo"),
		rowId: sessionStorage.getItem("rowIdentity"),
		lastActivePage: lastActivePage
	});
	logoutModel.save(); //POST
	
	LogoutVO.bind("ajaxError", function(record, xhr, settings, error) {
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
		alertify.alert("Unable to connect to the server.");
	});
	
	LogoutVO.bind("ajaxSuccess", function(record, xhr, settings, error) {
		sessionStorage.clear();
		window.location = "../signup-page.jsp";
	});

}

$(document).ready(function() {
	if(sessionStorage.getItem("totalTime")) {
	var totalTime = sessionStorage.getItem("totalTime");
	var mystring = totalTime.replace(/:/g, " : ");
	if (document.getElementById('timeId')) {
		document.getElementById('timeId').innerHTML = mystring;
	}
	}
});

/**
 * function used by the landing page to move to 
 * the last saved page by the user.
 */
function start(){
	window.location = sessionStorage.getItem("url");
}

/*  oval shape animate H to V */
$(document).ready(function() { 
	$(".right_arrow_oval").click(function () { 
		  $(".vertical_oval").animate({
		  left: "18%",
		  opacity: "1"
		  }, 500);
		  $(".horizontal_oval").animate({
		  left: "-74%",
		  opacity: "0"
		  }, 500);
		  $(".right_arrow_oval").animate({
			  opacity: "0"
			  }, 500);
		  $(".left_arrow_oval").animate({
			  opacity: "1"
			  }, 500);	  
		  
	});
	
	$(".left_arrow_oval").click(function () { 
		  $(".vertical_oval").animate({
		  left: "98%",
		  opacity: "1"
		  }, 500);
		  $(".horizontal_oval").animate({
		  left: "14%",
		  opacity: "1"
		  }, 500);
		  $(".right_arrow_oval").animate({
			  opacity: "1"
			  }, 500);		  
		  $(".left_arrow_oval").animate({
			  opacity: "0"
			  }, 500);
	});	
	
	});


function toogleResetDiv() {
	document.getElementById("passwordResetLink").style.display = "none";
	document.getElementById("passwordResetArea").style.display = "block";
}

/**
 * function to call forgot password against Enter key.
 */
function searchMyKeyEvent(e)
{
    // look for window.event in case event isn't passed in
    if (typeof e == 'undefined' && window.event) { e = window.event; }
    if (e.keyCode == 13)
    {
        document.getElementById('resetMyPassword').click();
    }
}

function resetMyPassword() {
	document.getElementById("rstMyPaswId").disabled=true;
	var selectedUserName = sessionStorage.getItem("jctEmail");
	var existingPassword = $('#inputEPass').val();
	var newPassword = $('#inputNPass').val();
	var confirmPassword = $('#inputCPass').val();
	if(validateMyPasswordEntries(existingPassword, newPassword, confirmPassword) == true) {
		$(".loader_bg").fadeIn();
	    $(".loader").fadeIn();
	    
		//Create a model
		var passwordResetModel = Spine.Model.sub();
		passwordResetModel.configure("/user/auth/resetPassword", "forUser", "initialPassword", "userPassword", "status", "clickStatus", "dist", "mailedPwd");
		passwordResetModel.extend( Spine.Model.Ajax );
		
		//Populate the model
		var modelPopulator = new passwordResetModel({  
			forUser: selectedUserName,
			initialPassword: encodePass(existingPassword+existingPassword),
			userPassword: encodePass(newPassword+newPassword),
			status: "3",
			clickStatus: "normal",
			dist: 2,
			mailedPwd: "NA"
		});
		//POST
		modelPopulator.save();
		
		passwordResetModel.bind("ajaxSuccess", function(record, xhr, settings, error) {
			var jsonStr = JSON.stringify(xhr);
			var obj = jQuery.parseJSON(jsonStr);
			var statusCode = obj.statusCode;
			if(statusCode == 200 || statusCode == 128){ //success
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			    document.getElementById("inputEPass").value = "";
			    document.getElementById("inputNPass").value = "";
			    document.getElementById("inputCPass").value = "";
			    document.getElementById("passwordResetLink").style.display = "block";
				document.getElementById("passwordResetArea").style.display = "none";
				
			    alertify.alert("Your password has been reset.");
			} else {
				$(".loader_bg").fadeOut();
			    $(".loader").hide();
			    alertify.alert("You have provided incorrect password. Please try again.");
			} 
		});
		
		passwordResetModel.bind("ajaxError", function(record, xhr, settings, error) {
			alertify.alert("Unable to connect to the server.");
			$(".loader_bg").fadeOut();
		    $(".loader").hide();
		});
		document.getElementById("rstMyPaswId").disabled=false;
	} else {
		document.getElementById("rstMyPaswId").disabled=false;
	}
}
/**
 * function encodes the new password which is 
 * then sent over the network
 * @param text
 * @returns {String}
 */
function encodePass(text) { 
    var ref = "P!QR#S$T%U&VW(X)Y*Z+[,-].^/_0`1a2b3c4d5e6f7g8h9i:j;k<l=m>n?o@pAqBrCsDtEuFvGwHxIyJzK{L|M}N~O";
	var result="";
	for (var count=0; count<text.length; count++) {
	    var char=text.substring (count, count+1); 
	    var num=ref.indexOf (char);
	    var encodeChar=ref.substring(num+1, num+2);
	    result += encodeChar;
	}
	return result;
}

/**
 * function validates the reset password entries
 * which are provided by the user.
 * @param mailedPassword
 * @param newPassword
 * @param confirmPassword
 * @returns {Boolean}
 */
function validateMyPasswordEntries(existingPassword, newPassword, confirmPassword){
	//var illegalChars = /\W/; // allow letters, numbers, and underscores
	  var illegalChars = /^(?=.*\d)(?=.*[a-z])(?=.*[A-Z])(?=.*[^a-zA-Z0-9])(?!.*\s).{8,16}$/;
	if (existingPassword == "") {
		alertify.alert('Please enter your existing Password.');
        return false;
    }   else if (newPassword == "") {
    	alertify.alert('Please enter the New Password.');
        return false;
    } else if ((newPassword.length < 8) || (newPassword.length > 16)) {
    	alertify.alert('New Password should be between 8 - 16 characters.');
        return false;
    } else if (!illegalChars.test(newPassword)) {
    	alertify.alert('New Password should contain atleast one lowercase letter, one uppercase letter, one numeric digit, and one special character.');
        return false;
    } else if (confirmPassword == "") {	//JIRA Defect Id: VMJCT-31
    	alertify.alert('Please enter the Confirm Password.');
        return false;
    } else if ((confirmPassword.length < 8) || (confirmPassword.length > 16)) {
    	alertify.alert('Confirm Password should be between 8 - 16 characters.');
        return false;
    } else if (newPassword != confirmPassword){
    	alertify.alert('New Password is different from Confirm Password. Please correct it.');
        return false;
    } else if (existingPassword == newPassword) {
    	alertify.alert('Your password should be different from the one we sent you.');
        return false;
    }
	return true;
}

function showTooltip(){				
	var x = '69%';
	var y = '30%';	
       $('.tooltip1').stop().show(100);
       $('.tooltip1').css({ left: x, top: y });	   
	   $('#tooltipText').text('Click to see the notification');		//	adding the description to the tooltip	
}
function hideTooltip(){
	$('.tooltip1').stop().hide(10);
}
function fetchTnstructions(){
	var userProfile = parseInt(sessionStorage.getItem("profileId"));
	var tnC = Spine.Model.sub();
	tnC.configure("/user/auth/fetchTnC", "userProfile", "userType");
	tnC.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new tnC({  
		userProfile: userProfile,
		userType: 1
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
		}  else if (statusCode == 100){	
			alertify.alert(obj.statusDesc);	//	no TnC found			
		} else {
			//Show error message
			alertify.alert("Some thing went wrong while fetching Terms & Conditions.");
		}
	});
}

$('#icon').click(function(){
	$('.tooltip1').stop().hide(10);
	//alert(sessionStorage.getItem("accountExp"));
	//var notificationText = sessionStorage.getItem("accountExp").replace("<br />", "");	//	string from server side	
	var notificationText = sessionStorage.getItem("accountExp");
	//notificationText = notificationText.replace(/<br >/g, "");replace(/\<br[\/]*\>/g, "\n"));
	notificationText = notificationText.replace(/\<br [\/]*\>/g, "\n");
	if(sessionStorage.getItem('notificationIconColor') == 'red')
			sessionStorage.setItem('notificationIconColor','yellow');
			
	if( $('#notificationTextDiv').css('display') == 'none' ) {//	when no notification is showing, fadeIn the notification
		$('#notificationTextDiv').fadeIn(300);
		$('#notificationTextDiv p').text(notificationText);
		
		setTimeout(function () {		//	automatically hide in 5 sec
			$('#notificationTextDiv').fadeOut(100);				
		},5000);
		
	} else {
		$('#notificationTextDiv').fadeOut(100);					
	}
	if( sessionStorage.getItem('notificationIconColor') == 'yellow' ) {
		$('.icon').attr("src","../img/alert_yellow.png");
		$('#notificationIconDiv').css({'opacity':'0.7'});
	}
	
});	

if( sessionStorage.getItem('notificationIconColor') == 'red' ){
	$('.icon').attr("src","../img/alert_red.png");	
	$('#notificationIconDiv').css({'opacity':'1'});			
} else {
	$('.icon').attr("src","../img/alert_yellow.png");
	$('#notificationIconDiv').css({'opacity':'0.7'});
}

function adjustLogoImage() {
	var img = new Image();
	img.src = "/user/img/logo.png";
	if (img.height >= 53) {
		$('#headerLogoId').css('background-size','contain');
	} 
	document.getElementById("headerLogoId").style.display = "block";
}