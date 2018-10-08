

$("#proceed").click(function(){
	var userType = document.getElementById("selectUserType").options[document.getElementById("selectUserType").selectedIndex].value;
	var action = $("input[name=actionRadio]:checked").val();
	
	if(userType == 0){		
		alertify.alert("Please select user type");
		return false;
	} else if(action == 'N'){
		sessionStorage.setItem("userType",userType);	
		window.location = 'addUser.jsp';
	} else if(action == 'R'){
		sessionStorage.setItem("userType",userType);
		window.location = 'renew.jsp';			
	}
});


function toggleRadio(){
	var select = $("#selectUserType").val();
	if(select == 1){	//	individual		
		//$("#renewRadio").html("Renew existing");
		$("#subRadioDiv").hide();
		$("#renewRadioDiv").show();
	} else {	//	facilitator
		$("#renewRadio").html("Subscribe existing");
		$("#renewRadioDiv").hide();
		$("#subRadioDiv").show();
	}
}