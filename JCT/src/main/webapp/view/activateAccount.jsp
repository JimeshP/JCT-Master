<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<head>
    <title><spring:message code="label.activateAccount" /></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link rel="shortcut icon" href="/user/img/crafting_ico.ico" />
    <!-- Bootstrap -->
    <link href="../css/bootstrap.css" rel="stylesheet">
    <link href="../css/style.css" rel="stylesheet">
    <link href="https://s3-us-west-2.amazonaws.com/jobcrafting/css/commonStyle.css" rel="stylesheet">
    <link href="../css/datepicker.css" rel="stylesheet">
    <link href="../css/bootstrap-multiselect.css" rel="stylesheet">
    <link rel="stylesheet" href="../css/alertify.core.css" />
	<link rel="stylesheet" href="../css/alertify.default.css" id="toggleCSS" />
    <!-- HTML5 Shim and Respond.js IE8 support of HTML5 elements and media queries -->
    <!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
      <script src="https://oss.maxcdn.com/libs/respond.js/1.3.0/respond.min.js"></script>
    <![endif]-->
	<script type="text/javascript" src="../lib/jquery.js"></script>
	<script type="text/javascript" src="../lib/spine.js"></script>
	<script type="text/javascript" src="../lib/ajax.js"></script>
	<!-- <script src="http://code.jquery.com/jquery-1.10.1.min.js"></script>  -->
	<script src="http://code.jquery.com/jquery-latest.min.js"></script>
	<script type="text/javascript" src="../js/jquery.fancybox.pack.js"></script>
	<link rel="stylesheet" type="text/css" href="../css/jquery.fancybox.css" media="screen" />
	<script>
	if (typeof jQuery == 'undefined') {
    	document.write(unescape("%3Cscript src='../js/latest_10.2_jquery.js' type='text/javascript'%3E%3C/script%3E"));
    }
	</script>	
	<script src="http://code.jquery.com/ui/1.10.3/jquery-ui.js" type="text/javascript"></script>
    <script type="text/javascript" src="../js/iscroll.js"></script>
	<script type="text/javascript">
	//<![CDATA[
        (window.jQuery && window.jQuery.ui && window.jQuery.ui.version === '1.10.3')||document.write('<script type="text/javascript" src="../js/jquery_ui.js"><\/script>');//]]>
    </script>	
	<script src="../js/alertify.min.js"></script>
	<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css">
	<script type="text/javascript" src="../js/jquery.ui.touch-punch.min.js"></script> 
    <script type="text/javascript" src="../js/tooltip.js"></script>
	<script type="text/javascript" src="../js/popover.js"></script>
	<script>
	$(document).ready(function() {
	$("input").focus(function() {
	//alert("focus sss");
	  $(this).popover('show');
	  });
	  
	  $("input").click(function() {
	  //alert("clicked");
	  $(this).popover('toggle');
	  });

	 $("input").blur(function() {
	 //alert("sdfdhfjld");
	  $(this).popover('hide');
	  });
	 });
	</script>
	<script src="../js/bootstrap.min.js"></script>
	<link href="../css/bootstrap.css" rel="stylesheet">
    <link href="../css/style.css" rel="stylesheet">  
    
	<script type="text/javascript" src="../js/common.js"></script>
	</head>
<body>
    <div class="main_warp">

      <!-- Header area start -->
        <jsp:include page="loginHeader.jsp"/> 
       <!-- Header area end -->

        <!-- Form area start -->        
        <div class="row-fluid">  
           <div class="container" >
             <div class="form_area" id="mainFrmDiv">
               <div class="form_area_top">
                  <div class="col-md-4">
                    <p class="sign_up"><img src="../img/Add-Male-User.png"><span id="resetPwdId"><spring:message code="label.resetPassword"/></span></p>
                  </div>
                  <div class="clearfix"></div>
               </div>
               <div id="passwordResetForm" style="display: none;">
               <div class="form_area_middle">
                 <div class="form_area_wraper">
                 <form class="form-horizontal" autocomplete="on" novalidate>
                 <div class="single_form_item">
                     <div class="col-md-4 " >
                     <label for="inputPwd" class="col-md-12 control-label align-right  mandatory_field"><spring:message code="label.systemGeneratePassword"/></label></div>
                     <div class="col-md-5" ><input type="password" class="form-control-general" id="mailedPassword" onkeypress="searchKeyEvent(event);" autofocus></div> 
                     <div class="col-md-3" >&nbsp;</div>  
                     <div class="clearfix"></div>
                   </div>
                   <div class="single_form_item">
                     <div class="col-md-4 ">
                     <label for="inputNPwd" class="col-md-12 control-label align-right mandatory_field"><spring:message code="label.newPassword"/></label></div>
                     <div class="col-md-5"><input data-container="body" data-toggle="popover" data-placement="right" data-content="New Password should contain at least one lowercase letter, one uppercase letter, one numeric digit, and one special character." type="password" class="form-control-general" value="" id="inputNPwd" onkeypress="searchKeyEvent(event);"></div> 
                     <div class="col-md-3">&nbsp;</div>  
                     <div class="clearfix"></div>
                   </div>
                   <div class="single_form_item">
                     <div class="col-md-4 ">
                     <label for="inputCPwd" class="col-md-12 control-label align-right mandatory_field"><spring:message code="label.confirmPassword"/></label></div>
                     <div class="col-md-5"><input  type="password" class="form-control-general" value="" id="inputCPwd" onkeypress="searchKeyEvent(event);"></div> 
                     <div class="col-md-3">&nbsp;</div> 
                     <div class="clearfix"></div>
                   </div>                            
                   <br />      
                   <div class="single_form_item">
                     <p><input type="button" class="btn btn-primary btn-sm" value="Reset" id="activateMyAccount" onClick="activateAccount()" /></p>
                   </div>
                   </form>
                   <div class="clearfix"><br /></div>
                 </div>
               </div>
               </div>
               <div id="loginForm" style="display: none;">
               <div class="form_area_middle">
                 <div class="form_area_wraper_justified_page">
                 <form class="form-horizontal" name="signupFrm" autocomplete="on" novalidate enctype="multipart/form-data" data-remote="true">
                   <div class="single_form_item">
                     <div class="col-md-4 ">
                      <label for="inputFN" class="col-md-12 control-label align-right"><spring:message code="label.firstName"/></label></div>
                    <!--  <label for="inputFN" class="col-md-12 control-label align-right">First Name*</label></div> -->
                     <div class="col-md-5"><input type="text" class="form-control-general" maxlength="30"  value="" id="inputFN" onkeypress="return validateDescription(event)" autofocus></div> 
                     <div class="col-md-3">&nbsp;</div>  
                     <div class="clearfix"></div>
                   </div>
                   <div class="single_form_item">
                     <div class="col-md-4 ">
                     <label for="inputFN" class="col-md-12 control-label align-right"><spring:message code="label.lastName"/></label></div>
                     <div class="col-md-5"><input type="text" class="form-control-general" maxlength="30" value="" id="inputLN" onkeypress="return validateDescription(event)"></div> 
                     <div class="col-md-3">&nbsp;</div>  
                     <div class="clearfix"></div>
                   </div>
                   
                   <%-- <div class="single_form_item">
                     <div class="col-md-4 ">
                     <label for="inputFN" class="col-md-12 control-label align-right"><spring:message code="label.location"/></label></div>
                     <div class="col-md-5">
                     	<!-- <input type="text" class="form-control-general" value="" id="inputLocation"> -->
                     	<select class="form-control-general" id="inputLocation">
                     		<option class="form-control-general" value=""><spring:message code="label.select.region"/> </option>
                     	</select>
                     </div> 
                     <div class="col-md-3">&nbsp;</div>  
                     <div class="clearfix"></div>
                   </div> --%>
                   
                   <div class="single_form_item">
                     <div class="col-md-4 ">
                     <label for="inputFN" class="col-md-12 control-label align-right label_device"><spring:message code="label.sex"/></label></div>
                     <div class="col-md-5">
                       <div class="col-md-6">
                          <input type="radio" name="gender" id="male" value="M"> <spring:message code="label.male"/>
                       </div>
                       <div class="col-md-6"><input type="radio" name="gender" id="female" value="F"> <spring:message code="label.female"/></div>
                     </div> 
                     <div class="col-md-3">&nbsp;</div>  
                     <div class="clearfix"></div>
                   </div>
                   
                   <!-- TENURE -->
                   <%-- <div class="single_form_item">
                     <div class="col-md-4 ">
                     <label for="inputFN" class="col-md-12 control-label align-right"><spring:message code="label.tenure"/></label></div>
                     <div class="col-md-5">
                     	<select class="form-control-general" id="inputTenure">
                     		<option class="form-control-general" value=""><spring:message code="label.select.tenure"/> </option>
                     		<option class="form-control-general" value="< 1 year"> &lt; 1 year </option>
                     		<option class="form-control-general" value="1-2 years">1-2 years</option>
                     		<option class="form-control-general" value="2-3 years">2-3 years</option>
                     		<option class="form-control-general" value="3-4 years">3-4 years</option>
                     		<option class="form-control-general" value="4-5 years +">4-5 years</option>
                     		<option class="form-control-general" value="5 years +">5 years +</option>
                     	</select>
                     </div> 
                     <div class="col-md-3">&nbsp;</div>  
                     <div class="clearfix"></div>
                   </div> --%>
                   
                   <!-- SUPERVISE PEOPLE -->
                   <%-- <div class="single_form_item">
                     <div class="col-md-4 ">
                     <label for="inputFN" class="col-md-12 control-label align-right"><spring:message code="label.supervise"/></label></div>
                     <div class="col-md-5">
                       <div class="col-md-6">
                          <input type="radio" name="supervisePeople" id="superviseYes" value="Y"> <spring:message code="label.yes"/>
                       </div>
                       <div class="col-md-6"><input type="radio" name="supervisePeople" id="superviseNo" value="N"> <spring:message code="label.no"/></div>
                     </div> 
                     <div class="col-md-3">&nbsp;</div>  
                     <div class="clearfix"></div>
                   </div> --%>
                   
                   <!-- FUNCTION GROUP -->
                   <%-- <div class="single_form_item">
                     <div class="col-md-4 ">
                     <label for="inputFN" class="col-md-12 control-label align-right mandatory_field"><spring:message code="label.functionGrp"/></label></div>
                     <div class="col-md-5">
                     	<select class="form-control-general" id="inputFunctionGroup">
                     		<option class="form-control-general" value=""><spring:message code="label.select.functionGrp"/> </option>
                     	</select>
                     </div> 
                     <div class="col-md-3">&nbsp;</div>  
                     <div class="clearfix"></div>
                   </div> --%>
                   
                   <!-- JOB LEVEL -->
                   <%-- <div class="single_form_item">
                     <div class="col-md-4 ">
                     <label for="inputFN" class="col-md-12 control-label align-right mandatory_field"><spring:message code="label.jobLevel"/></label></div>
                     <div class="col-md-5">
                     	<select class="form-control-general" id="inputJobLevel">
                     		<option class="form-control-general" value=""><spring:message code="label.select.jobLevel"/> </option>
                     	</select>
                     </div> 
                     <div class="col-md-3">&nbsp;</div>  
                     <div class="clearfix"></div>
                   </div> --%>
                   
                   <!-- OCCUPATION SEARCH -->
                   <div class="single_form_item">
	                   <div class="col-md-4 ">
	                   	   <label for="inputFN" class="col-md-12 control-label align-right mandatory_field"><spring:message code="label.occupation.search"/></label>
	                   </div>
	                   <div class="col-md-4"><input type="text" class="form-control-general" placeholder="Enter key words describing your occupation" value="" id="ocptnSrchTxtId" oninput="validateOnetData()"></div> 
	                   <div class="col-md-4"><input type="button" class="btn btn-primary btn-sm searchClass" value="Search" id="searchOccupation" onclick="searchOccupationSearch()" /></div>  
	                   <div class="clearfix"></div>
	               </div>
	               <div class="single_form_item">
                       <div class="col-md-4 ">
	                     <label for="inputFN" class="col-md-12 control-label align-right mandatory_field"><spring:message code="label.how.long.worked"/></label>
	                   </div>
	                   <div class="col-md-5">
	                        <!-- <input type="text" class="form-control-general" maxlength="2" value="" id="nosOfYrs" onkeypress="return validateNumber(event)"> -->
	                        <select class="form-control-general" id="nosOfYrsInput">
                     			<option class="form-control-general" value=""><spring:message code="label.select.tenure"/> </option>
                     		</select>
	                   </div> 
	                   <div class="col-md-3">&nbsp;</div>  
	                   <div class="clearfix"></div>
                   </div>
                   
	               <!-- Survey Question -->
	              <div id="surveyQuestionsDivId" style="display: block;">
                  </div>
	               
	              <div class="single_form_item">
                     <p><input type="checkbox" name="termsAndCondition" id="iAgree" value="" data-original-title="" >
                     	<spring:message code="label.tnc1"/> <a class="link" onclick="fetchTnstructions()" style="cursor: pointer;" title="Click to see our terms & conditions"><spring:message code="label.tnc2"/></a></p>
                   </div> 
                  <div class="single_form_item">
                     <p><input type="button" class="btn btn-primary btn-sm" value="Save" id="saveDetails" onclick="saveUserDeatils()" /></p>
                   </div>
                  </form>
                   <div class="clearfix"></div>
               </div>
               </div>
               <div class="form_area_bottom"></div>
             </div>
             <div id="userValidatedSection" style="display: none; width: 100%; text-align: center; color: red; top: 10%; float: left; position: relative;">
               <div class="form_area_middle">
                 <div class="form_area_wraper">
                 <form class="form-horizontal" autocomplete="on" novalidate>
                 <div class="single_form_item" >
                     <div class="col-md-12 activate_accnt_note" >
                     	<img src = "../img/failed-icon.png" /><br /><br/>
                     	<span id="noSuchUserId"><spring:message code="activate.account.password.created.note" /></span>
                     	<br />
                     	<spring:message code="activate.account.click.here.note" />
                     	<%-- <spring:message code="label.contact.sys.admin" />&nbsp;<a class="click_highlight" href='mailto:JCTSupport_vmw@interrait.com?cc=gtdops@vmware.com'><spring:message code="label.clickHere" /></a> --%>
                     </div>                    
                   </div>
                   </form>
                   <div class="clearfix"><br /></div>
                 </div>
               </div>
               </div>
             </div>
           </div>
           
           <%-- <div class="container" id="surveyQuestionsDivContId" style="display: block;">
             <div class="form_area">
               <div class="form_area_top">
                  <div class="col-md-4">
                    <p class="sign_up"><img src="../img/Add-Male-User.png"><span><spring:message code="label.survey.qtns"/></span></p>
                  </div>
                  <div class="clearfix"></div>
               </div>
               <div id="surveyQuestionsDivId" style="display: block;">
               </div>
               <div class="single_form_item">
                     <p>
                     	<input type="button" class="btn btn-primary btn-sm" value="Save Survey Questions" id="saveDetails" onclick="saveSurveyDeatils()" />
                     </p>
               </div>
             </div>
           </div> --%>
           
        </div>
         <!-- Form area end -->
        <!-- Footer area start -->
        <%-- <div class="row-fluid footer">
          <h5 class="align_center footer_heading"><spring:message code="label.forgotPasswordTest"/></h5>
        </div> --%>
        <jsp:include page="footer.jsp"/>
        <!-- Footer area end -->
        <div class="loader_bg" style="display:none"></div>
        <div class="loader" style="display:none"><img src="../img/Processing.gif" alt="loading"></div>
    </div>

<!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
<!-- <script src="https://code.jquery.com/jquery.js"></script> -->   
<!-- Include all compiled plugins (below), or include individual files as needed -->
<script src="../js/bootstrap.min.js"></script>
<script type="text/javascript" src="../js/activateAccount.js"></script>
<script src="../js/jquery.placeholder.js"></script>
<script>
    $('input[type=text], textarea').placeholder();
</script>
<div id='occupationDiv' class="scroll">
	<div id="headerPanel" class="popupHeader custom_header">
		Select the option that best describes your occupation. Feel free to try different search words to find a better fit.
	</div>
	<div align="center" id="occupationPlottingDiv">
		<br /><br />
		<img src="../img/Processing.gif" />
	</div>
	<div align="center" id="dataPlottingDiv" style="height: 82%; overflow: auto;">
	</div>
	<div align="center" style="display: none;" id="goDiv">
		<input type="button" class="btn btn-primary btn-sm" value="Done" onclick="closeDiv()" />
	</div>
</div>
<!-- Terms & Conditions modal box -->
	<div class="modal fade" id="myModal" tabindex="-1" data-backdrop="static">
	  <div class="modal-dialog" >
	    <div class="modal-content">
	      <div class="tcModal_header modal-header">
	        <button type="button" class="custom_close_btn" data-dismiss="modal"></button>      
	        <h5 class="modal-title tcModal_title" id="myModalLabel">Terms and Conditions</h5>
	      </div>
	      <br>
		   <div class="instruction_panel tcText_style" id="terms_condtn_div_id_survey"></div>
	      <div class="modal-footer" style="padding: 5px 20px 8px;">
	      		<input type="submit" class="btn btn-primary btn-sm cstm-login-btn reset_pass tcClose-btn" value="Close" data-dismiss="modal">
	      </div>
	    </div>
	  </div>
	  
	</div>


</body>
</html>