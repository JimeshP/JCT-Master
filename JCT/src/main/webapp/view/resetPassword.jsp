<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<head>
    <title>..: Reset Password :..</title>
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
<body onload="showDescOnLoad();">
    <div class="main_warp">

      <!-- Header area start -->
        <jsp:include page="loginHeader.jsp"/> 
       <!-- Header area end -->

        <!-- Form area start -->        
        <div class="row-fluid">  
           <div class="container" >
             <div class="form_area">
               <div class="form_area_top">
                  <div class="col-md-4">
                    <p class="sign_up"><img src="../img/Add-Male-User.png"><span id="resetPwdId"><spring:message code="label.resetPassword"/></span></p>
                  </div>
                  <div class="clearfix"></div>
               </div>
               <div id="passwordResetForm">
               <div class="form_area_middle">
                 <div class="form_area_wraper">
                 <form class="form-horizontal" autocomplete="on" novalidate>
                 <div class="single_form_item">
                     <div class="col-md-4 ">
                     <label for="inputPwd" class="col-md-12 control-label align-right"></label></div>
                     <div class="single_form_item"><span id="dispalyMessage"></span></div> 
                     <div class="col-md-3">&nbsp;</div>  
                     <div class="clearfix"></div>
                   </div>
                   <div class="single_form_item">
                     <div class="col-md-4 ">
                     <label for="inputPwd" class="col-md-12 control-label align-right mandatory_field"><spring:message code="label.systemGeneratePassword"/></label></div>
                     <div class="col-md-5"><input type="password" class="form-control-general" id="mailedPassword" onkeypress="searchKeyEvent(event);" autofocus></div> 
                     <div class="col-md-3">&nbsp;</div>  
                     <div class="clearfix"></div>
                   </div>
                   <div class="single_form_item">
                     <div class="col-md-4 ">
                     <label for="inputNPwd" class="col-md-12 control-label align-right mandatory_field"><spring:message code="label.newPassword"/></label></div>
                     <div class="col-md-5"><input data-container="body" data-toggle="popover" data-placement="right" type="password" data-content="New Password should contain atleast one lowercase letter, one uppercase letter, one numeric digit, and one special character."  class="form-control-general" value="" id="inputNPwd" onkeypress="searchKeyEvent(event);"></div> 
                     <div class="col-md-3">&nbsp;</div>  
                     <div class="clearfix"></div>
                   </div>
                   <div class="single_form_item">
                     <div class="col-md-4 ">
                     <label for="inputCPwd" class="col-md-12 control-label align-right mandatory_field"><spring:message code="label.confirmPassword"/></label></div>
                     <div class="col-md-5"><input type="password" class="form-control-general" value="" id="inputCPwd" onkeypress="searchKeyEvent(event);"></div> 
                     <div class="col-md-3">&nbsp;</div> 
                     <div class="clearfix"></div>
                   </div>                                  
                   <div class="single_form_item">
                     <p><input type="button" class="btn btn-primary btn-sm" value="Reset" id="resetMyPassword" /></p>
                   </div>
                   </form>
                   <div class="clearfix"></div>
                 </div>
               </div>
               </div>
               <div id="loginForm" style="display: none;">
               <div class="form_area_middle">
                 <div class="form_area_wraper">
                 <form class="form-horizontal" autocomplete="on" novalidate>
                 <div class="single_form_item">
                     <div class="col-md-4 ">
                     <label for="inputPwd" class="col-md-12 control-label align-right"></label></div>
                     <div class="single_form_item"><span id="dispalyMessageLogin"></span></div> 
                     <div class="col-md-3">&nbsp;</div>  
                     <div class="clearfix"></div>
                   </div>
                   <div class="single_form_item">
                     <div class="col-md-4 ">
                     <label for="inputPwd" class="col-md-12 control-label align-right"><spring:message code="label.emailId"/></label></div>
                     <div class="col-md-5"><input type="text" class="form-control-general" placeholder="Email Id" id="inputEmails" onkeypress="searchKeyPress(event);" autofocus></div> 
                     <div class="col-md-3">&nbsp;</div>  
                     <div class="clearfix"></div>
                   </div>
                   <div class="single_form_item">
                     <div class="col-md-4 ">
                     <label for="inputNPwd" class="col-md-12 control-label align-right"><spring:message code="label.password"/></label></div>
                     <div class="col-md-5"><input type="password" class="form-control-general" value="" id="pwd" onkeypress="searchKeyPress(event);"></div> 
                     <div class="col-md-3">&nbsp;</div>  
                     <div class="clearfix"></div>
                   </div>
                   <div class="single_form_item">
                     <p><input type="button" class="btn btn-primary btn-sm" value="Log In" id="loginBtn" /></p>
                   </div>
                   </form>
                   <div class="clearfix"></div>
                 </div>
               </div>
               </div>
               <div class="form_area_bottom"></div>
             </div>
           </div>
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
<script src="https://code.jquery.com/jquery.js"></script>    
<script src="../js/bootstrap-datepicker.js"></script>
<script src="../js/datepicker-initialize.js"></script>
<script src="../js/bootstrap-multiselect.js"></script>
<!-- Include all compiled plugins (below), or include individual files as needed -->
<script src="../js/bootstrap.min.js"></script>
<script type="text/javascript" src="../js/resetPassword.js"></script>
<script src="../js/login.js"></script>
<script src="../js/jquery.placeholder.js"></script>
<script>
    $('input[type=text], textarea').placeholder();
</script>
</body>
</html>