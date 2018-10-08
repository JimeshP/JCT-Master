<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<head>
    <title><spring:message code="label.signup.title"/> </title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <!-- Bootstrap -->
    <link href="css/bootstrap.css" rel="stylesheet">
    <link href="css/style.css" rel="stylesheet">
    <link href="css/bootstrap-multiselect.css" rel="stylesheet">
    <link rel="stylesheet" href="css/alertify.core.css" />
	<link rel="stylesheet" href="css/alertify.default.css" id="toggleCSS" />
	<meta name="viewport" content="width=device-width">
	<link rel="shortcut icon" href="/admin/img/crafting_ico.ico" />

    <!-- HTML5 Shim and Respond.js IE8 support of HTML5 elements and media queries -->
    <!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
      <script src="https://oss.maxcdn.com/libs/respond.js/1.3.0/respond.min.js"></script>
    <![endif]-->
	<script type="text/javascript" src="lib/jquery.js"></script>
	<script type="text/javascript" src="lib/spine.js"></script>
	<script type="text/javascript" src="lib/ajax.js"></script>
	<script src="js/alertify.min.js"></script>
 
    </head>
<body onload="changeHashOnLoad();">

    <div class="main_warp">

      <!-- Header area start -->
         <jsp:include page="view/loginHeader.jsp"/> 
       <!-- Header area end -->

        <!-- Form area start -->
        
        <div class="row-fluid">  
           <div class="main-cont-wrapper" >
           		<div class="col-sm-12">
           			<div class="welcome_title row">
           			<div class="col-sm-8 col-md-8">Job Crafting<sup style='font-size:20px'>TM</sup> Facilitator</div>           			
           			<div class="col-sm-4 col-md-4 logoTitle" >
           			<img src="img/user_logo.png" class="img-responsive" onclick="window.open('http://www.jobcrafting.com','_blank');" alt="loading">
           			</div>           			
           			</div>           			
           		</div>
           		<div class="col-sm-12 main-login-area">
           			<form class="form-horizontal" name="signupFrm" autocomplete="on" novalidate enctype="multipart/form-data" data-remote="true">
	                  	<div class="single_form_item m-t-0">
	                  	 <div class="col-sm-12 heading_text"></div>
	                  	</div>
	                   <div class="single_form_item">
	                     <div class="col-sm-4 ">
	                      <label for="inputFN" class="col-sm-12 control-label align-right label_text" autofocus>User Name:</label></div>
	                     <div class="col-sm-6"><input type="text" class="form-control-general" maxlength="30" value=""  placeholder="Email ID" id="inputEmails" onkeypress="searchKeyPress(event);"></div> 
	                     <div class="col-sm-2">&nbsp;</div>  
	                     <div class="clearfix"></div>
	                   </div>
	                   <div class="single_form_item">
	                     <div class="col-sm-4 ">
	                     <label for="inputFN" class="col-sm-12 control-label align-right label_text">Password:</label></div>
	                     <div class="col-sm-6"><input type="password" class="form-control-general" maxlength="30" value="" placeholder="Password" id="pwd" onkeypress="searchKeyPress(event);"></div> 
	                     <div class="col-sm-2">&nbsp;</div>  
	                     <div class="clearfix"></div>
	                   </div>   	                 

	                   <div class="single_form_item align-center" style="margin-top:0%">
	                     <input type="button" class="btn btn-sm login-btn" value="Login" id="loginBtn" />
	                   </div>
	                   
	                   
					  <div class="single_form_item forgot-pass-link">
	                     <div class="col-md-2 ">&nbsp;</div>
	                     <div class="col-md-8 password_area">
	                     	<div class="col-md-12">
	                     		<!-- <i>If you haven forgotten your login credential,<a href="#" data-toggle="modal" data-target="#myModal-frgtpass" id="click"> click here</a> to reset.</i> -->	       
	                     		<a href="#" data-toggle="modal" data-target="#myModal-frgtpass">Forgot / Reset Password</a>
	                     	</div>
	                     	<div class="col-md-12"><a href="http://tool.jobcrafting.com/user/">Tool Login</a></div>
                     
	                     <div class="clearfix"></div>
	                     </div> 
	                     <div class="col-md-3">&nbsp;</div>  
	                     <div class="clearfix"></div>
	                   </div>                                   
	
	                   </form>
           		</div>
           		
           		
           </div>
        </div>
         <!-- Form area end -->
        <!-- Footer area start -->
        <%-- <div class="row-fluid footer">
          <h5 class="align_center footer_heading"><spring:message code="label.footer"/></h5>
        </div> --%>
                <jsp:include page="view/footerLanding.jsp"/>
        <!-- Footer area end -->
        <div class="loader_bg" style="display:none"></div>
        <div class="loader" style="display:none"><img src="img/Processing.gif" alt="loading"></div>
    </div>

<!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
<script src="https://code.jquery.com/jquery.js"></script>    
<!--<script src="js/modal.js"></script>
<script src="js/bootstrap-multiselect.js"></script>
<!-- Include all compiled plugins (below), or include individual files as needed -->
<script src="js/bootstrap.min.js"></script>
<script src="js/bootstrap-filestyle.js"></script>
<script src="js/login.js"></script>
<!-- <script src="js/disableRC.js"></script>  -->
<script type="text/javascript">
$(":file").filestyle();
</script>
  <script type = "text/javascript">
   function changeHashOnLoad() {
       window.location.href += "#";
       setTimeout(changeHashAgain, 50);
   }

   function changeHashAgain() 
   {          
       window.location.href += "1";
   }

   var storedHash = window.location.hash;
   window.setInterval(function () {
       if (window.location.hash != storedHash) {
           window.location.hash = storedHash;
       }
   }, 50);
   </script>
   
   
   
   <!-- Modal -->
<div class="modal fade" id="myModal-frgtpass" tabindex="-1">
  <div class="modal-dialog">
    <div class="modal-content">
      <div class="modal-header">
        <button type="button" class="close" data-dismiss="modal" style="margin-right: 3%;">&times;</button>
        <h4 class="modal-title" id="myModalLabel" style="    margin-left: 3%;margin-top: 3%;"><spring:message code="label.forgotPassword" /></h4>
      </div>
      <hr>
      <div class="modal-body">
                   <div class="single_form_item">
                     <div class="col-md-4 enter_email_lbl">
                     	<label for="inputFN" class="col-md-12 control-label align-right">Enter E-mail ID :</label>
                     </div>
                     <div class="col-md-6" style="margin-left: 4%;margin-top: -1%;">
                     	<input type="text" class="form-control-general input-sm" maxlength="30" id="forgotEmail" onkeypress="onKeyPress(event);" placeholder="Enter email" style="width: 95%; !important;margin-left: -3%;">
                     </div>  
                     <div class="clearfix"></div>
                   </div>
      </div>
       <hr>
      <div class="modal-footer" style="margin-top: -3%;">
        <input type="submit" class="btn btn-primary btn-sm cstm-login-btn reset_pass_admin" value="Reset Password" id="fgtMyPwd" onclick="forgotPwd()">
        <!-- <button type="button" class="btn btn-primary">Save changes</button> -->
      </div>
    </div>
  </div>
</div>

</body>
</html>