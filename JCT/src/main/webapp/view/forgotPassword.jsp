<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<head>
    <title>..: Forgot Password :..</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <!-- Bootstrap -->
    <link href="../css/bootstrap.css" rel="stylesheet">
    <link href="../css/style.css" rel="stylesheet">
    <link href="../css/datepicker.css" rel="stylesheet">
    <link href="../css/bootstrap-multiselect.css" rel="stylesheet">
	<link rel="stylesheet" href="../css/alertify.core.css" />
	<link rel="stylesheet" href="../css/alertify.default.css" id="toggleCSS" />
	<meta name="viewport" content="width=device-width">
    <!-- HTML5 Shim and Respond.js IE8 support of HTML5 elements and media queries -->
    <!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
      <script src="https://oss.maxcdn.com/libs/respond.js/1.3.0/respond.min.js"></script>
    <![endif]-->
	<script type="text/javascript" src="../lib/jquery.js"></script>
	<script type="text/javascript" src="../lib/spine.js"></script>
	<script type="text/javascript" src="../lib/ajax.js"></script>
	<script src="../js/alertify.min.js"></script>
    </head>
<body>
    <div class="main_warp">

      <!-- Header area start -->
        <jsp:include page="header.jsp"/>
       <!-- Header area end -->

        <!-- Form area start -->
        
        <div class="row-fluid">  
           <div class="container" >
             <div class="form_area">
               <div class="form_area_top">
                  <div class="col-md-4">
                    <p class="sign_up"><img src="../img/Add-Male-User.png"><span><spring:message code="label.forgotPassword"/></span></p>
                  </div>
                  <div class="clearfix"></div>
               </div>
               <div class="form_area_middle">
                 <div class="form_area_wraper">
                 <form class="form-horizontal" autocomplete="on" novalidate>
                  <div class="single_form_item">
                   &nbsp;
                   </div>
                   <div class="single_form_item"><p><spring:message code="label.forgotPasswordTest"/></div>                  
                   <div class="single_form_item">
                     <div class="col-md-4 ">
                     <label for="inputUN" class="col-md-12 control-label align-right"><spring:message code="label.userName"/>*</label></div>
                     <div class="col-md-5"><input type="text" class="form-control" value="" id="inputUN" autofocus></div> 
                     <div class="col-md-3">&nbsp;</div>  
                     <div class="clearfix"></div>
                   </div>
  		           <div class="single_form_item">
                     <p><button type="button" class="btn btn-primary btn-sm" id="mailMyPwd"><spring:message code="label.submit"/></button></p>
                   </div>
                   </form>
                   <div class="clearfix"></div>
                 </div>
               </div>
               <div class="form_area_bottom"></div>
             </div>
           </div>
        </div>
         <!-- Form area end -->
        <!-- Footer area start -->
                <jsp:include page="footer.jsp"/>
        <!-- Footer area end -->
        <div class="loader_bg" style="display:none"></div>
        <div class="loader" style="display:none"><img src="../img/Processing.gif" alt="loading"></div>
    </div>

<!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
<script src="https://code.jquery.com/jquery.js"></script>    
<!--<script src="js/modal.js"></script>
<script src="js/tooltip.js"></script>-->
<script src="../js/bootstrap-datepicker.js"></script>
<script src="../js/datepicker-initialize.js"></script>
<script src="../js/bootstrap-multiselect.js"></script>
<!-- Include all compiled plugins (below), or include individual files as needed -->
<script src="../js/bootstrap.min.js"></script>
<script src="../js/bootstrap-filestyle.js"></script>
<script src="../js/login.js"></script>

</body>

</html>