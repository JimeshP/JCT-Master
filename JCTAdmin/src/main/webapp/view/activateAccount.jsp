<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<head>
    <title><spring:message code="label.activateAccount" /></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <!-- Bootstrap -->
    
    <link href="../css/bootstrap.css" rel="stylesheet">
    <link href="../css/style.css" rel="stylesheet">
    <link rel="stylesheet" href="../css/alertify.core.css" />
	<link rel="stylesheet" href="../css/alertify.default.css" id="toggleCSS" />
	<meta name="viewport" content="width=device-width">
	<link rel="shortcut icon" href="/admin/img/crafting_ico.ico" />
    <!-- HTML5 Shim and Respond.js IE8 support of HTML5 elements and media queries -->
    <!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
      <script src="https://oss.maxcdn.com/libs/respond.js/1.3.0/respond.min.js"></script>
    <![endif]-->
	<script type="text/javascript" src="../lib/jquery.js"></script>
	<script type="text/javascript" src="../lib/spine.js"></script>
	<script type="text/javascript" src="../lib/ajax.js"></script>
	<script src="http://code.jquery.com/jquery-latest.min.js"></script>
	<script>
	if (typeof jQuery == 'undefined') {
    	document.write(unescape("%3Cscript src='../js/latest_10.2_jquery.js' type='text/javascript'%3E%3C/script%3E"));
    }
	</script>	
	<script src="http://code.jquery.com/ui/1.10.3/jquery-ui.js" type="text/javascript"></script>
	<script type="text/javascript">
	//<![CDATA[
        (window.jQuery && window.jQuery.ui && window.jQuery.ui.version === '1.10.3')||document.write('<script type="text/javascript" src="../js/jquery_ui.js"><\/script>');//]]>
    </script>	
	<script src="../js/alertify.min.js"></script>
	<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css"> 
</head>
	<body>
    <div class="main_warp">

      <!-- Header area start -->
        <jsp:include page="loginHeader.jsp"/> 
       <!-- Header area end -->

        <!-- Form area start -->        
        <div class="row-fluid">                  
           <form class="form-horizontal" name="signupFrm" autocomplete="on" novalidate enctype="multipart/form-data" data-remote="true">
           <div class="container" id="surveyQuestionsDivContId">
             <div class="form_area">
               <div class="form_area_top">
                  <div class="col-md-4">
                    <p class="sign_up"><img src="../img/Add-Male-User.png"><span><spring:message code="label.survey.qtns"/></span></p>
                  </div>
                  <div class="clearfix"></div>
               </div>
               <div id="surveyQuestionsDivId">
               </div>
               
               <div id="surveyQtnSavebtnId" style="display: none;">
               <div class="single_form_item">
                     <p>
                     	<input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Save" id="saveDetails" onclick="saveSurveyDeatils()" />
                     </p>
               </div>
               </div>
               <div class="single_form_item">
                   <p><input type="checkbox" name="termsAndCondition" id="iAgree" value="" data-original-title="" >
                     	<spring:message code="label.tnc1"/> <a class="link" onclick="fetchTnstructions()" title="Click to see our terms & conditions"><spring:message code="label.tnc2"/></a></p>
               </div>
               <div id="surveyQtnNextbtnId" style="display: none;">
               <div class="single_form_item">
                     <p>
                     	<input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Next" id="saveDetails" onclick="updateActiveStatus()" />
                     </p>
               </div>
               </div>
             </div>
           </div>
           </form>
           
        </div>
         <!-- Form area end -->
        <!-- Footer area start -->
    <!-- Terms & Conditions modal box -->
	<div class="modal fade" id="tnCModal" tabindex="-1" data-backdrop="static">
	  <div class="modal-dialog" >
	    <div class="modal-content">
	      <div class="tcModal_header modal-header">
	        <button type="button" class="custom_close_btn" data-dismiss="modal"></button>      
	        <h5 class="modal-title tcModal_title" id="myModalLabel">Terms and Conditions</h5>
	      </div>
	      <br>
		   <div class="instruction_panel tcText_style" id="terms_condtn_div_id"></div>
	      <div class="modal-footer" style="padding: 5px 20px 8px;">
	      		<input type="submit" class="pop-up-btn tcClose-btn" value="Close" data-dismiss="modal">
	      </div>
	    </div>
	  </div>	  
	</div>
        
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
<script src="../js/common.js"></script>
<script src="../js/activateAccount.js"></script>
</body>
</html>