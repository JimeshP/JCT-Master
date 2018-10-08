<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<head>
    <title><spring:message code="label.reflectionQtn" /></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link rel="shortcut icon" href="/user/img/crafting_ico.ico" />
    <!-- Bootstrap -->
    <link href="../css/bootstrap.css" rel="stylesheet">
    <link href="../css/style.css" rel="stylesheet">
    <link href="https://s3-us-west-2.amazonaws.com/jobcrafting/css/commonStyle.css" rel="stylesheet">
    <link href="../css/less-768.css" rel="stylesheet">
    <link href="../css/less-1280.css" rel="stylesheet">  
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
	<script type="text/javascript" src="../lib/html2canvas.js"></script>
	<!-- <script src="http://code.jquery.com/jquery-1.10.1.min.js"></script>  -->
	<script src="http://code.jquery.com/jquery-latest.min.js"></script>
	<!-- <script type="text/javascript" src="../js/jquery.fancybox.pack.js"></script>
	<link rel="stylesheet" type="text/css" href="../css/jquery.fancybox.css" media="screen" /> -->
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
	    <jsp:include page="header.jsp"/>
	    <!-- Header area end --> 
	       
	    <!-- Menu area start -->
	    <jsp:include page="mainMenu.jsp"/>
	    <!-- Menu area end -->   
	        	          	 
	        <div class="container" id="header-wrap">  
	            <div class="row" >
	                <div class="form_area">
	                    <div class="form_area_top col-md-12 prev_nxt_sec">
	                        <div class="col-md-2 col-xs-3 Prev_btn_area">
	                    		<a href="#" class="btn btn-info prev-btn" onclick="navigateToPrevious()"><span class="glyphicon glyphicon-chevron-left"></span> <spring:message code="label.previousStep"/></a>
                    			<div class="clearfix"></div>
	                  		</div>
		                  	<div class="col-md-8 col-xs-12 text-center heading-title">
		                    	<spring:message code="label.reflectionQuestions"/>
		                    </div>
		                  	<div class="col-md-2 col-xs-3 Next_btn_area">
                  				<a href="#" class="btn btn-info prev-btn pull-right" onclick="saveQuestionnaireData('N')"><spring:message code="label.nextStep"/> <span class="glyphicon glyphicon-chevron-right"></span></a>
                 			</div>
		                  	<div class="clearfix"></div>
	               		</div>
	               		
		      			<div class="clearfix"></div> 
				      	 <!-- Instruction area start -->
      	 
	                 	<div class="breadcrumb_area">
	                 		<div class="clearfix"></div>
	                 	</div>  
	                  	<div class="clearfix"></div> 
	               		
	               		<div class="form_area_middle">
	                 		<div class="form_area_wraper Questionaire">
	                 			<form class="form-horizontal col-xs-11" autocomplete="on" novalidate>  
                 					<div id="questionArea">
                 	
                 					</div>                  
	                   			</form>
	                   
	                    		<!-- Progress Bar start -->
     					<jsp:include page="progressBarBottom.jsp"/>
     				<!-- Progress Bar end -->
     				
                     <!-- Diagram Bar start -->
     					<jsp:include page="diagramBarBs.jsp"/>
     				<!-- Diagram Bar end -->  
	     						<!-- Progress Bar end -->  
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
	        <!-- Footer area end -->
	        <div class="loader_bg" style="display:none"></div>
	        <div class="loader" style="display:none"><img src="../img/Processing.gif" alt="loading"></div>
	        <!-- Footer area end -->
		</div>
		<script src="../js/bootstrap-datepicker.js"></script>
		<script src="../js/datepicker-initialize.js"></script>
		<script src="../js/bootstrap-multiselect.js"></script>
		<script src="../js/bootstrap.min.js"></script>
		<script src="../js/bootstrap-filestyle.js"></script>
		<script src="../js/session_storage.js"></script>
		<script src="../js/questionaire.js"></script>
		<script src="../js/instructionBar.js"></script>
		<script src="../js/stopWatch.js"></script>
		<script src="../js/common.js"></script>
		
<!-- <div class="modal fade" id="myModal" tabindex="-1" data-backdrop="static">
  <div class="modal-dialog" >
    <div class="modal-content">
      <div class="new_modal_header modal-header">
        <button type="button" class="custom_close_btn" data-dismiss="modal" ></button>
        <button type="button" class="close">
		<div id="close_button" class="close_style">
		<img src="../img/cross-black.png" height="20px" width="20px"/>
		</div>
		</button>
        <h4 class="modal-title new_modal_title" id="myModalLabel">Reflection Question</h4>
      </div>
	  <div id="instruct_header" class="page_title_instruct">
	 <h3 class="page-title">Instructions:</h3>
	  </div>
	   
      <div class="modal-body video_style" id="instruction_panel_video">
    
         </div>           
                    <div class="instruction_panel text_style" id="instruction_panel_text">
                               
                  </div>
     
      <div class="modal-footer">
       <input type="submit" class="btn btn-primary btn-sm cstm-login-btn reset_pass" value="Reset Password" id="fgtMyPwd" onclick="forgotPwd()"> 
      </div>
    </div>
  </div>
</div> -->
	</body>
</html>