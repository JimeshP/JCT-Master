<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<head>
    <title>..: Before Sketch 1 :..</title>
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
	<script type="text/javascript" src="../lib/html2canvas.js"></script>
	<!-- <script src="http://code.jquery.com/jquery-1.10.1.min.js"></script>  -->
	<script src="http://code.jquery.com/jquery-latest.min.js"></script>	
    <script src="http://code.jquery.com/ui/1.10.3/jquery-ui.js"></script>
    <script src="../js/alertify.min.js"></script>
	 <link rel="stylesheet" href="http://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css">
	 <script type="text/javascript" src="../js/jquery.ui.touch-punch.min.js"></script>  
</head>
<body>
    <div class="main_warp">
    
     <!-- Header area start -->
        <jsp:include page="header.jsp"/>
     <!-- Header area end --> 
       
     <!-- Menu area start -->
     	<jsp:include page="mainMenu.jsp"/>
     <!-- Menu area end -->   
        
        <div class="row-fluid">  
           <div class="container" >
             <div class="form_area">
               <div class="form_area_top">
                  <div class="col-md-2 ">
                    <a href="#" class="btn btn-info prev-btn"><span class="glyphicon glyphicon-chevron-left"></span> <spring:message code="label.previousStep"/></a>
                    <div class="clearfix"></div>
                  </div>
                  <div class="col-md-8 text-center area-title">
                   <spring:message code="label.afterSketch"/>
                  </div>
                  <div class="col-md-2 ">
                  <a href="#" onclick="goToNextPageAftSktch()" class="btn btn-info prev-btn pull-right"><spring:message code="label.nextStep"/> <span class="glyphicon glyphicon-chevron-right"></span></a>
                  </div>
                  <div class="clearfix"></div>
               </div>
                      
                  
                 <div class="breadcrumb_area">
                 <div class="col-md-6">
                  <ol class="breadcrumb nobg">
                <li><spring:message code="label.stepOne"/></li>
                <li class="active"><spring:message code="label.defineYourTask"/></li>
                </ol>
                 </div>
                
                 </div> 
                  <div class="clearfix"></div> 
                   
                      
                <!--  <div class="form_area_top1">
                  <div class="col-md-3"><p><button type="button" class="btn_prev">PREV STEP</button></p></div>
                  <div class="col-md-6"><h6 class="sub_heading">BEFORE SKETCH</h6> </div>
                  <div class="col-md-3"><p><button type="button" class="btn_next" onclick="goToNextPage()">NEXT STEP</button></p></div>
                  <div class="clearfix"></div>
               </div>  -->             
               <div class="form_area_middle">
                 <div class="form_area_wraper">
                 <form class="form-horizontal" autocomplete="on" novalidate>                 
                 <!--  <div class="single_form_item">
                 <div class="col-md-1"><h6 class="sub_text">Step 1 ></h6></div>    
                 <div class="col-md-2"><h6 class="sub_text">Define Your Task</h6></div>
                 <div class="col-md-5"></div>              
                 <div class="col-md-4"><p><button type="button" class="btn_add_task" id="addtask">+ Add New Task</button></p></div>
                   <div class="clearfix"></div> 
                 </div> -->
                
                   <div class="col-md-12"> 
                 <div id="outderDiv" class="pageWrap_after_sketch">
                 	
                 </div> 
                   
                 </div>                  
                   </form>
                   
                    <!-- Progress Bar start -->
     					 <jsp:include page="progressBar.jsp"/>
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
        <div class="row-fluid footer">
          <h5 class="align_center footer_heading"><spring:message code="label.footer"/></h5>
        </div>
        <!-- Footer area end -->
    </div>
<script src="../js/bootstrap-datepicker.js"></script>
<script src="../js/datepicker-initialize.js"></script>
<script src="../js/bootstrap-multiselect.js"></script>
<script src="../js/bootstrap.min.js"></script>
<script src="../js/bootstrap-filestyle.js"></script>
<script src="../js/session_storage.js"></script>
<script src="../js/afterSketch3.js"></script>
<script src="../js/stopWatch.js"></script>
<script src="../js/common.js"></script>
</body>
</html>