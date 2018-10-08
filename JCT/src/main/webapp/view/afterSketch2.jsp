<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<head>
    <title><spring:message code="label.puttingItTogether"/></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <link rel="shortcut icon" href="/JCT/img/crafting_ico.ico" />
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
	<script type="text/javascript" src="../lib/html2canvas.js"></script>
	<!-- <script src="http://code.jquery.com/jquery-1.10.1.min.js"></script>  -->
	<script src="http://code.jquery.com/jquery-latest.min.js"></script>	
    <!-- <script src="http://code.jquery.com/ui/1.10.3/jquery-ui.js"></script>  -->
	<script src="../js/ui-10.js"></script>
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
        
	        <div class="container content-area" id="header-wrap">	          	 
	        <div class="row"> 
             <div class="form_area"> 
               <div class="form_area_top col-md-12 prev_nxt_sec">
                  <div class="col-xs-3 col-md-2 Prev_btn_area">
                    <a href="#" class="btn btn-info prev-btn" onclick="backToPreviousPage()"><span class="glyphicon glyphicon-chevron-left"></span> <spring:message code="label.previousStep"/></a>
                    <div class="clearfix"></div>
                  </div>
                  <div class="col-xs-12 col-md-8 text-center heading-title">
                   <spring:message code="label.puttingItTogetherCaps"/>
                  </div>
                  <div class="col-xs-3 col-md-2 Next_btn_area">
                  <a href="#" onclick="goToNextPageAftSktch()" class="btn btn-info prev-btn pull-right"><spring:message code="label.nextStep"/> <span class="glyphicon glyphicon-chevron-right"></span></a>
                  </div>
                  <div class="clearfix"></div>
               </div>
                      
               <!-- Instruction area start -->
               <jsp:include page="instructionBar.jsp"/>                        
               <!-- Instruction area end -->

               <div class="form_area_middle af_two_area">
                 <div class="form_area_wraper">
                 <form class="form-horizontal" autocomplete="on" novalidate>
                 <div class="breadcrumb_area">
                 <div class="button">
    				<a href="#" id="addtask"><spring:message code="label.addYourTask"/></a>
				 </div>
                 </div> 
                  <div class="clearfix"></div>                                
                 <div class="col-xs-12 col-md-12">
                 <div class="col-xs-4 col-md-3 draw_oval_Shape_area">
                 <div id="drawShape" class="draw_oval_Shape">
                 <div>
                 <!-- <output class="role_frame">Add Role Frame</output> -->
                 </div>
                 
                 <!-- Implementing New resizable Circle -->
                 
 				<div class="">
 				<div class="horizontal_oval_container col-xs-12 col-md-12 ">
 				 <div class="oval_draggable ui-widget-content horizontal_oval" style="border-style: none; width: 71%; display: block;"><output class="role_frame">Add Role Frame</output><img class="ellipse_small_oval" src="../img/Ellipse_small.png" alt="Ellipse small" /></div>
                 <!-- Image oval horizontal area start -->
                 <div id="ovalDraggable" class="oval_draggable ellipse_main ui-widget-content horizontal_oval" style="border-style: none; width: 71%; display: block;">
                 <img src="../img/left_top.png" width="20%" draggable="false"/>
                 <div class="removeService delete_oval idGenClass" onclick="deleteElement(this);">
                 <img src="../img/cross-black.png" alt="Delete" /></div>                 
                <!--  <div class="speech-bubble"><textarea maxlength="120"></textarea></div> -->
                
                 <div class="clearfix"></div>
                 <!-- <div id="ovalDraggable" class="oval_draggable ui-widget-content">	
                 <div class="speech-bubble"><textarea></textarea></div> -->                
				 </div>
				 <!-- Image oval horizontal area end -->
				 <div class="clearfix"></div>
				 </div>
				 <!--<div class="speech-bubble"><div class="role_text" >Role:</div><textarea maxlength="120"></textarea>
				
				<div class="ui-resizable-handle ui-resizable-se" id="nwgrip"></div>								
				</div>			 -->	 				 
				  <div class="clearfix"></div>	
				 </div>		 

                 </div> 
                 </div>
                 <div class="col-xs-8 col-md-9 shape_area"> 
                 <div id = "innerDiv" class="inner_div">
                 	<!-- Changed the order of Strength, Value , Passion boxes as Value, Strength , Passion 
                  	   	 Dated : 16.07.2014 -->
                 	<div id="value_pageWrap" class="col-md-4 value_inner_div common_div" ></div>
                 	<div id="strength_pageWrap" class="col-md-4 strength_inner_div common_div" ></div>
                    <div id="passion_pageWrap" class="col-md-4 passion_inner_div common_div" ></div>        
                 </div> 
                 </div>
                 <div class="clearfix"></div>
                  </div>
                  
                    <div class="clearfix"></div>                            
                   </form>
                   <div class="clearfix"></div>
               <div class="col-xs-12 col-md-11">
                 <div id ="pageWrapCopy" class="as_canvas_area" style="position: relative;">
	                 <div id="outderDiv" class="pageWrap_after_sketch as_canvas_area" style="height:600px;">
						<!-- <ul class="role-frame-menu">
						<p class="custom-menu-text">Change Position</p>
							<li id = "topLeft" data-action = "top-left">Top Left</li>
							<li id = "topRight" data-action = "top-right">Top Right</li>
							<li id = "bottomLeft" data-action = "bottom-left">Bottom Left</li>
							<li id = "bottomRight" data-action = "bottom-right">Bottom Right</li>
						</ul> -->
	                 	<div class="clearfix"></div>
	                 </div> 
	                 <div class="clearfix"></div>
                 </div>
                 </div>
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
       <%--  <div class="row-fluid footer">
          <h5 class="align_center footer_heading"><spring:message code="label.footer"/></h5>
        </div> --%>
         <!-- Footer area start -->
             <jsp:include page="footer.jsp"/>
        <!-- Footer area end -->
        <!-- Footer area end -->
         <div class="loader_bg" style="display:none"></div>
         <div class="loader" style="display:none"><img src="../img/Processing.gif" alt="loading"></div>
    </div>

    
    
    
<script src="../js/bootstrap-datepicker.js"></script>
<script src="../js/datepicker-initialize.js"></script>
<script src="../js/bootstrap-multiselect.js"></script>
<script src="../js/bootstrap.min.js"></script>
<script src="../js/bootstrap-filestyle.js"></script>
<script src="../js/session_storage.js"></script>
<script src="../js/fixDesign.js"></script>
<script src="../js/afterSketch2.js"></script>
<script src="../js/stopWatch.js"></script>
<script src="../js/common.js"></script>
<script src="../js/jquery.placeholder.js"></script>
<script src="../js/instructionBar.js"></script>
<script>
  	$('input[type=text], textarea').placeholder();
</script>

<div class="modal fade" id="myModal" tabindex="-1" data-backdrop="static">
  <div class="modal-dialog" >
    <div class="modal-content">
      <div class="new_modal_header modal-header">
        <button type="button" class="custom_close_btn" data-dismiss="modal"></button>      
        <h4 class="modal-title new_modal_title" id="myModalLabel">After Diagram</h4>
      </div>
	  <div class="instruction_panel text_style" id="instruction_panel_video_text"><br>
     	 		<div class="instruction_panel text_style" id="text_ins_before_video_id"></div> 
     	 		<div class="modal-body video_style" id="video_instruction_id"></div>
     	 		<!-- <div class="page_title_instruct"></div><br>   -->
     	 		<div class="instruction_panel text_style" id="text_ins_after_video_id"></div> 
     	 	</div>
      <div class="modal-footer">
      </div>
    </div>
  </div>
</div>

</body>
</html>