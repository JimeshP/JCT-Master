<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<head>
    <title><spring:message code="label.mappingYourself" /></title>
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
	
	 <script src="http://code.jquery.com/jquery-1.10.1.min.js"></script>
   
	<script src="http://code.jquery.com/ui/1.10.3/jquery-ui.js"></script>
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
        <!-- Form area start -->
        
	        <div class="container" id="header-wrap">	          	 
	        <div class="row"> 
             <div class="form_area">              
               <div class="form_area_top col-md-12 prev_nxt_sec">
                  <div class="col-md-2 col-xs-3 Prev_btn_area">
                    <a href="#" class="btn btn-info prev-btn" onclick="backToPreviousPage()"><span class="glyphicon glyphicon-chevron-left"></span><spring:message code="label.previousStep"/></a>
                    <div class="clearfix"></div>
                  </div>
                  <div class="col-md-8 col-xs-12 text-center heading-title">
                   <spring:message code="label.mappingYourselfCaps"/>
                  </div>
                  <div class="col-md-2 col-xs-3 Next_btn_area">
                  <a href="#" class="btn btn-info prev-btn pull-right" onclick="goToNextPage()"><spring:message code="label.nextStep"/><span class="glyphicon glyphicon-chevron-right"></span></a>
                  </div>
                  <div class="clearfix"></div>
               </div>
               
             	<!-- Instruction area start -->
             	<jsp:include page="instructionBar.jsp"/> 
             	<%-- <div class="instruction_area">
			    <div class="row_custom instruction">
		       		<div id="panel" style="display: none;">
			          	<ul class="ins-options">
				           	<li>To represent central aspects of yourself in your After Diagram, select Strengths, Values, and Passions from the options below.</li>
				           	<li>You can select two to four of each.</li>
				           	<li>These will help you decide how to craft your job in the next step.</li>
			         	</ul>
			        </div>
		        		<p class="slide"><a id="drp" href="#" class="btn-slide active">Slide Panel</a></p>
		       			<h4 class="ins-title"><spring:message code="label.instructions"/></h4>
		       			<div class="clearfix"></div> 
		      	</div>
		      	<div class="clearfix"></div> 
		      	</div> --%>
		      	 <!-- Instruction area start -->                                     
                  
               <div class="form_area_middle">
                 <div class="form_area_wraper after_sketchone_area">
                 <form class="form-horizontal col-xs-12" autocomplete="on" novalidate>                                  
                  <!-- Changed the order of Strength, Value , Passion boxes as Value, Strength , Passion 
                  	   Dated : 16.07.2014 -->
                  <!-- Value population div -->                    
                  <div class="col-md-4 col-xs-4 single_item">
                  	<div class="col-md-12 div_value value_grad">
                  		<div class="col-md-12 border value_border_grad"><spring:message code="label.value.desc"/></div> 
                 		<div class="div_scroll">  
                  			<div id="valuePopulateDiv"></div>
                 		</div>
                	</div>
                 </div>
                  
                 <!-- Strength population div -->    
                 <div class="col-md-4 col-xs-4 single_item ">
                 	<div class="col-md-12 col-xs-12 div_strength strength_grad">                 
                 		<div class="col-md-12 border strength_border_grad"><spring:message code="label.strength.desc"/></div>  
                 		<div class="div_scroll">
                 			<div id="strengthPopulateDiv"></div>
                		 </div>
                 	</div>
                 </div>

                 <!-- Passion population div -->                                     
                 <div class="col-md-4 col-xs-4 single_item">
                 	<div class="col-md-12 div_passion passion_grad">
                  		<div class="col-md-12 border passion_border_grad"><spring:message code="label.passion.desc"/></div> 
                  		<div class="div_scroll">  
                 			<div id="passionPopulateDiv"></div>
                 		</div>
                 	 </div>
                 </div> 
                 
                  <div class="clearfix"></div> 
              
                   </form>
                   <!-- Progress Bar start -->
     					<jsp:include page="progressBar.jsp"/>
     				<!-- Progress Bar end --> 
                   <div class="clearfix"></div> 
                   
                    <div id="pageWrap" class="col-xs-10 col-md-11 single_form_item outer_div_item">
                    	<!-- Changed the order of Strength, Value , Passion boxes as Value, Strength , Passion 
                  	  		 Dated : 16.07.2014 -->
                      	<div id="value_pageWrap" class="col-xs-4 col-md-4 value_outer_div" ></div>
                        <div id="strength_pageWrap" class="col-xs-4 col-md-4 strength_outer_div" ></div>                       
                        <div id="passion_pageWrap" class="col-xs-4 col-md-4 passion_outer_div" ></div>
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
                <jsp:include page="footer.jsp"/>
        <!-- Footer area end -->
        <div class="loader_bg" style="display:none"></div>
        <div class="loader" style="display:none"><img src="../img/Processing.gif" alt="loading"></div>
        
    </div>

<!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
 <!-- <script src="https://code.jquery.com/jquery.js"></script> --> <!-- ****commented out to make element dragable***** -->
<!--<script src="js/modal.js"></script>
<<!-- script src="js/tooltip.js"></script>-->
<script src="../js/bootstrap-datepicker.js"></script>
<script src="../js/datepicker-initialize.js"></script>
<script src="../js/bootstrap-multiselect.js"></script>
<!-- Include all compiled plugins (below), or include individual files as needed -->
<script src="../js/bootstrap.min.js"></script>
<script src="../js/bootstrap-filestyle.js"></script>
<script src="../js/session_storage.js"></script>
<script src="../js/afterSketch.js"></script>
<script src="../js/stopWatch.js"></script>
<script src="../js/common.js"></script>
<script src="../js/instructionBar.js"></script>
<!-- popover.js added to show the attribute description on mouse hover
	 Dated : 16.07.201 -->
<script src="../js/tooltip.js"></script>
<script src="../js/popover.js"></script>

  <div class="modal fade" id="myModal" tabindex="-1" data-backdrop="static">
  <div class="modal-dialog" >
    <div class="modal-content">
      <div class="new_modal_header modal-header">
        <button type="button" class="custom_close_btn" data-dismiss="modal"></button>      
        <h4 class="modal-title new_modal_title" id="myModalLabel">Mapping Yourself</h4>
      </div>
	 	  	<!-- <div id="instruct_header" class="page_title_instruct"></div>    -->
      		<!-- <div class="modal-body video_style" id="instruction_panel_video"></div>           
      		<div class="instruction_panel text_style" id="instruction_panel_text"></div> -->
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