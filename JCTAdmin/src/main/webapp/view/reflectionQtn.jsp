<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<head>
    <title><spring:message code="menu.sub.create.refQtn.group" /></title>
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
	    <jsp:include page="header.jsp"/>
	    <!-- Header area end --> 	       
	        <div class="row">  
	            <!-- <div class="col-sm-12 container container_mid" > -->
	            <div class="main-cont-wrapper" >
	            <div class="col-sm-2 pd-0 col">
	            	<!-- Menu area start -->
	    			<jsp:include page="sideMenuNew.jsp"/>
	   				 <!-- Menu area end -->   
	            </div>
	            
	            <div class="col-sm-10 col">
	            <div class= "heading_label">
	            	<ol class="breadcrumbAdmin">
           				<li class="first_nav"><spring:message code="reports.breadcrumb.manage.user.home"/> &nbsp;|</li>
           				<li class="first_nav"><spring:message code="menu.cont.configs"/> &nbsp;|</li>
           				<li class="second_nav" id='updateId'><spring:message code="menu.sub.create.refQtn.group"/></li>           				          				
             		</ol>
	            </div>
	            <div class= "sub_contain"><spring:message code="menu.sub.create.refQtn.group"/></div>
	            <div class= "main_contain">	  
	             <div class="form_area_top upper_div">
	             <div class= "inner_container" id="addHeaderDiv"></div> 
	            <form class="form-horizontal main_div" name="crtRefQtn" autocomplete="on" novalidate enctype="multipart/form-data" data-remote="true"> 	            
	           <!--  Reflection Question  -->
	           	<div class="single_form_item">
                     <div class="col-sm-4 select_option">
                     <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="menu.sub.select.option"/></label></div>
                     <div class="col-sm-6 sub_text_label select_option_radio">
                       <div class="col-sm-7 radio_text">
                          <input type="radio" name="refQtnFld" id="reftnQtn" value="R" onclick="checkSelectedOption(this)" checked = "checked"> <spring:message code="label.reflectionQtnRadio"/>
                       </div>
                       <div class="col-sm-5 radio_text">
                       	  <input type="radio" name="refQtnFld" id="actnPln" value="A" onclick="checkSelectedOption(this)"> <spring:message code="label.actionPlanRadio"/>
                       </div>
                     </div> 
                     <div class="col-sm-2 hidden-ipad">&nbsp;</div>  
                     <div class="clearfix"></div>
                 </div>
                  
                  
                 <!--  User Profile -->
                  <div class="single_form_item">
                     <div class="col-sm-4 select_option">
                     <label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="label.user.select.profile"/></label></div>
                     <div class="col-sm-5 select_option_radio">
                     	<select class="form-control-general form-control-reflectn-qtn" id="inputUserProfileInpt" onchange="changeUserProfile(this)">
                     		<option class="form-control-general" value=""><spring:message code="label.user.select.profile_drpDwn"/> </option>
                     	</select>
                     </div> 
                     <div class="col-sm-3 hidden-ipad">&nbsp;</div>  
                     <div class="clearfix"></div>
                 </div>  
                
                <!--  Reflection Question Description-->
                <div id="createRefQtnUpdate"> 
               <div class="single_form_item">
                   <div class="col-sm-4 select_option">
                   <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="label.reflectionQtnDecs"/></label></div>
                   <div class="col-sm-6 select_option_radio"><input type="text" class="form-control-general-form-redefined" maxlength="200" value="" id="qtnDescInptId" onkeypress="return validateDescription(event)" autofocus></div> 
                   <div class="col-sm-2 hidden-ipad">&nbsp;</div>  
                   <div class="clearfix"></div>
                </div>                           
               
               	<!--  Reflection Question Sub Description For Edit-->
				     
				<div id ="sudQuestionTotalEditDiv" >     
				   	
               <div id ="sudQuestionEditDiv"  style = 'display: none;'>
                              
               	<div class="single_form_item" >
                   <div class="col-sm-4 select_option">
                   <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="label.reflectionQtnSub"/></label></div>
                   <div class="col-sm-6 select_option_radio"><input type="text" class="form-control-general-form-redefined" maxlength="200" value="" id="qtnSubDescEditId" onkeypress="return validateDescription(event)" autofocus></div> 
                   <div class="col-sm-2 hidden-ipad">&nbsp;</div>  
                   <div class="clearfix"></div>
                </div> 
                
                 <!--  ADDED FOR JCT PUBLIC VERSION -->
               <!--  Global profile change div -->
               <div id= "globalProfileChgDiv" style="display: none;">
                <div class="single_form_item">                
                <div class="col-sm-4 select_option">
                	<label for="inputAD" class="col-sm-12 control-label text_label "><spring:message code="label.make.global.change"/></label>
                </div>
                <div class="col-sm-8 global_chg_checkbox">
                <input type="checkbox" class="" id="globalProfileCheckBoxId"> <span class="global_chg_text"><spring:message code="label.make.global.change.desc"/></span>
                </div>                  
                <div class="col-sm-2 hidden-ipad">&nbsp;</div>  
                   <div class="clearfix"></div>
                </div>
                </div>
                
               </div>
                               
                 <!--  Reflection Question Sub Description-->
                 <div id ="sudQuestionDiv">                
                  <!-- No Of Sub Qtn -->
                  <div id="noOFSubQtnDiv">
                   <div class="single_form_item">
                     <div class="col-sm-4 select_option">
                     <label for="inputQtnNo" class="col-sm-12 control-label text_label"><spring:message code="label.NoOFSubQtn"/></label></div>
                     <div class="col-sm-5 select_option_radio">
                     	<select class="form-control-general form-control-reflectn-qtn" id="inputNoOFSubQtn">
                     		<option class="form-control-general" value=""><spring:message code="label.select.noOfQtn_drpDwn"/> </option>                    		
                     	</select>
                     </div> 
                     <div class="col-sm-3 hidden-ipad">&nbsp;</div>  
                     <div class="clearfix"></div>
                   </div>
                   </div>
                
                <div id=sunQtnAdditionDiv></div>
                
                </div>
                
                <div class="single_form_item">
                 <div id ="refQuestionAddDiv" class="col-sm-6 right_allgn">                 
                  <div class="single_form_item">
                     <p>
	                     <input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Add" id="addRefQtnBtnId" onclick="saveRefQtn()" />                      
	                     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	                     <input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Reset" id="cancelRefQtnBtnId" onclick="resetValue()" />	                   	                     
                     </p>
                   </div>	
                  </div>
                  
                  <div id ="actionPlanAddDiv" class="col-sm-6 right_allgn" style = 'display: none;'>
                   <div class="single_form_item">
                     <p>
	                     <input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Add" id="addActnPlanBtnId" onclick="saveActionPlan()" />
	                     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	                     <input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Reset" id="cancelActnPlanBtnId" onclick="resetValue()" />
                     </p>
                   </div>	
                 </div>
                 
                 <div class="col-sm-3 chgOrderDivButton" id="chgOrderDiv">            
                   <input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Change Order" id="chgOrderBtnId" data-toggle="modal" data-target="#myModal" onclick="changeOrder()"/>
                   </div>
                   <div class="clearfix"> </div>
                 </div>                 
                 
                </div>
                </div>
                </form>
	            </div>	       
							
			<!-- Modal -->
			<div class="modal fade" id="myModal" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true">
			  <div class="modal-dialog">
			    <div class="modal-content">
			      <div class="modal-header">
			         <button type="button" class="custom_close_btn" data-dismiss="modal" aria-hidden="true"></button> 
			        
			        <div class="modal-header-desc" id="headerDescMainId">Change Main Question Order</div>
			        <div class="modal-header-desc" id="headerDescSubId">
			         <div class="col-md-4" id ="backBtnId"><button type="button" class="popupBackButton" onclick="backToPrevious()">Back</button></div>
			          <div class="col-md-8 sub-header-desc">Change Sub Question Order</div>
			        </div>
			         <div class="single_form_item">
			        <div class="col-md-6 profile-desc-area" id="profileDescId"></div>
			        <div class="col-md-6 chg-sub-qtn-link" id="chgSubQtnLinkId"><a class="click_highlight" href='#' id="chgSubQtnOrderBtnId" onclick="changeSubOrder()">Change Sub Question Order</a></div>	  
			        <div class="clearfix"></div>		        
			        </div> 
			        
			        <div class="main-qtn-desc-area" id="mainQtnDescId"></div>
			        <!-- <div class=" profile-desc-area" id="mainQtnDescId"></div> -->
			        
			        <div id="mainQtnListId" class="main-qtn-area"></div>
			        <div id="subQtnListId" class="main-qtn-area"></div>
			        
			      </div>
			      
			      <div class="modal-footer" id="mainQtnSaveBtnId">
			        <!-- <button type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" data-dismiss="modal">Close</button> -->
			        <button type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" onclick="saveMainQtnOrder()">Save</button>
			      </div>
			      <div class="modal-footer" id="subQtnSaveBtnId">
			        <!-- <button type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" data-dismiss="modal">Close</button> -->
			        <button type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" onclick="saveSubQtnOrder()">Save</button>
			      </div>
			    </div>
			  </div>
			</div>
						
	             <div class="existing_data_div">             
	             <div  id="existing_list_Id" class= "existing_list">Existing Reflection Question</div>
	            <div id="existingRefQtnListId">
	            	<div align="center">
	            		<img src="../img/dataLoading.GIF" />
	            		<p> Loading Existing Reflection Question </p>
	            	</div>
	            </div> 
	            </div>
	            </div>
	            
	                <div class="clearfix"></div>     
	            </div>	            
	            
				</div>
	    	</div>
	    	</div>
	        <!-- Form area end -->
	    <!-- Footer area start -->
	    <jsp:include page="footer.jsp"/>
	    <!-- Footer area end --> 	          
		
		
		
		<script src="https://code.jquery.com/jquery.js"></script>    
<!-- Include all compiled plugins (below), or include individual files as needed -->
<script src="../js/bootstrap.min.js"></script>
<script src="../js/reflectionQuestion.js"></script>
<script src="../js/common.js"></script>
<script src="../js/jquery.tablesorter.js"></script>
<script src="../js/jquery.tablednd.js"></script>
<script type="text/javascript">


$(document).ready(function() {
    $("#table-1").tableDnD();
});


</script>
</body>
</html>