<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<head>
    <title><spring:message code="menu.sub.create.userRole.group" /></title>
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
	            <div class="col-sm-2 pd-0">
	            	<!-- Menu area start -->
	    			<jsp:include page="sideMenuNew.jsp"/>
	   				 <!-- Menu area end -->   
	            </div>
	            
	            <div class="col-sm-10 col">
	            <div class= "heading_label">
	            	<ol class="breadcrumbAdmin">
           				<li class="first_nav"><spring:message code="reports.breadcrumb.manage.user.home"/> &nbsp;|</li>
           				<li class="first_nav"><spring:message code="menu.cont.configs"/> &nbsp;|</li>
           				<li class="second_nav" id='updateId'><spring:message code="menu.sub.create.userRole.group"/></li>           				          				
             		</ol>
	            </div>
	            <div class= "sub_contain"><spring:message code="menu.sub.create.userRole.group"/></div>
	            <div class= "main_contain">	  
	             <div class="form_area_top">
	             <div class= "inner_container" id="addUserRoleText"></div> 
	            <form class="form-horizontal main_div" name="crtUserRole" autocomplete="on" novalidate enctype="multipart/form-data" data-remote="true"> 	            
	          
	           <!--  User Role  -->
	           	<div class="single_form_item">
                     <div class="col-sm-4 select_option">
                     <label for="inputUR" class="col-sm-12 control-label text_label"><spring:message code="label.selectUserRole"/></label></div>
                     <div class="col-sm-6 sub_text_label select_option_radio">
                       <div class="col-sm-7 radio_text">
                          <input type="radio" name="userRoleFld" id="funcGroup" value="F" onclick="checkSelectedOption(this)" checked = "checked"> <spring:message code="label.functionGrp"/>
                       </div>
                       <div class="col-sm-5 radio_text">
                       	  <input type="radio" name="userRoleFld" id="jobLevel" value="J" onclick="checkSelectedOption(this)"> <spring:message code="label.jobLevel"/>
                       </div>
                     </div> 
                     <div class="col-sm-2 hidden-ipad">&nbsp;</div>  
                     <div class="clearfix"></div>
                   </div>
                               
               	
				     
				<div id ="sudQuestionTotalEditDiv" > 								
                
				<!--  Function Group -->
				<div id ="funcGroupDiv">
               	<div class="single_form_item" >
                   <div class="col-sm-4 select_option function_group">
                   <label for="inputFG" class="col-sm-12 control-label text_label"><spring:message code="label.functionGrpName"/></label></div>
                   <div class="col-sm-6 select_option_radio"><input type="text" class="form-control-general-form-redefined" maxlength="50" value="" id="fuctionGrpId" onkeypress="return alphaNumericOnly(event)" autofocus></div> 
                   <div class="col-sm-2 hidden-ipad">&nbsp;</div>  
                   <div class="clearfix"></div>
                </div> 
               </div>  
               
               <!--  Added for order-->
				<%-- <div id ="funcGroupOrderDiv">
               <div class="single_form_item">
                   <div class="col-sm-4 select_option">
                   <label for="inputQTNO" class="col-sm-12 control-label text_label"><spring:message code="label.functionGrpOrder"/></label></div>
                   <div class="col-sm-6 select_option_radio"><input type="text" class="form-control-general-form-redefined" maxlength="4" value="" id="fuctionGrpOrdertId" onkeypress="return numberOnly(event)"></div> 
                   <div class="col-sm-2 hidden-ipad">&nbsp;</div>  
                   <div class="clearfix"></div>
                </div> 
                </div> --%>
                
               <!--  Job Level -->                                
                    	
               <div id ="jobLevelDiv" style = 'display: none;'>
               	<div class="single_form_item" >
                   <div class="col-sm-4 select_option">
                   <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="label.jobLevelName"/></label></div>
                   <div class="col-sm-6 select_option_radio"><input type="text" class="form-control-general-form-redefined" maxlength="50" value="" id="jobLevelId" onkeypress="return alphaNumericOnly(event)" autofocus></div> 
                   <div class="col-sm-2 hidden-ipad">&nbsp;</div>  
                   <div class="clearfix"></div>
                </div> 
               </div>                            
                
                <%-- <div id ="jobLevelOrderDiv" style = 'display: none;'>
                <div class="single_form_item">
                   <div class="col-sm-4 select_option">
                   <label for="inputQTNO" class="col-sm-12 control-label text_label"><spring:message code="label.jobLevelOrder"/></label></div>
                   <div class="col-sm-6 select_option_radio"><input type="text" class="form-control-general-form-redefined" maxlength="4" value="" id="jobLevelOrderId" onkeypress="return numberOnly(event)" onblur="return changejobLevelOrder(this)"></div> 
                   <div class="col-sm-2 hidden-ipad">&nbsp;</div>  
                   <div class="clearfix"></div>
                </div> 
                </div>
                --%>
                
                 <div id ="functionGrpAddDiv" >                 
                  <div class="single_form_item">
                     <p><input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Add" id="addFuncGrpBtnId" onclick="saveFunctionGroup()" />                      
                     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                     <input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Reset" id="cancelFuncGrpBtnId" onclick="resetValue()" />
                     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;          
                     <input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Change Order" id="changeOrderFGBtnId" data-toggle="modal" data-target="#myModal" onclick="changeOrderFG()"/></p>
                   </div>	
                  </div>
                  
                 <div id ="jobLevelAddDiv" style = 'display: none;'>
                   <div class="single_form_item">
                     <p><input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Add" id="addJobLevelBtnId" onclick="saveJobLevel()" />
                     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                     <input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Reset" id="cancelJobLevelBtnId" onclick="resetValue()" />
					 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                     <input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Change Order" id="changeOrderJLBtnId" data-toggle="modal" data-target="#myModal" onclick="changeOrderJL()"/></p>
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
			        
			        <div class="modal-header-desc" id="headerDescMainId"></div>			       		         
			        <div class="pd-15"></div>
			        <!-- <div class="main-qtn-desc-area" id="mainQtnDescId"></div> -->
			        
			        <div id="mainAreaListId" class="main-qtn-area"></div>
			        <!-- <div id="subQtnListId" class="main-qtn-area"></div>  -->
			        

			      </div>
			      
			      <div class="modal-footer" id="FGSaveBtnId">
			        <!-- <button type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" data-dismiss="modal">Close</button> -->
			        <button type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" onclick="saveFuncGrpOrder()">Save</button>
			      </div>
			      
			      <div class="modal-footer" id="JLSaveBtnId" style="display: none">
			        <!-- <button type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" data-dismiss="modal">Close</button> -->
			        <button type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" onclick="saveJobLevelOrder()">Save</button>
			      </div>
		
			    </div>
			  </div>
			</div>
			      
	            <div class="existing_data_div">
	             
	             <div class= "existing_list" id="existing_data_list"></div>
	            <div id="existingUserRoleListId">
	            	<div align="center">
	            		<img src="../img/dataLoading.GIF" />
	            		<p> Loading Existing Data </p>
	            	</div>
	            </div> 
	            </div>
	            </div>
	            
	                     
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
<script src="../js/userRole.js"></script>
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