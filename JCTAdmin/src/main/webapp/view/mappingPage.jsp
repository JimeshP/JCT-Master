<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<head>
    <title><spring:message code="menu.sub.create.mapping.values" /></title>
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
           				<li class="second_nav" id='updateId'><spring:message code="menu.sub.create.mapping.values"/></li>           				          				
             		</ol>
	            </div>
	            <div class= "sub_contain"><spring:message code="menu.sub.create.mapping.values"/></div>
	            <div class= "main_contain">	  
	             <div class="form_area_top">
	             <div class= "inner_container" id="addHeaderDiv"></div> 
	            <form class="form-horizontal main_div" name="crtMapping" autocomplete="on" novalidate enctype="multipart/form-data" data-remote="true"> 	            
	           <!--  Mapping Values  -->
	           	<div class="single_form_item">
                     <div class="col-sm-4 select_option">
                     <label for="inputMP" class="col-sm-12 control-label text_label"><spring:message code="label.mappingValueDesc"/></label></div>
                     <div class="col-sm-6 sub_text_label select_option_radio">
                       <div class="col-sm-4 radio_text">
                          <input type="radio" name="mappingValFld" id="strengthId" value="S" onclick="checkSelectedOption(this)" checked = "checked"> <spring:message code="label.mappingStrength"/>
                       </div>
                       <div class="col-sm-4 radio_text value_radio_text">
                       	  <input type="radio" name="mappingValFld" id="valueId" value="V" onclick="checkSelectedOption(this)"> <spring:message code="label.mappingValue"/>
                       </div>
                       <div class="col-sm-4 radio_text">
                       	  <input type="radio" name="mappingValFld" id="passionId" value="P" onclick="checkSelectedOption(this)"> <spring:message code="label.mappingPassion"/>
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
                
                <!--  Attribute Description-->
                 <div id="createMappingValUpdate"> 
                <div class="single_form_item">
                   <div class="col-sm-4 select_option function_group">
                   <label for="inputAFD" class="col-sm-12 control-label text_label"><spring:message code="label.attrName"/></label></div>
                   <div class="col-sm-6 select_option_radio"><input type="text" class="form-control-general-form-redefined" maxlength="30" value="" id="attrDescInptId" onkeypress="return alphaNumericInput(event)" autofocus></div> 
                   <div class="col-sm-2 hidden-ipad">&nbsp;</div>  
                   <div class="clearfix"></div>
                </div>
                
                 <div class="single_form_item">
                   <div class="col-sm-4 select_option function_group">
                   <label for="inputAD" class="col-sm-12 control-label text_label"><spring:message code="label.attrDesc"/></label></div>
                   <div class="col-sm-6 select_option_radio"><input type="text" class="form-control-general-form-redefined" maxlength="100" value="" id="attrFullDescInptId" onkeypress="return alphaNumericInput(event)" autofocus></div> 
                   <div class="col-sm-2 hidden-ipad">&nbsp;</div>  
                   <div class="clearfix"></div>
                </div>
                
               <!--  ADDED FOR JCT PUBLIC VERSION -->
               <!--  Global profile change div -->
               <div id= "globalProfileChgDiv" style="display: none;">
                <div class="single_form_item">                
                <div class="col-sm-4 select_option">
                	<label for="inputAD" class="col-sm-12 control-label text_label"><spring:message code="label.make.global.change"/></label>
                </div>
                <div class="col-sm-8 global_chg_checkbox">
                <input type="checkbox" class="" id="globalProfileCheckBoxId"> <span class="global_chg_text"><spring:message code="label.make.global.change.desc"/></span>
                </div>                  
                <div class="col-sm-2 hidden-ipad">&nbsp;</div>  
                   <div class="clearfix"></div>
                </div>
                </div>
                
                
                <div class="single_form_item">
                 <div id ="strengthAddDiv" class="col-sm-6 right_allgn right_allgn_strength">                 
                  <div class="single_form_item">
                     <p><input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Add" id="addStrengthBtnId" onclick="saveStrength()" />                      
                     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                     <input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn reset_class" value="Reset" id="cancelBtnId" onclick="resetValue()" /></p>
                   </div>	
                  </div>
                  
                 <div id ="valueAddDiv" style = 'display: none;' class="col-sm-6 right_allgn right_allgn_strength">
                   <div class="single_form_item">
                     <p><input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Add" id="addValueBtnId" onclick="saveValue()" />
                     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                     <input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn reset_class" value="Reset" id="cancelBtnId" onclick="resetValue()" /></p>
                   </div>	
                 </div>
                 
                   <div id ="passionAddDiv" style = 'display: none;' class="col-sm-6 right_allgn right_allgn_strength">
                   <div class="single_form_item">
                     <p><input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Add" id="addPassionBtnId" onclick="savePassion()" />
                     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                     <input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn reset_class" value="Reset" id="cancelBtnId" onclick="resetValue()" /></p>
                   </div>	
                 </div>
                 
                 <div class="col-sm-3 chgOrderDivButton" id="chgOrderDiv">  
                   <input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn changeOrdBtn_Ipad" value="Change Order" id="chgOrderBtnId" data-toggle="modal" data-target="#myModal" onclick="changeOrder()"/>
                   </div>
                   <div class="clearfix"> </div>
                   
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
			      
			      <div class="modal-footer" id="attrSaveBtnId">
			        <!-- <button type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" data-dismiss="modal">Close</button> -->
			        <button type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" onclick="saveAttributeOrder()">Save</button>
			      </div>
			      			      		
			    </div>
			  </div>
			</div>
			     
	             <div class="existing_data_div">
	             
	             <div class= "existing_list" id="existing_data_list"></div>
	            <div id="existingMappingListId">
	            	<div align="center">
	            		<img src="../img/dataLoading.GIF" />
	            		<p> Loading Existing Attributes  </p>
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
<script src="../js/mappingValue.js"></script>
<script src="../js/common.js"></script>
<script src="../js/jquery.tablesorter.js"></script>
<script src="../js/jquery.tablednd.js"></script>
</body>
</html>