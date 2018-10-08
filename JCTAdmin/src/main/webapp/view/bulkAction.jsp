<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<head>
    <title><spring:message code="menu.sub.create.user" /></title>
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
           				<li class="first_nav"><spring:message code="menu.app.manage.user"/> &nbsp;|</li>
           				<li class="second_nav" id='updateId'><spring:message code="menu.sub.create.user"/></li>        				          				
             		</ol>
	            </div>
	            <div class= "sub_contain"></div>
	            <div class= "main_contain">	  
	             <div class="form_area_top">
	             <div class= "inner_container"><spring:message code="label.bulk.action"/></div> 
	            <form class="form-horizontal main_div" name="bulkAction" autocomplete="on" novalidate enctype="multipart/form-data" data-remote="true"> 	            
	                        
	             <!-- USER TYPE -->
	              <!--  User Group -->
                  <div class="single_form_item">
                     <div class="col-sm-4 ">
                     <label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="label.select.user.type"/></label></div>
                     <div class="col-sm-5">
                     	<select class="form-control-general" id="inputUserTypeInpt" onchange="enableDisableSelection()" disabled>
                     		<option class="form-control-general" value="I"><spring:message code="label.user.type.ind"/> </option>
                     		<option class="form-control-general" value="F"><spring:message code="label.user.type.fac"/> </option>
                     	</select>
                     </div> 
                     <div class="col-sm-3">&nbsp;</div>  
                     <div class="clearfix"></div>
                </div> 
	              
                 <!--  User Group -->
                  <div class="single_form_item">
                     <div class="col-sm-4 ">
                     <label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="label.user.select.group"/></label></div>
                     <div class="col-sm-5">
                     	<select class="form-control-general" id="inputUserGroupInpt" onchange="openRadioDiv()">
                     		<option class="form-control-general" value=""><spring:message code="label.user.select.group_drpDwn"/> </option>
                     	</select>
                     </div> 
                     <div class="col-sm-3">&nbsp;</div>  
                     <div class="clearfix"></div>
                </div>      
               
                <!--  Action  -->
	           	<div class="single_form_item" id='radioDrpDwnList' style="display: none;">
                      <div class="col-sm-4 ">
                     <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="label.user.choose.action"/></label></div>
                     <div class="col-sm-8 sub_text_label">
                     <div class="col-sm-4 radio_text">
                          <span id="activeIdTxt"><input type="radio" name="actionFld" id="activeId" value="active" onclick='decidePlotting("activeId")'><spring:message code="label.user.active"/></span>
                     </div> 
                     <div class="col-sm-4 radio_text">
                          <span id="inctiveIdTxt"><input type="radio" name="actionFld" id="inctiveId" value="inactive" onclick='decidePlotting("inctiveId")'> <spring:message code="label.user.inactive"/></span>
                     </div> 
                      <div class="col-sm-4 radio_text" id='passwordDivId'>
                          <input type="radio" name="actionFld" id="passwordId" value="password" onclick='decidePlotting("passwordId")'> <spring:message code="label.user.reset.password"/>
                     </div> 
                     </div>
                     <div class="clearfix"></div> 
                     
                     <%-- <div class="col-sm-4 "></div>
                     <div class="col-sm-6 sub_text_label">
                          <input type="radio" name="actionFld" id="activeId" value="active" onclick='decidePlotting("activeId")'> <spring:message code="label.user.active"/>
                     </div> 
                     <div class="col-sm-2">&nbsp;</div>  
                     <div class="clearfix"></div>
                     
                     <div class="col-sm-4 "></div>
                     <div class="col-sm-6 sub_text_label">
                          <input type="radio" name="actionFld" id="inctiveId" value="inactive" onclick='decidePlotting("inctiveId")'> <spring:message code="label.user.inactive"/>
                     </div> 
                     <div class="col-sm-2">&nbsp;</div>  
                     <div class="clearfix"></div>
                     
                     <div class="col-sm-4 "></div>
                     <div class="col-sm-6 sub_text_label">
                          <input type="radio" name="actionFld" id="passwordId" value="password" onclick='decidePlotting("passwordId")'> <spring:message code="label.user.reset.password"/>
                     </div> 
                     <div class="col-sm-2">&nbsp;</div>  
                     <div class="clearfix"></div> --%>
                   </div>
                   
                   <div class="single_form_item"><br></div> 
                                                                             
                </form>
	            </div>
	            	            
	             <div class="existing_data_div" id='existing_data_div_id' style="display: none;">	             
	             <div class= "existing_list"></div>
	            	 <div id="mainPopulationDiv" style="display: none;">
	            	
	            	 </div>
	            	 <div id='buttonDivId' style="display: none;">
	            	 
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
		<div class="loader_bg" style="display:none"></div>
	    <div class="loader" style="display:none"><img src="../img/Processing.gif" alt="loading"></div>
		
<script src="https://code.jquery.com/jquery.js"></script>    
<!-- Include all compiled plugins (below), or include individual files as needed -->
<script src="../js/bootstrap.min.js"></script>
<script src="../js/common.js"></script>
<script src="../js/bulkAction.js"></script>
<script src="../js/jquery.tablesorter.js"></script>
</body>
</html>