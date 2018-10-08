<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<head>
    <title><spring:message code="menu.sub.create.myAccount" /></title>
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
	            <div class="main-cont-wrapper" >
	            <div class="col-sm-2 pd-0">
	            	<!-- Menu area start -->
	    			<jsp:include page="sideMenuNew.jsp"/>
	   				 <!-- Menu area end -->   
	            </div>
	            
	            <div class="col-sm-10">
	            <div class= "heading_label">
	            	<ol class="breadcrumbAdmin">
           				<li class="first_nav"><spring:message code="reports.breadcrumb.manage.user.home"/> &nbsp;|</li>
           				<li class="second_nav"><spring:message code="menu.sub.create.myAccount"/></li>
             		</ol>
	            </div>
	            <br>
	            <div class= "sub_contain"></div>
	            <div class= "main_contain">	  
	            <form class="form-horizontal" name="actionPlanRprt" autocomplete="on" novalidate enctype="multipart/form-data" data-remote="true"> 
	           <div class="single_form_item"><br/></div>
	            <!-- First Name -->
	            <div class="single_form_item">	            	
                     <div class="col-sm-4 ">
                      <label for="inputFN" class="col-sm-12 control-label text_label_myAccnt left_text"><spring:message code="label.MyAccnt.fName"/></label></div>
                     <div class="col-sm-3"><output class="output_label output_fac_Id" id="firstNameInptId"></output></div> 
                     <div class="col-sm-2"></div> 
                     <div class="col-sm-3"></div>  
                     <div class="clearfix"></div>
                  </div>
                  
                  
                  <div id="firstNameChgMainDiv" class="input_chg_div" style="display: none">
                   <div id="firstNameChgDiv" class="input_chg">
                   	<p id="firstNameChgText" class="input_chg_label"><spring:message code="label.MyAccnt.enter.first.name"/></p>
					<input type="text" id="firstNameChgInptId" class="form-control-general-form-redefined" maxlength="50">
                    <div class="input_chg_button_div">
                   <input type="button" value="OK" class="input_chg_button" id="firstNameChgOkBtnId">
                   <input type="button" value="Cancel" class="input_chg_button input_chg_button_cancel" id="firstNameChgCancelBtnId">
                   </div>
                   </div>
                  </div>
                  
                  <div id="divGeneralId" class="general_div" style="display: none;"></div>
                 
                 <!-- Last Name -->
	            <div class="single_form_item">	            	
                     <div class="col-sm-4 ">
                      <label for="inputFN" class="col-sm-12 control-label text_label_myAccnt left_text"><spring:message code="label.MyAccnt.lName"/></label></div>
                     <div class="col-sm-5"><output class="output_label" id="lastNameInptId"></output></div> 
                     <div class="col-sm-3"></div>  
                     <div class="clearfix"></div>
                  </div>
                  
                  <div id="lastNameChgMainDiv" class="input_chg_div" style="display: none">
                   <div id="lastNameChgDiv" class="input_chg">
                   	<p id="lastNameChgText" class="input_chg_label"><spring:message code="label.MyAccnt.enter.last.name"/></p>
					<input type="text" id="lastNameChgInptId" class="form-control-general-form-redefined" maxlength="50">
                    <div class="input_chg_button_div">
                   <input type="button" value="OK" class="input_chg_button" id="lastNameChgOkBtnId">
                   <input type="button" value="Cancel" class="input_chg_button input_chg_button_cancel" id="lastNameChgCancelBtnId">
                   </div>
                   </div>
                  </div>
                  
                  <div id="divGeneralId" class="general_div" style="display: none;"></div>
                  
                  <!-- Password -->
	            <div class="single_form_item">	            	
                     <div class="col-sm-4 ">
                      <label for="inputFN" class="col-sm-12 control-label text_label_myAccnt left_text"><spring:message code="label.MyAccnt.pwd"/></label></div>
                     <div class="col-sm-5"><output class="output_label" id="chgPwdInptId"></output></div> 
                     <div class="col-sm-3"></div>  
                     <div class="clearfix"></div>
                  </div>
                  
                  <div id="pwdChgMainDiv" class="input_chg_div" style="display: none">
                   <div id="pwdChgDiv" class="input_chg">
                   	<p id="pwdChgText" class="input_chg_label"><spring:message code="label.MyAccnt.enter.old.pwd"/></p>
					<input type="password" id="oldPwdChgInptId" class="form-control-general-form-redefined" maxlength="30">
					<p id="pwdChgText" class="input_chg_label"><spring:message code="label.MyAccnt.enter.new.pwd"/></p>
					<input type="password" id="newPwdChgInptId" class="form-control-general-form-redefined" maxlength="30">
					<p id="pwdChgText" class="input_chg_label"><spring:message code="label.MyAccnt.enter.confirm.pwd"/></p>
					<input type="password" id="confirmPwdChgInptId" class="form-control-general-form-redefined" maxlength="30">
                    <div class="input_chg_button_div">
                   <input type="button" value="OK" class="input_chg_button" id="pwdChgOkBtnId">
                   <input type="button" value="Cancel" class="input_chg_button input_chg_button_cancel" id="pwdChgCancelBtnId">
                   </div>
                   </div>
                  </div>
                  
                  <div id="divGeneralId" class="general_div" style="display: none;"></div>   
                  
                   <!-- Email Id -->
	            <div class="single_form_item">	            	
                     <div class="col-sm-4 ">
                      <label for="inputFN" class="col-sm-12 control-label text_label_myAccnt left_text"><spring:message code="label.MyAccnt.email"/></label></div>
                     <div class="col-sm-5"><output class="output_label" id="emailIdInptId"></output></div> 
                     <div class="col-sm-3"></div>  
                     <div class="clearfix"></div>
                  </div>
                          
                  <div id="emailIdChgMainDiv" class="input_chg_div" style="display: none">
                   <div id="emailIdChgDiv" class="input_chg">
                   	<p id="emailIdChgText" class="input_chg_label"><spring:message code="label.MyAccnt.enter.email.id"/></p>
					<input type="text" id="emailIdChgInptId" class="form-control-general-form-redefined" maxlength="30">
                    <div class="input_chg_button_div">
                   <input type="button" value="OK" class="input_chg_button" id="emailIdChgOkBtnId">
                   <input type="button" value="Cancel" class="input_chg_button" id="emailIdChgCancelBtnId">
                   </div>
                   </div>
                  </div>
                  
                  <div id="divGeneralId" class="general_div" style="display: none;"></div>       
                  
                  <div class="single_form_item"><br/></div>
                   <div class="single_form_item"><br/></div>
                   
                   <!-- <div class="single_form_item">
                     <p><input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Cancel" id="cancel" /></p>
                   </div> -->
                   
                   <div class="single_form_item"><br/></div>
                   <div class="single_form_item"><br/></div>
                   <div class="single_form_item"><br/></div>
	            </form>

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
<script src="../js/myAccount.js"></script>
<script src="../js/common.js"></script>
<script src="../js/sortTable.js"></script>
</body>
</html>