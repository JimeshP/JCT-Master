<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<head>
    <title><spring:message code="menu.sub.create.user.group" /></title>
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
           				<li class="first_nav"><spring:message code="menu.app.manage.user"/> &nbsp;|</li>
           				<li class="second_nav" id='updateId'><spring:message code="menu.sub.create.user.group"/></li>
             		</ol>
	            </div>
	            <div class= "sub_contain"><spring:message code="menu.sub.user.group"/></div>
	            <div class= "main_contain">
	            <div class= "inner_container" id="addUserGrpText"><spring:message code="label.add.group"/></div>	  
	            <form class="form-horizontal main_div" name="actionPlanRprt" autocomplete="on" novalidate enctype="multipart/form-data" data-remote="true" id="ugForm"> 
	            <div class="single_form_item"><br/></div>
	            
	            <div id="createUserGroupUpdate"> 
	            <div class="single_form_item">
                     <div class="col-sm-4 ">
                     <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="label.user.select.profile"/></label></div>
                     <div class="col-sm-5">
                     	<select class="form-control-general" id="inputUserProfileInpt" onchange="changeUserProfile(this)">
                     		<option class="form-control-general" value=""><spring:message code="label.user.select.profile_drpDwn"/> </option>
                     	</select>
                     </div> 
                     <div class="col-sm-3">&nbsp;</div>  
                     <div class="clearfix"></div>
                </div>  
                </div>
                
	            <div class="single_form_item">
                   <div class="col-sm-4 ">
                   <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="label.user.group"/></label></div>
                   <div class="col-sm-6"><input type="text" class="form-control-general-form-redefined" maxlength="50" value=""  placeholder="User Group Name" id="userGroupInptId" onkeypress="return alphaNumericOnly(event)" autofocus></div> 
                   <div class="col-sm-2">&nbsp;</div>  
                   <div class="clearfix"></div>
                </div>
                                                               
                 <div class="single_form_item">
                   <p><input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Create" id="createUserGroupBtnId" onclick="saveUserGroup()" /></p>
                 </div>	
                   <div class="single_form_item"><br/></div>
	            </form>
	            
	           <%--  
	             <div class= "existing_list"><spring:message code="label.user.existingGroup"/></div>
	            <div id="existingUserGroupListId" class="table_contant">
	            	<div align="center">
	            		<img src="../img/dataLoading.GIF" />
	            		<p> Loading Existing User Groups </p>
	            	</div>
	            </div> --%>
	            
	               <div class="existing_data_div" id="progessDataDiv">             
	             <div id="existing_list_Id" class= "existing_list"><spring:message code="label.user.existingGroup"/></div>
	            <div id="existingUserGroupListId">
	            	<div align="center">
	            		<img src="../img/dataLoading.GIF" />
	            		<p> Loading Existing User Groups </p>
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
<script src="../js/userGroup.js"></script>
<script src="../js/common.js"></script>
<script src="../js/jquery.tablesorter.js"></script>
</body>
</html>