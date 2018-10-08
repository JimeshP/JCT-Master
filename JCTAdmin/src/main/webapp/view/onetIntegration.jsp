<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@ taglib uri="http://java.sun.com/jsp/jstl/core" prefix="c" %>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<head>
    <title><spring:message code="menu.sub.create.onet.occupation" /></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <!-- Bootstrap -->
    
    <link href="/admin/css/bootstrap.css" rel="stylesheet">
    <link href="/admin/css/style.css" rel="stylesheet">
    <link rel="stylesheet" href="/admin/css/alertify.core.css" />
	<link rel="stylesheet" href="/admin/css/alertify.default.css" id="toggleCSS" />
	<meta name="viewport" content="width=device-width">
	<link rel="shortcut icon" href="/admin/img/crafting_ico.ico" />
    
	<script type="text/javascript" src="/admin/lib/jquery.js"></script>
	<script type="text/javascript" src="/admin/lib/spine.js"></script>
	<script type="text/javascript" src="/admin/lib/ajax.js"></script>
	<script src="http://code.jquery.com/jquery-latest.min.js"></script>
	<script>
	if (typeof jQuery == 'undefined') {
    	document.write(unescape("%3Cscript src='/admin/js/latest_10.2_jquery.js' type='text/javascript'%3E%3C/script%3E"));
    }
	</script>	
	<script src="http://code.jquery.com/ui/1.10.3/jquery-ui.js" type="text/javascript"></script>
	<script type="text/javascript">
	//<![CDATA[
        (window.jQuery && window.jQuery.ui && window.jQuery.ui.version === '1.10.3')||document.write('<script type="text/javascript" src="/admin/js/jquery_ui.js"><\/script>');//]]>
    </script>	
	<script src="/admin/js/alertify.min.js"></script>
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
	            
	            <div class="col-sm-10 col">
	            <div class= "heading_label">
	            	<ol class="breadcrumbAdmin">
           				<li class="first_nav"><spring:message code="reports.breadcrumb.manage.user.home"/> &nbsp;|</li>
           				<li class="first_nav"><spring:message code="menu.cont.configs"/> &nbsp;|</li>
           				<li class="second_nav"><spring:message code="menu.sub.create.onet.occupation"/></li>
             		</ol> 
	            </div>
	             <div class= "sub_contain"><spring:message code="menu.sub.create.onet.occupation"/></div>
	            <div class= "main_contain">	  
	            <div class= "inner_container"><spring:message code="menu.sub.create.onet.occupation"/></div>	                 
	            
	            
	            <form class="form-horizontal main_div" modelAttribute="uploadForm" method="POST" action="/admin/uploadFile" enctype="multipart/form-data" onsubmit="return validateFile()">
       				<br>
       				<div class="single_form_item">
       					<div class="col-sm-4 ">
                      	<label for="inputOC" class="col-sm-12 control-label text_label"><spring:message code="label.upload.occupation"/></label></div>	
                       	<div class="col-sm-5">
                     		<input type="file" class="filestyle" data-classButton="btn choose_file" name="file" id="filename" accept=".xls,.xlsx ">
                     	</div>  
                     	<div class="col-sm-3">&nbsp;</div>  
                     	<div class="clearfix"></div>  
       			 	</div>
       			 
       			 	<div class="single_form_item">
                    	 <p><input type="submit" id= "onetDataSaveBtnId" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Save"  /></p>
                   	</div>
                </form>  
   			     			
	            
	            <div class="existing_data_div">             
	             <div  id="existing_list_Id" class= "existing_list"></div>
	             <div id="onetOcuupationListId" class="table_div_scroll">
	            	<div align="center">
	            		<img src="/admin/img/dataLoading.GIF" />
	            		<p> Loading Existing Onet Occupation Data </p>
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
 <script src="/admin/js/bootstrap.min.js"></script>
<script src="/admin/js/bootstrap-filestyle.js"></script>
<script src="/admin/js/common.js"></script>
<script src="/admin/js/onetIntegration.js"></script>
<script src="../js/jquery.tablesorter.js"></script>
<script src="/admin/js/equalHeight.js"></script>
<script type="text/javascript">
$(":file").filestyle();
</script>
<script type="text/javascript">
var msg = '<c:out value="${message}" />';
if(msg != "") {
	alertify.alert(msg);
}
</script>	
</body>
</html>