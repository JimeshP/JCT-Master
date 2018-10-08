<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<head>
    <title><spring:message code="menu.diagram.remove"/></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <!-- Bootstrap -->
    
    <link href="../css/bootstrap.css" rel="stylesheet">
    <link href="../css/style.css" rel="stylesheet">
    <meta name="viewport" content="width=device-width">
	<link rel="shortcut icon" href="/admin/img/crafting_ico.ico" />
    <!-- HTML5 Shim and Respond.js IE8 support of HTML5 elements and media queries -->
    <!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
      <script src="https://oss.maxcdn.com/libs/respond.js/1.3.0/respond.min.js"></script>
    <![endif]-->
	<script src="https://code.jquery.com/jquery.js"></script>
	<script src="../js/bootstrap.min.js"></script>
    <script type="text/javascript" src="../lib/jquery.js"></script>
	<script type="text/javascript" src="../lib/spine.js"></script>
	<script type="text/javascript" src="../lib/ajax.js"></script>
	
	<link href="../css/bootstrap.css" rel="stylesheet">
	
    <link href="../css/style.css" rel="stylesheet">  
    <script type="text/javascript" src="../js/jquery.fancybox.pack.js"></script>
    <script type="text/javascript" src="../js/iscroll.js"></script>
	<link rel="stylesheet" type="text/css" href="../css/jquery.fancybox.css" media="screen" />
	<script src="../js/alertify.min.js"></script>
	<link rel="stylesheet" href="../css/alertify.core.css" />
	<link rel="stylesheet" href="../css/alertify.default.css" id="toggleCSS" />
	
<!-- Include all compiled plugins (below), or include individual files as needed -->

<script src="../js/common.js"></script>
<script src="../js/removeDiagram.js"></script>
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
           				<li class="first_nav"><spring:message code="menu.cont.configs"/> &nbsp;|</li>
           				<li class="second_nav"><spring:message code="menu.diagram.remove"/></li>
             		</ol> 
	            </div>
	             <div class= "sub_contain"><spring:message code="menu.diagram.remove"/></div>
	            <div class= "main_contain">	  
	            <div class= "inner_container" id="addUserProfileText"><spring:message code="menu.diagram.remove"/></div>
	            <form class="form-horizontal main_div" name="actionPlanRprt" autocomplete="on" novalidate enctype="multipart/form-data" data-remote="true"> 
	            <div class="single_form_item"><br/></div>
	            <div class="single_form_item">
                     <div class="col-sm-4 ">
                      <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="label.userName"/></label></div>
                     <div class="col-sm-6"><input type="text" class="form-control-general-form" maxlength="50" value=""  id="userName" onkeypress="" autofocus></div> 
                     <div class="col-sm-2">&nbsp;</div>  
                     <div class="clearfix"></div>
                   </div>
                   <div class="single_form_item">
                     <p><input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Search" onclick="searchUserDiagram()" /></p>
                   </div>	
                   <div class="single_form_item"><br/></div>
	            </form>
	            	           
	        
	            
	            <div class="existing_data_div" id="diagramPlottingDiv" style="display:none;">             
	            <div id="existing_Diagram_Id" class= "existing_list"><spring:message code="menu.existing.diagram"/></div>
	            
	             <div id="totalDiagramCountId" class= "total_diagram_note" style = "display:block;"></div>
	             
	            <div id="existingUserDiagramId" class= "existingUserDiagramId">
	            	<div align="center">
	            		<img src="../img/dataLoading.GIF" />
	            		<p>Loading Existing Diagrams </p>
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
		<div id="diagramDiv"></div>
		<div id="diagramDiv2"></div>
</body>
<script type="text/javascript">
		$(document).ready(function() {
			$("#dispId").click(function() {
				$.fancybox.open({
					href : '#dd',
					//type : 'iframe',
					closeClick : false,
					openEffect : 'elastic',
					openSpeed  : 150,
					closeEffect : 'elastic',
					closeSpeed  : 150,
					padding : 5
				});
			});
		});
		$(document).ready(function() {
			$("#dispId2").click(function() {
				$.fancybox.open({
					href : '#dd',
					//type : 'iframe',
					closeClick : false,
					openEffect : 'elastic',
					openSpeed  : 150,
					closeEffect : 'elastic',
					closeSpeed  : 150,
					padding : 5
				});
			});
		});
</script>

</html>