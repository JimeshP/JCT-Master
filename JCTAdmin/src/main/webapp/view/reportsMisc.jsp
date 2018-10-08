<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<head>
    <title><spring:message code="label.reportmisc" /></title>
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
           				<li class="first_nav"><spring:message code="reports.breadcrumb.bs.home"/> &nbsp;|</li>
           				<li class="first_nav"><spring:message code="reports.breadcrumb.bs.rpts"/> &nbsp;|</li>
           				<li class="second_nav"><spring:message code="reports.breadcrumb.bs.misc"/></li>
             		</ol>
	            </div>
	            <div class= "sub_contain"><spring:message code="label.miscRprt"/></div>
	            <div class= "main_contain">	  
	             <div class="form_area_top">
                 <form class="form-horizontal" name="actionPlanRprt" autocomplete="on" novalidate enctype="multipart/form-data" data-remote="true">       
	            
	            <div class="header_descriptn" >
	            <div id='noDataDiv' align='center' style='display: none;'><br /><br /><br /><img src='../img/no-record.png'><br /><div class='textStyleNoExist'>No Records found</div></div>
	            <div id='mainHeaderDivCont'></div>
	            <br>
	            <div id='miscHeaderContentDiv'>
	            
	            </div>
                 
                 
                 <div id="detailedDiv" style="display: none">
                 	<div id="detailedHdr" style="width: 100%;"></div>
                 	 <br>
                 	<div id="detailedDivContent">
                 	</div>
                 </div>
                 </div>
                   	
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
		
		</div>
		<div id='detailedDiv' class="scroll">
		
		</div>
		<script src="https://code.jquery.com/jquery.js"></script>    
<!-- Include all compiled plugins (below), or include individual files as needed -->
<script src="../js/bootstrap.min.js"></script>
<script src="../js/miscReport.js"></script>
<script src="../js/menuHandler.js"></script>
<script src="../js/common.js"></script>
<script src="../js/equalHeight.js"></script>
<script src="../js/sortTable.js"></script>
</body>
</html>