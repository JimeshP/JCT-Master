<%@page contentType="text/html" pageEncoding="UTF-8"%>    
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>    
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title><spring:message code="label.survey.rpt" /></title>
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
	            <div class="col-sm-10 col">
	            
	             <div class= "heading_label">
	            	<ol class="breadcrumbAdmin">
           				<li class="first_nav"><spring:message code="reports.breadcrumb.bs.home"/> &nbsp;|</li>
           				<li class="first_nav"><spring:message code="reports.breadcrumb.bs.rpts"/> &nbsp;|</li>
           				<li class="second_nav">Survey Reports</li>
             		</ol>
	            </div>
	              <div class= "sub_contain">Survey Reports</div>
	              
	          <div class= "main_contain">	  
	             <div class="form_area_top">
	             <div class= "inner_container">Search Facilitator/Individual Survey Reports</div> 
                 <form class="form-horizontal main_div" name="beforeSktchRprt" autocomplete="on" novalidate enctype="multipart/form-data" data-remote="true">
                 
                 <!-- User Profile -->
                 <br>
                   <div class="single_form_item">
                     <div class="col-sm-4 ">
                     <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="label.fac.ind"/></label></div>
                     <div class="col-sm-5">
                     	<select class="form-control-general" id="inptUserType">
                     		<option class="form-control-general" value=""><spring:message code="label.user.type.select.drpDwn"/> </option>
                     		<option class="form-control-general" value="0">All </option>
                     		<option class="form-control-general" value="1"><spring:message code="label.user.type.select.drpDwn.ind"/> </option>
                     		<option class="form-control-general" value="3"><spring:message code="label.user.type.select.drpDwn.fac"/> </option>
                     	</select>
                     </div> 
                     <div class="col-sm-3">&nbsp;</div>  
                     <div class="clearfix"></div>
                </div>  
                <div class="single_form_item">
                    <p><a id='generateXL' href='#' class="btn_admin btn_admin-primary btn_admin-sm search_btn cstm-text-format" onclick='generateXLReport(event)'>Generate Excel</a></p>
                </div> 
                
                 <!--   <div class="single_form_item">
                   <p>
					<input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Create" id="addNewUserBtnId" onclick="generateReport()" />
				   </p>
                 </div>-->
                  
              </form>
              </div>
           <!--    <div class="result_area">   		
                  <div class="col-sm-7 result_text" id="beforeSketchResultDiv">Results for Survey Reports</div> 
                  <div class="col-sm-4 inner_container" id="genearteXLDiv"></div> 
                  <div class="clearfix"></div> 
                              
	             <div class="main_div table_contant pd-r-0" style="overflow:auto !important;">
	              	 <div id="surveyRptDiv"></div>
	             </div>
                </div>-->
                            
              </div>
	            </div>
	            <div class="clearfix"></div>
	      </div>
	     </div>	
	         <!-- Footer area start -->
	    <jsp:include page="footer.jsp"/>
	    <!-- Footer area end -->                    
 </div>
 
 <script src="https://code.jquery.com/jquery.js"></script>    
<!-- Include all compiled plugins (below), or include individual files as needed -->
<script src="../js/bootstrap.min.js"></script>
<script src="../js/surveyReport.js"></script>
<script src="../js/menuHandler.js"></script>
<script src="../js/common.js"></script>
<script src="../js/equalHeight.js"></script>
</body>
</html>