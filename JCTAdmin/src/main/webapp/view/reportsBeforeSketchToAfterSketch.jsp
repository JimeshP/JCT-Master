<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<head>
    <title><spring:message code="label.reportBSToAS" /></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <!-- Bootstrap -->
    
    <link href="../css/bootstrap.css" rel="stylesheet">
    <link href="../css/style.css" rel="stylesheet">
    <link rel="stylesheet" href="../css/alertify.core.css" />
	<link rel="stylesheet" href="../css/alertify.default.css" id="toggleCSS" />
	<link rel="stylesheet" type="text/css" href="../css/jquery.fancybox.css" media="screen" />
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
	           <!--  <div class="col-sm-12 container container_mid" > -->
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
           				<li class="second_nav"><spring:message code="reports.breadcrumb.bs.asbs"/></li>
             		</ol>
	            </div>
	            
	            <div class= "sub_contain"><spring:message code="label.beforeAfterSketchRprt"/></div>
	            
	            <div class= "main_contain">	  
	             <div class="form_area_top" style="display: none;">
	             <div class= "inner_container"><spring:message code="label.searchBeforeAfterSketchRprt"/></div> 
                 <form class="form-horizontal main_div" name="beforeSktchRprt" autocomplete="on" novalidate enctype="multipart/form-data" data-remote="true">       
	            
	            <!-- OCCUPATION SEARCH -->
	            <div class="single_form_item">
	                   <div class="col-md-4 ">
	                   	   <label for="inputFN" class="col-md-12 control-label align-right mandatory_field"><spring:message code="label.occupation.search"/></label>
	                   </div>
	                   <div class="col-md-4"><input type="text" class="form-control-general" maxlength="30"  placeholder="Enter key words describing occupation" value="" id="ocptnSrchTxtId" oninput="validateOnetData()"></div> 
	                   <div class="col-md-4"><input type="button" class="btn_admin btn_admin-primary btn_admin-sm" value="Search Occupation" id="searchOccupation" onclick="searchOccupationSearch()" /></div>  
	                   <div class="clearfix"></div>
	               </div>
	               
	            <!-- FUNCTION GROUP -->
                   <%-- <div class="single_form_item">
                     <div class="col-sm-4 ">
                     <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="label.functionGrp"/></label></div>
                     <div class="col-sm-5">
                     	<select class="form-control-general" id="inputFunctionGroup">
                     		<option class="form-control-general" value=""><spring:message code="label.select.functionGrp"/> </option>
                     	</select>
                     </div> 
                     <div class="col-sm-3">&nbsp;</div>  
                     <div class="clearfix"></div>
                   </div> --%>
                 
                 <!-- JOB LEVEL -->
                   <%-- <div class="single_form_item">
                     <div class="col-sm-4 ">
                     <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="label.jobLevel"/></label></div>
                     <div class="col-sm-5">
                     	<select class="form-control-general" id="inputJobLevel">
                     		<option class="form-control-general" value=""><spring:message code="label.select.jobLevel"/> </option>
                     	</select>
                     </div> 
                     <div class="col-sm-3">&nbsp;</div>  
                     <div class="clearfix"></div>
                   </div> --%>
                   <div class="single_form_item">
                     <p><input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Search Report" id="searchBtnBSToAS" onclick="populateBeforeSketchToAfterSketchSelectedReport('newid')" /></p>
                   </div>	
	             </form>
	             </div>
	            
	            <div class="result_area">	            
	              <div class="col-sm-7 result_text" id="beforeSketchResultDiv"><spring:message code="label.resultBeforeAfterSketchRprt"/></div>  
	              <div class="col-sm-4 inner_container" id="genearteXLDiv"></div>  
	              <div class="clearfix"></div> 
	              	              
	              <div class="col-sm-6 display_value pd-l-7" id="beforeToafterSketchRptHdrDiv"></div>
	              <div class="col-sm-5"></div>
	              <div class="clearfix"></div> 
	              </div>
	             <!-- <div class="main_div main_header_div" id="beforeToafterSketchRptHdrDiv">
	             
	             </div> -->
	            <!--  <div class="table_contant main_div" id="beforeToafterSketchRptDiv">
	             
	             </div>  -->  
	             
	              <div class="main_div table_contant pd-r-0" >
	              	 <div id="beforeToafterSketchRptDiv"></div>
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
		<div id='occupationDiv' class="scroll">
	<div id="headerPanel" class="popupHeader custom_header">
		Select the option that best describes your occupation. Feel free to try different search words to find a better fit.
	</div>
	<div align="center" id="occupationPlottingDiv">
		<br /><br />
		<img src="../img/Processing.gif" />
	</div>
	<div align="center" id="dataPlottingDiv" style="height: 82%; overflow: auto;">
	</div>
	<div align="center" style="display: none;" id="goDiv">
		<input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Go" onclick="closeDiv()" />
	</div>
</div>
		<script src="https://code.jquery.com/jquery.js"></script>    
<!-- Include all compiled plugins (below), or include individual files as needed -->
<script src="../js/bootstrap.min.js"></script>
<script src="../js/beforeToafterSketchReport.js"></script>
<script src="../js/menuHandler.js"></script>
<script src="../js/common.js"></script>
<script src="../js/equalHeight.js"></script>
<script type="text/javascript" src="../js/jquery.fancybox.pack.js"></script>
	<script type="text/javascript">
		$(document).ready(function() {
			$("#searchOccupation").click(function() {
				$.fancybox.open({
					href : '#occupationDiv',
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
</body>
</html>