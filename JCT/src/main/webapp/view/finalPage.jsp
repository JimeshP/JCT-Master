<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<head>
    <title><spring:message code="label.result.page" /></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link rel="shortcut icon" href="/user/img/crafting_ico.ico" />
    <!-- Bootstrap -->
    <link href="../css/bootstrap.css" rel="stylesheet">
    <link href="../css/style.css" rel="stylesheet">
    <link href="https://s3-us-west-2.amazonaws.com/jobcrafting/css/commonStyle.css" rel="stylesheet">
    <link rel="stylesheet" href="../css/alertify.core.css" />
	<link rel="stylesheet" href="../css/alertify.default.css" id="toggleCSS" />
	<meta name="viewport" content="width=device-width">
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
	<!-- <script src="http://code.jquery.com/jquery-latest.min.js"></script> -->
	
<script src="http://code.jquery.com/ui/1.10.3/jquery-ui.js" type="text/javascript"></script>
<script type="text/javascript">
//<![CDATA[
(window.jQuery && window.jQuery.ui && window.jQuery.ui.version === '1.10.3')||document.write('<script type="text/javascript" src="../js/jquery_ui.js"><\/script>');//]]>
</script>	
	
	
		
    <!-- <script src="http://code.jquery.com/ui/1.10.3/jquery-ui.js"></script> -->
    <script src="../js/alertify.min.js"></script>
	 <link rel="stylesheet" href="http://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css"> 
</head>
<body>
    <div class="main_warp">
    
     <!-- Header area start -->
        <jsp:include page="header.jsp"/>
     <!-- Header area end --> 
       
     <!-- Menu area start -->
     	<jsp:include page="mainMenu.jsp"/>
     <!-- Menu area end -->   
        
	        <div class="container" id="header-wrap">	          	 
	        <div class="row"> 
             <div class="form_area">
               <div class="form_area_top col-md-12 prev_nxt_sec">
                  <%-- <div class="col-md-2 col-xs-3 Prev_btn_area ">
                    <a href="#" class="btn btn-info prev-btn" onclick="goToResultPage()"><span class="glyphicon glyphicon-chevron-left"></span>&nbsp;&nbsp;&nbsp;<spring:message code="label.finalPageBack"/>&nbsp;&nbsp;&nbsp;</a>
                    <div class="clearfix"></div>
                  </div> --%>
                  <div class="col-md-8 col-xs-12 text-center heading-title">
                   
                  </div>
                 <%--  <div class="col-md-2 ">
                  <a href="#" onclick="goToNextPage()" class="btn btn-info prev-btn pull-right"><spring:message code="label.nextStep"/> <span class="glyphicon glyphicon-chevron-right"></span></a>
                  </div> --%>
                  <div class="clearfix"></div>
               </div>
                                 
               <div class="form_area_middle">
                 <div class="form_area_wraper_justified_page">
                 <form class="col-md-12 pd_l_0 pd_r_0" autocomplete="on" novalidate>                 
                            
               <div class="finalPage-wrapper">
                  <h4 class="finalPage_link"><a href="#" onclick="populateLandingPage()" class="finalPage_desc"><spring:message code="label.finalPage.intoduction"/></a></h4>
                  
                  <div class="col-md-12 finalPage_divider"><spring:message code="label.finalPage.bsHeader"/></div> 
                   <h4 class="finalPage_link"><a href="#" onclick="populateBeforeSketch()" class="finalPage_desc"><spring:message code="label.finalPage.beforeSketch"/></a></h4> 
                   <h4 class="finalPage_link"><a href="#" onclick="populateQuestionnaire()" class="finalPage_desc"><spring:message code="label.finalPage.refQtn"/></a></h4>  
                   
                  <div class="col-md-12 finalPage_divider"><spring:message code="label.finalPage.asHeader"/></div> 
                   <h4 class="finalPage_link"><a href="#" onclick="populateMapping()" class="finalPage_desc"><spring:message code="label.finalPage.mapping"/></a></h4> 
                   <h4 class="finalPage_link"><a href="#" onclick="populateAfterDiag()" class="finalPage_desc"><spring:message code="label.finalPage.afterSketch"/></a></h4>   
                   <h4 class="finalPage_link"><a href="#" onclick="populateActionPlan()" class="finalPage_desc"><spring:message code="label.finalPage.actionPlan"/></a></h4>                              
                   <h4 class="finalPage_link"><a href="#" onclick="showOldPDF()" class="finalPage_desc">Results</a></h4>
                  <div class="clearfix"></div>
                  
                  <div class="col-md-12 finalPage_divider"><spring:message code="label.finalPage.excersize.header"/></div> 
                  <h4 class="finalPage_link"><a href="#" onclick="restartExcersize()" class="finalPage_desc"><spring:message code="label.restart.excersize"/></a></h4>
					<br>
               <%--    <div class="start-btn">
                    <a class="multi-line-button green" href="#" style="width:14em" onClick="submitDecision()">
                      <span class="title"><spring:message code="label.finish.flow"/></span>
                    </a>
                  </div> --%>
               </div>

                </form>
                   <div class="clearfix"></div>
                 </div>
               </div>
               <div class="form_area_bottom"></div>
             </div>
           </div>
        </div>
         <!-- Form area end -->
        <!-- Footer area start -->
        
        <%-- <div class="row-fluid header">
          <h5 class="col-md-5"><spring:message code="label.footer"/></h5>
        </div> --%>
        <jsp:include page="footer.jsp"/>
       
        <!-- Footer area end -->
        <%-- <div class="loader_bg" style="display:none"></div>
        <div class="loader" style="display:none"><img src="../img/Processing.gif" alt="loading"></div>
        <div class="row-fluid footer">
          <h5 class="align_center footer_heading"><spring:message code="label.footer"/></h5>
        </div> --%>
        <!-- Footer area end -->
        <div class="loader_bg" style="display:none"></div>
        <div class="loader" style="display:none"><img src="../img/Processing.gif" alt="loading"></div>
        <!-- Footer area end -->
    </div>
<script src="../js/bootstrap.min.js"></script>
<script src="../js/session_storage.js"></script>
<script src="../js/resultPage.js"></script>
<script src="../js/stopWatch.js"></script>
<script src="../js/common.js"></script>
<script src="../js/finalPage.js"></script>
</body>
</html>