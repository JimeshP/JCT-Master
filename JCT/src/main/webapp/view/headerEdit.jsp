<!DOCTYPE html>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<html>
<head>
    <title>..: Header :..</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <!-- <meta name="viewport" content="width=device-width, initial-scale=1.0"> -->
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
    <!-- <meta name="viewport" content="width=device-width" /> -->
    <!-- Bootstrap -->
    <link href="../css/bootstrap.css" rel="stylesheet"> <!-- Bootstrap framework css  -->
    <link href="../css/style.css" rel="stylesheet"> <!-- Main stylesheet for every resolution -->  
    <link href="../css/less-600.css" rel="stylesheet"> <!-- mainly effected for google Nexus & less than 601 resolution -->
    <link href="../css/less-550.css" rel="stylesheet"> <!-- mainly effected for Htc & less than 560 resolution -->
    <link href="../css/less-768.css" rel="stylesheet"> <!-- mainly effected for S3 & less than 768 resolution --> 
    <link href="../css/less-1280.css" rel="stylesheet"> <!-- mainly effected for Ipad & less than 1280 resolution -->
	<link href="../css/style-1152.css" rel="stylesheet"> <!-- mainly effected for less than 1152-1270 resolution -->
    <link href="../css/iphone_specific.css" rel="stylesheet"> <!-- mainly effected for Iphone & less than 481 resolution -->  
    <script type="text/javascript" src="/user/js/common.js"></script>  
</head>

<body> 	
	<!-- Header area start -->
        <div class="container-fluid header">
            <div class="header_wrap_area container">
            <div class="row">
            <div class="col-xs-12 col-md-4 logo_area">
            	<h5 style="display: none;" id="headerLogoId" class="heading_main logo">&nbsp;</h5>
            </div>
            
            
            <div class="col-xs-12 col-md-4 tool_title">
            <spring:message code="label.job.crafting.exercise"/></div>
            
           <div class="col-xs-12 col-md-3 welcome_area">
              
              <%-- <a href="#" id="fancybox-manual-my-account" onclick="openMyAccountDetails()"><spring:message code="label.menu.myaccount"/></a> &nbsp;&nbsp;| &nbsp;&nbsp;<a href="#" id="logMeOut"><spring:message code="label.logOut"/></a> --%>
              </div>   
             <div class="clearfix"></div>
              </div>              
            </div>
        </div>
        <!-- Header area end -->
</body>
</html>