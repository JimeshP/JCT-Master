<!DOCTYPE html>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<html>
<head>
    <title>..: Header :..</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <!-- Bootstrap -->
    <link href="css/bootstrap.css" rel="stylesheet">
    <link href="css/style.css" rel="stylesheet">  
    <script type="text/javascript" src="../js/common.js"></script>
</head>

<body> 
	<!-- <div class="row-fluid header">
    	<div class="header_wrap_area container" >
     		<div class="col-sm-3">
     			<h5 class="heading_main logo">&nbsp;</h5>
     		</div>
        	<div class="col-sm-9">
       			<h5 class="heading_main"></h5>         
			</div>              
     	</div>
	</div>     -->
	
	<!-- Header area start -->
        <div class="row-fluid header">
            <div class="header_wrap_area container" >
            <div class="col-sm-3">
            	<h5 class="heading_main logo" onclick="window.open('http://www.jobcrafting.com','_blank');">&nbsp;</h5>
            </div>
              <div class="col-sm-9 pull-right time_section">

             <div class="col-sm-3 total_time" style="display: none;">&nbsp;<span id="timeId"></span></div> 

              <div class="col-sm-5">
              </div>

              <div class="col-sm-4 welcome_area pull-right"><spring:message code="label.greetUser"/> <span id="nameId"></span> | <a href="#" id="log_out" onclick="logoutGeneral('landingPage')"><spring:message code="label.logOut"/></a></div>             
			  </div>              
            </div>
        </div>
        <!-- Header area end -->
</body>
</html>