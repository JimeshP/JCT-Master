<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<head>
    <title>..: Header :..</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <!-- <meta name="viewport" content="width=device-width, initial-scale=1.0"> -->
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
    <!-- <meta name="viewport" content="width=device-width" /> -->
    <!-- Bootstrap -->
  <!--   <link href="../css/bootstrap.css" rel="stylesheet"> Bootstrap framework css 
    <link href="css/style.css" rel="stylesheet"> Main stylesheet for every resolution  --> 
    <link href="/admin/css/style_ipad.css" rel="stylesheet"> <!-- mainly effected for Ipad & less than 1280 resolution -->   
    <script type="text/javascript" src="/admin/js/common.js"></script>
</head>

<body>

	<!-- Header area start -->
        <div class="container-fluid header">
            <div class="header_wrap_area container">
            <div class="row">
            <div class="col-sm-4 logo_area">
            	<h5 class="heading_main logoGeneral" onclick="window.open('http://www.jobcrafting.com','_blank');">&nbsp;</h5>
            </div>
            <div class="col-sm-4 tool_title" id="interfaceTitle">
            	<%-- <spring:message code="label.job.crafting.panel"/> --%>
            </div>
        
            
           <div class="col-sm-4 welcome_area pull-right align-right">
              <spring:message code="label.greetUser"/> <span id="nameId"></span>&nbsp;&nbsp;| &nbsp;&nbsp;<a href="#" id="logMeOut" onclick="logout()"><spring:message code="label.logOut"/></a>
              </div>   
                        
              </div>              
            </div>
        </div>
     <!-- Header area end -->
</body>
</html>