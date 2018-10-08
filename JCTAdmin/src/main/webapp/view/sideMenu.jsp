<!DOCTYPE html>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<html>
<head>
    <title></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <!-- Bootstrap -->
    <link href="../css/bootstrap-combined.min.css" rel="stylesheet">  
    <link href="../css/style.css" rel="stylesheet">       
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
</head>

<body> 
	<ul class="dropdown-menu main_menu col" role="menu" >
	
	
  	<li class="dropdown-submenu"> 	
    <a href="#" class="first-level content_conf"><span class="content_conf_text"><spring:message code="menu.app.manage.user" /></span></a>  
	   	<ul class="dropdown-menu">
		    <li><a href="userProfile.jsp" class= "second-level"><spring:message code="menu.sub.create.user.profile" /></a></li>
		    <li class="sub-divider"></li>
		    <li><a href="userGroup.jsp" class= "second-level"><spring:message code="menu.sub.create.user.group" /></a></li>
		    <li class="sub-divider"></li>
		    <li><a href="addUser.jsp" class= "second-level"><spring:message code="menu.sub.create.user.list" /></a></li>
	    </ul>
  	</li>
  <!-- <li class="dropdown-submenu">
    <a tabindex="-1" href="#" class="first-level">Reports</a>
    <ul class="dropdown-menu">
      <li><a href="#" onclick="populateReports('BS')" class="second-level">Before Sketch</a></li>
      <li class="sub-divider"></li>     
      <li><a href="#" onclick="populateReports('AS')" class="second-level">After Sketch</a></li>
      <li class="sub-divider"></li>
      <li><a href="#" onclick="populateReports('BSTOAS')" class="second-level">Before Sketch To After Sketch</a></li>
      <li class="sub-divider"></li>
      <li><a href="#" onclick="populateReports('AP')" class="second-level">Action Plan</a></li>
      <li class="sub-divider"></li>
      <li><a href="#" onclick="populateReports('MISC')" class="second-level">Misc Report</a></li>
      <li class="sub-divider"></li>
    </ul>
  </li>	 -->
  	
  <li class="divider"></li>
	
	<li class="dropdown-submenu"> 	
    <a href="appSettings.jsp" class="first-level content_conf"><span class="content_conf_text"><spring:message code="menu.app.settings" /> </span></a>  
  	</li>
  	<li class="divider"></li>
  
  <li class="dropdown-submenu"> 	
    <a href="#" class="first-level content_conf"><span class="content_conf_text"><spring:message code="menu.cont.configs" /></span></a>  
   <ul class="dropdown-menu">
      <li><a href="reflectionQtn.jsp" class= "second-level"><spring:message code="menu.refl.questin" /></a></li>
      <li class="sub-divider"></li>
      <%-- <li class="dropdown-submenu">
        <a href="#" class= "second-level"><spring:message code="menu.user.role" /></a>
        <ul class="dropdown-menu">
        	<li><a href="#">3rd level</a></li>
        	<li><a href="#">3rd level</a></li>
        </ul>
      </li> --%>
      <li><a href="userRole.jsp" class= "second-level"><spring:message code="menu.user.role" /></a></li>
      <li class="sub-divider"></li>     
      <li><a href="mappingPage.jsp" class= "second-level"><spring:message code="menu.sub.create.mapping.values"/></a></li>
      <li class="sub-divider"></li>
      <li><a href="region.jsp" class= "second-level"><spring:message code="menu.region" /></a></li>
       <li class="sub-divider"></li>
      <li><a href="instruction.jsp" class= "second-level">Add / Edit Instructions</a></li>
    </ul>
  </li>
    
    <li class="divider"></li>
   <li class="dropdown-submenu"> 	
    <a href="myAccount.jsp" class="first-level content_conf"><span class="content_conf_text"><spring:message code="menu.my.accounts" /></span></a>  
  	</li>
  	

  <li class="divider"></li>
  <li class="dropdown-submenu">
    <a tabindex="-1" href="#" class="first-level"><spring:message code="menu.my.reports" /></a>
    <ul class="dropdown-menu">
      <li><a href="#" onclick="populateReports('BS')" class= "second-level"><spring:message code="menu.reports.bs" /></a></li>
      <li class="sub-divider"></li>     
      <li><a href="#" onclick="populateReports('AS')" class= "second-level"><spring:message code="menu.reports.as" /></a></li>
      <li class="sub-divider"></li>
      <li><a href="#" onclick="populateReports('BSTOAS')" class= "second-level"><spring:message code="menu.reports.bs.to.as" /></a></li>
      <li class="sub-divider"></li>
      <li><a href="#" onclick="populateReports('AP')" class= "second-level"><spring:message code="menu.reports.ap" /></a></li>
      <li class="sub-divider"></li>
      <li><a href="#" onclick="populateReports('MISC')" class= "second-level"><spring:message code="menu.reports.misc" /></a></li>
      <li class="sub-divider"></li>
      <li><a href="reportsAll.jsp" class= "second-level"><spring:message code="menu.reports.merged" /></a></li>
      <li class="sub-divider"></li>
    </ul>
  </li>
  
    <li class="divider"></li>
    <li class="dropdown-submenu"><a href="#" class="first-level content_conf"></a></li>
  <div class="clearfix"></div> 
</ul>              
	
<script src="../js/bootstrap.min.js"></script>
<script src="../js/menuHandler.js"></script>

</body>
</html>