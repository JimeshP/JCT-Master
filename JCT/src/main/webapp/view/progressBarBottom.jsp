<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<!-- <head>
    <title>..: Header :..</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    Bootstrap
    <link href="../css/bootstrap.css" rel="stylesheet">
    <link href="../css/style.css" rel="stylesheet">  
</head> -->

<body> 
	
	<div class="progressCol progressbar_bottom">
		<img height="159" width="61" src="../img/progressNew.png" class="progressFlip"> 
		<div class="progressBox">
			<img height="28" width="56" src="../img/menu_hide.png" class="progress-hide">
			<title class= "heading-progress"><spring:message code="label.progress"/></title>
			<ul class="progressList">
				<li class="green"><spring:message code="label.beforeSketch"/></li>
				<ul class="subprogressList">
					<li class="bs_step1 activelist"><span>1. <spring:message code="label.defineYourTask"/></span></li>
					<li class="bs_step2 "><span>2. <spring:message code="label.reflectionQtn"/></span></li>
				</ul>
			<ul class="progressList">
				<li class="green"><spring:message code="label.afterSketch"/></li>
				<ul class="subprogressList">
					<li class="as_step1 activelist"><span>1. <spring:message code="label.mappingYourself"/></span></li>
					<li class="as_step2"><span>2. <spring:message code="label.puttingItTogether"/></span></li>
					<li class="as_step3"><span>3. <spring:message code="label.actionPlan"/></span></li>
				</ul>
			</ul>
		</div>
      </div>    
</body>
</html>