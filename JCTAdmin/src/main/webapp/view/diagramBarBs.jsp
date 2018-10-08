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
	
	<div class="diagramCol">
		<img height="159" width="61" src="../img/diagram_flip.png" class="diagramFlip"> 
		<div class="diagramBox">
			<div class="diagram_head">
			<title class="heading-diagram"><spring:message code="label.beforeDiagram"/></title>
			<img height="27" width="56" src="../img/hide_diagram.png" class="diagram-hide">			
			</div>
             <div id="bsDiag"></div>
		</div>
      </div>    
</body>
</html>