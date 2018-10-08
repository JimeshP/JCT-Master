<!DOCTYPE html>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<html>
<head>
    <title></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <!-- Bootstrap -->
    <link href="../css/bootstrap.css" rel="stylesheet">
    <link href="../css/style.css" rel="stylesheet">  
    <script type="text/javascript" src="../js/jquery.fancybox.pack.js"></script>
    <script type="text/javascript" src="../js/iscroll.js"></script>
	<link rel="stylesheet" type="text/css" href="../css/jquery.fancybox.css" media="screen" />
</head>

<body> 
	<div class="row-fluid nav-header">
    	<div class="header_wrap_area container" >
             <!-- <div class="question_icon"><a href="#"><img src="../img/question_icon.png" alt="question_icon" /></a></div> -->

            <ul class="nav navbar-nav main_nav">
            <li><a href="#"><spring:message code="label.menu.myaccount"/></a></li>
            <li class="dropdown"><a href="#" class="dropdown-toggle" data-toggle="dropdown"><spring:message code="label.menu.search"/><b class="caret"></b></a>
             <ul class="dropdown-menu">
              <li><a id="fancybox-manual-b" href="#" onclick="searchBeforeSketch()"><spring:message code="label.menu.beforesketch"/></a></li>
              <li><a id="fancybox-manual-a" href="#" onclick="searchAfterSketch()"><spring:message code="label.menu.afterdiagram"/></a></li>
              </ul>
            </li>
            <li><a href="#"><spring:message code="label.menu.abt.jct"/></a></li>
            </ul>
         </div>
	</div>  
	<script src="../js/sketchSearch.js"></script> 
	<div id='sketchSearchDiv' class="scroll">
		<div id="headerDiv" class="popupHeader"></div>
		<div id="imageDiv">
		</div>
		<div id="paginator" class="paginator">
		<div id="contoller">
		</div> 
		</div>
		<div id="imageEnlargeDiv" style="display: none;">
			<div id="backButton">
			</div>
			<div id="mainPic">
			</div>
		</div>
		<div id="imageBS" style="text-align: center;"></div>
	</div> 
	
	<div id='ASsketchSearchDiv' class="scroll">
		<div id="ASheaderDiv" class="popupHeader"></div>
		<div id="ASimageDiv">
		</div>
		<div id="ASpaginator" class="paginator">
		<div id="AScontoller">
		</div> 
		</div>
		<div id="ASimageEnlargeDiv" style="display: none;">
			<div id="ASbackButton">
			</div>
			<div id="ASmainPic">
			</div>
		</div>
		<div id="imageAS" style="text-align: center;"></div>
	</div> 
</body>
<script type="text/javascript">
		$(document).ready(function() {
			$("#fancybox-manual-b").click(function() {
				$.fancybox.open({
					href : '#dd',
					//type : 'iframe',
					closeClick : false,
					openEffect : 'elastic',
					openSpeed  : 150,
					closeEffect : 'elastic',
					closeSpeed  : 150,
					padding : 5
				});
			});
			
			$("#fancybox-manual-a").click(function() {
				$.fancybox.open({
					href : '#dd',
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
</html>