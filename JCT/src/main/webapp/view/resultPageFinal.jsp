<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<head>
<title><spring:message code="label.result.page" /></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
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
<script type="text/javascript" src="../js/jquery.fancybox.pack.js"></script>
<script type="text/javascript" src="../js/iscroll.js"></script>
<link rel="stylesheet" type="text/css" href="../css/jquery.fancybox.css"
	media="screen" />
<script src="http://code.jquery.com/jquery-latest.min.js"></script>
<script>
if (typeof jQuery == 'undefined') {
    document.write(unescape("%3Cscript src='../js/latest_10.2_jquery.js' type='text/javascript'%3E%3C/script%3E"));
}
</script>
<!-- <script src="http://code.jquery.com/jquery-latest.min.js"></script> -->

<script src="http://code.jquery.com/ui/1.10.3/jquery-ui.js"
	type="text/javascript"></script>
<script type="text/javascript">
//<![CDATA[
(window.jQuery && window.jQuery.ui && window.jQuery.ui.version === '1.10.3')||document.write('<script type="text/javascript" src="../js/jquery_ui.js"><\/script>');//]]>
</script>



<!-- <script src="http://code.jquery.com/ui/1.10.3/jquery-ui.js"></script> -->
<script src="../js/alertify.min.js"></script>
<link rel="stylesheet"
	href="http://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css">
</head>
<body>
	<div class="main_warp">

		<!-- Header area start -->
		<jsp:include page="header.jsp" />
		<!-- Header area end -->

		<!-- Menu area start -->
		<jsp:include page="mainMenu.jsp" />
		<!-- Menu area end -->

		<div class="container" id="header-wrap">
			<div class="row">
				<div class="form_area">
					<div class="form_area_top col-md-12 prev_nxt_sec">
						<div class="col-md-2 col-xs-3 Prev_btn_area ">
							<a href="#" class="btn btn-info prev-btn" onclick="goPrevious()"><span
								class="glyphicon glyphicon-chevron-left"></span> <spring:message
									code="label.previousStep" /></a>
							<div class="clearfix"></div>
						</div>
						<div class="col-md-8 col-xs-12 text-center heading-title">
							<spring:message code="label.result" />
						</div>
						<%--  <div class="col-md-2 ">
                  <a href="#" onclick="goToNextPage()" class="btn btn-info prev-btn pull-right"><spring:message code="label.nextStep"/> <span class="glyphicon glyphicon-chevron-right"></span></a>
                  </div> --%>
						<div class="clearfix"></div>
					</div>

					<div class="form_area_middle">
						<div class="form_area_wraper_justified_page">
							<form class="form-horizontal" autocomplete="on" novalidate>


								<div class="cont-middle">
									<div class="result-wrapper">
										<h4 class="notification">
											<spring:message code="label.youCan" />
											<a href="#" class="click_here" id="generateURL"
												onclick="downloadFile()" target="_blank"> <spring:message
													code="label.clickHere" />
											</a>
											<spring:message code="label.download" />
											<br /> <br />
										</h4>
									</div>

									<div id="infoMsgDispFinal" style="display: block;">
										<h4 class="notification">
											<spring:message code="label.final.notification" />
										</h4>
									</div>
									<div class="clearfix"></div>

									<form role="form" class="cstm-frm-wrapper">
										<div class="radio">
											<label> <input type="radio" name="optionsRadios"
												id="keepPreviousDiagId" value="option1" checked> <spring:message
													code="label.final.radioOnePartOne" />&nbsp;<a href="#"
												id='fancybox-manual-b-final' onclick="searchDiags()"
												class="click_here"><spring:message
														code="label.final.view" /></a>
											</label>
										</div>
										<div class="radio">
											<label> <input type="radio" name="optionsRadios"
												id="addWithoutEditsId" value="option2"> <spring:message
													code="label.final.radioTwo" />
											</label>
										</div>
										<div class="radio">
											<label> <input type="radio" name="optionsRadios"
												id="addAfterEditId" value="option2"> <spring:message
													code="label.final.radioThree" />
											</label>
										</div>
									</form>
									<div class="clearfix"></div>
									<br>
									<div class="start-btn">
										<a class="multi-line-button green" href="#" id='finishId'
											style="width: 14em" onClick="submitDecision()"> <span
											class="title"><spring:message code="label.finish.flow" /></span>
										</a>
									</div>



									<div class="result-wrapper">
										<h4 class="notification">
											<spring:message code="label.previous.pdf.link1" />
											<a href="#" onclick="showOldPDF()" class="finalPage_desc">click here</a>.
											<br /><br />
										</h4>
									</div>



									<%-- <div class="dwn-pdf">
                  <spring:message code="label.youCan"/>
                  <a href="#" class="click_here" id="generateURL" onclick="downloadFile()" target="_blank">
                  		<spring:message code="label.clickHere"/>
                  </a>
                  <spring:message code="label.download"/>                   
                  </div> --%>

								</div>
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
	<jsp:include page="footer.jsp" />

	<!-- Footer area end -->
	<%-- <div class="loader_bg" style="display:none"></div>
        <div class="loader" style="display:none"><img src="../img/Processing.gif" alt="loading"></div>
        <div class="row-fluid footer">
          <h5 class="align_center footer_heading"><spring:message code="label.footer"/></h5>
        </div> --%>
	<!-- Footer area end -->
	<div class="loader_bg" style="display: none"></div>
	<div class="loader" style="display: none">
		<img src="../img/Processing.gif" alt="loading">
	</div>
	<!-- Footer area end -->

	<div id='sketchSearchDivFinal' class="scroll">
		<div id="headerDivFinal" class="popupHeader">
			<table width='100%'>
				<tr>
					<td class="header_align" width='60%'><spring:message
							code="lbl.bsd" /></td>
					<td align='right'><input type="button" class="popupBSButton"
						onclick="fetchASDiag()"
						value="<spring:message code='label.view.after.sketch'/>" /></td>
			</table>
		</div>
		<div id="headerDivFinal2" class="popupHeader" style="display: none;">
			<table width='100%'>
				<tr>
					<td class="header_align" width='60%'><spring:message
							code="lbl.asd" /></td>
					<td align='right'><input type="button" class="popupBSButton"
						onclick="populateBSDiagram()"
						value="<spring:message code='label.view.before.sketch'/>" /></td>
			</table>
		</div>
		<div id="imageEnlargeDivFinal">
			<div id="backButtonFinal"></div>
			<div id="mainPicFinal" style="align: center;"></div>
			<%-- <div align='center' id='bsLinkDiv' style='display: none;'>
				<a href="#" id='fancybox-manual-b-final' onclick="fetchASDiag()" class="click_here"><spring:message code="label.view.after.sketch"/></a>
			</div>
			<div align='center' id='asLinkDiv' style='display: none;'>
				<a href="#" id='fancybox-manual-b-final' onclick="populateBSDiagram()" class="click_here"><spring:message code="label.view.before.sketch"/></a>
			</div> --%>
			<!--<a id="AvailabilityFinal" style="display: none;  text-align:center; cursor: pointer;">
				<spring:message code="label.view.before.sketch"/>
			</a>
			<div id="backButton">
				<input type="button" class="popupBackButton" onclick="goBack()" value="Back">
			</div>-->
		</div>
		<div id="imageFinal" style="text-align: center;"></div>
	</div>
	<div id="rectificationId" class="scroll"></div>
	</div>
	<script src="../js/bootstrap.min.js"></script>
	<script src="../js/session_storage.js"></script>
	<script src="../js/resultPageFinal.js"></script>
	<script src="../js/stopWatch.js"></script>
	<script src="../js/common.js"></script>

	<!-- <div id='occupationDiv' class="scroll">
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
		<input type="button" class="btn btn-primary btn-sm" value="Go" onclick="closeDiv()" />
	</div>
</div> -->
</body>
<script type="text/javascript">
		$(document).ready(function() {
			$("#fancybox-manual-b-final").click(function() {
				$.fancybox.open({
					href : '#dd',
					closeClick : false,
					openEffect : 'elastic',
					openSpeed  : 150,
					closeEffect : 'elastic',
					closeSpeed  : 150,
					padding : 5
				});
			});
			
			$("#nonExistentId").click(function() {
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