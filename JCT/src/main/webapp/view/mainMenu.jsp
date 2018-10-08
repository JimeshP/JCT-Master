<!DOCTYPE html>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<html>
<head>
    <title></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <!-- Bootstrap -->
    <link href="../css/bootstrap.css" rel="stylesheet">
    <link href="../css/style.css" rel="stylesheet">  
    <script type="text/javascript" src="../js/jquery.fancybox.pack.js"></script>
    <script type="text/javascript" src="../js/iscroll.js"></script>
	<link rel="stylesheet" type="text/css" href="../css/jquery.fancybox.css" media="screen" />
	<script type="text/javascript" src="../js/common.js"></script>
</head>

<body>

	<form name="ajaxform" id="ajaxform" method="post">
		<input type="hidden" id="custType" name="attributes[ind_customer_type_0]" value="Individual" />
  		<input type="hidden" id="subsType" name="attributes[ind_subs_type_0]" value="Renew"/>
    	<input type="hidden" id="custEmail" name="attributes[ind_customer_email_0]" />
        <input type="hidden" id="prdVarient" name="id" />
        <input type="hidden" name="quantity" value="1" />
   </form>
        		
   <nav class="navbar navbar-default">
	<div class="container-fluid nav-header">
    	<div class="header_wrap_area container" >
            <div class="navbar-header" > 
            <button type="button" class="navbar-toggle" data-toggle="collapse" data-target="#bs-example-navbar-collapse-1">
        <span class="sr-only">Toggle navigation</span>
        <span class="icon-bar"></span>
        <span class="icon-bar"></span>
        <span class="icon-bar"></span>
      </button>
      <a class="navbar-brand" href="#" data-toggle="collapse" data-target="#bs-example-navbar-collapse-1">Show Menu</a>
            </div> 
            <div class="row navbar-collapse collapse" id="bs-example-navbar-collapse-1">
            <div class="col-xs-12 col-md-8">
            <ul class="nav navbar-nav main_nav">
           <%--  <li><a id="fancybox-manual-my-account" href="#" onclick="openMyAccountDetails()"><spring:message code="label.menu.myaccount"/></a></li> --%>
            <li class="dropdown"><a href="#" class="dropdown-toggle" data-toggle="dropdown"><spring:message code="label.menu.search"/><b class="caret"></b></a>
             <ul class="dropdown-menu">
              <li><a id="fancybox-manual-b" href="#" onclick="searchBeforeSketch()"><spring:message code="label.menu.beforesketch"/></a></li>
              <li><a id="fancybox-manual-a" href="#" onclick="searchAfterSketch()"><spring:message code="label.menu.afterdiagram"/></a></li>
              </ul>
            </li>
            </ul>
            </div>
          
	            <div class="notificationNew">
				<div id="notificationIconDiv" class="notificationIconDiv">
						<img id="icon" class="icon" onmouseenter="showTooltip()" onmouseleave="hideTooltip()"/>
				</div>
				<div id="notificationTextDiv" class="notificationTextDiv">
					<p></p>
				</div>
				<div class="tooltip1">
					<div class="tooltipTextHolder">
						<p id="tooltipText" class="tooltipText"></p>
					</div>					
                </div>
		    </div>
           </div>
            
           </div> 
         </div>
	<!-- </div> -->
	</nav>  
	<script src="../js/sketchSearch.js"></script>
	<script src="../js/accountRenew.js"></script>
	<!-- internal defect found & fix job label & field closer -->
	<div id='sketchSearchDiv' class="diagramSearch">
			<div id="searchingPanel" class="popupSearchHeader">
				    <%--  <div class="col-md-2 job_label">
                     	<label for="inputFN" class="col-md-12 control-label same-line"><spring:message code="label.occupation.search"/></label>
                     </div> --%>
                     <div class="col-md-7 job_label_dropdown">
                     	<input type="text" class="form-control-general" id="occupationBS" value="" placeholder="Enter words describing occupation" />
                     </div> 
                     <div class="col-md-3 search_diagram"><input type="button" value="Search Occupations" onclick="searchOccupation('BS')" id="bsSearchButtonId" class="btn btn-primary btn-sm"/></div>
                     <div class="col-md-1 search_diagram"><input type="button" value="Reset" onclick="searchBeforeSketchAll()" id="bsSearchButtonId" class="btn btn-primary btn-sm" style="float: right;"/></div> 
                	<div class="clearfix"></div>
			 </div>
		<div id="headerDiv" class="popupHeader">
		<table id='bsTable' style="display: none;">
				<tr>
					<td width="15%">
						<input type='button' id="popupBackButton" class='popupBackButton' onClick='goBack()' value='Back'>
					</td>
					<td width="70%" class= "job_level_text">
						<span id='bsHeader'></span>
					</td>
					<td width="15%">
						<input type='button' class='popupBackButton' id="asAvailability" value='<spring:message code="label.view.after.sketch"/>'>
						<input type='button' class='popupBackButton' id="bsAvailability1" value='<spring:message code="label.view.before.sketch"/>' style="display: none;">
					</td>
				</tr>
			</table>
			<table id='bsTable2'>
				<tr>
					<td>
						&nbsp;&nbsp;&nbsp;
					</td>
					<td class= "occupation_text" width='100%'>
						<span id='bsHeaderInitial'></span>
					</td>
					<td>
						&nbsp;&nbsp;&nbsp;
					</td>
				</tr>
			</table>
		</div>
		<div id="imageDiv" style="height: 370px; overflow: scroll;">
		</div>
		<div id="imageEnlargeDiv" style="display: none;">
			<div id="mainPic" style="height: 440px; overflow: scroll;">
			</div>
		</div>
		<div id="contoller" style="position:absolute; bottom:0; text-align:center; width:100%"></div>
		<div id="contollerOccupation" style="display:none; position:absolute; bottom:1; text-align:center; width:100%">
			<table width="20%" align="center">
				<tr>
					<td>
						<input type="button" class='popupBackButton' onclick="processSelectedOccupation('BS')" value="Select">
					</td>
				</tr>
			</table>
			
		</div>
	</div> 
	
	<!-- After Sketch -->
	<div id='ASsketchSearchDiv' class="diagramSearch">
			<div id="ASsearchingPanel" class="popupSearchHeader">
				    <%--  <div class="col-md-2 job_label">
                     	<label for="inputFN" class="col-md-12 control-label same-line"><spring:message code="label.occupation.search"/></label>
                     </div> --%>
                     <div class="col-md-7 job_label_dropdown">
                     	<input type="text" class="form-control-general" id="occupationAS" value="" placeholder="Enter words describing occupation"/>
                     </div> 
                     <div class="col-md-3 search_diagram"><input type="button" value="Search Occupations" onclick="searchOccupation('AS')" id="asSearchButtonId" class="btn btn-primary btn-sm"/></div>
                     <div class="col-md-1 search_diagram"><input type="button" value="Reset" onclick="searchAfterSketchAll()" id="asSearchButtonId" class="btn btn-primary btn-sm" style="float: right;"/></div> 
                	<div class="clearfix"></div>
			 </div>
		<div id="ASheaderDiv" class="popupHeader">
		<table id='asTable' style="display: none;">
				<tr>
					<td width="15%">
						<input type='button' id="ASpopupBackButton" class='popupBackButton' onClick='goBackAS()' value='Back'>
					</td>
					<td width="70%" class= "job_level_text">
						<span id='asHeader'></span>
					</td>
					<td width="15%">
						<input type='button' class='popupBackButton' id="asAvailability1" value='<spring:message code="label.view.after.sketch"/>' style="display: none;">
						<input type='button' class='popupBackButton' id="bsAvailability" value='<spring:message code="label.view.before.sketch"/>'>
					</td>
				</tr>
			</table>
			<table id='asTable2'>
				<tr>
					<td>
						&nbsp;&nbsp;&nbsp;
					</td>
					<td class= "occupation_text" width='100%'>
						<span id='asHeaderInitial'></span>
					</td>
					<td>
						&nbsp;&nbsp;&nbsp;
					</td>
				</tr>
			</table>
		</div>
		<div id="ASimageDiv" style="height: 370px; overflow: scroll;">
		</div>
		<div id="ASimageEnlargeDiv" style="display: none;">
			<div id="ASmainPic" style="height: 440px; overflow: scroll;">
			</div>
		</div>
		<div id="AScontoller" style="position:absolute; bottom:0; text-align:center; width:100%"></div>
		<div id="AScontollerOccupation" style="display:none; position:absolute; bottom:1; text-align:center; width:100%">
			<table width="20%" align="center">
				<tr>
					<td>
						<input type="button" class='popupBackButton' onclick="processSelectedOccupation('AS')" value="Select">
					</td>
				</tr>
			</table>
			
		</div>
	</div> 
	
	<!-- internal defect found & fix job label & field closer --> 
	<!-- MY ACCOUNT DIV -->
	<div id='myAccountDiv' class="scroll">
		<div class=" clock_area " style="text-align: center; margin: 0 auto; width: 60%;">
              <div class="current_time_hrs_txt"><spring:message code="label.currentTimeHours"/>&nbsp;:</div>                         
               <div class="current_time_hrs">
               <form name="clock">
                <input type="hidden" id="stwa" size="18" name="stwa" value="01 : 00 : 00" style="text-align:center" class="form-control_current_time input-sm current_time_field"  disabled />
                <input type="text" id="stwa2" size="18" name="stwa2" value="01 : 00 : 00" style="text-align:center" class="form-control_current_time input-sm current_time_field"  />
                <!-- <span id="stwa2" class="total_time_field"></span>  -->
                </form>
                </div>               
                <%-- <div class="current_hrs"><spring:message code="label.hrs"/></div> --%>                 
                <div class="total_time_text"><spring:message code="label.totalHours"/>
              		
              <span id="timeId" class="total_time_field"></span> 
               <%-- <spring:message code="label.hrs"/>  --%>
               </div>
               <div class="clearfix"></div>
         </div>
         <br />
         <div class="clearfix"></div>
		<div id="myAccountheaderDiv" class="popupHeader"></div>
		<div id="myAcValDiv" class="padding-table"></div>
		<!-- <a href="#" id="passwordResetLink" class="padding-table" onclick="toogleResetDiv()">Reset Password</a> -->
		<div align="center">
			<input type="button" class="btn btn-primary btn-sm" value='<spring:message code="label.resetPassword"/>' id="passwordResetLink" onClick="toogleResetDiv()" />
		</div>
		
		<div id="passwordResetArea" class="padding-table" style="display: none;">
			<div class="single_form_item">
            	<div class="col-md-4 ">
                    <label for="inputFN" class="col-md-12 control-label align-right"><spring:message code="label.existing.password"/>*</label>
                </div>
                <div class="col-md-6"><input type="password" class="form-control-general" maxlength="30"  value="" id="inputEPass" onkeypress="searchMyKeyEvent(event);" autofocus></div> 
                <div class="col-md-2">&nbsp;</div>  
                <div class="clearfix"></div>
            </div>
            
            <div class="single_form_item">
            	<div class="col-md-4 ">
                    <label for="inputFN" class="col-md-12 control-label align-right"><spring:message code="label.new.password"/>*</label>
                </div>
                <div class="col-md-6"><input data-container="body" data-toggle="popover" data-placement="right" data-content="New Password should contain atleast one lowercase letter, one uppercase letter, one numeric digit, and one special character." type="password" class="form-control-general" maxlength="30"  value="" id="inputNPass" onkeypress="searchMyKeyEvent(event);" autofocus></div> 
                <div class="col-md-2">&nbsp;</div>  
                <div class="clearfix"></div>
            </div>
            
            <div class="single_form_item">
            	<div class="col-md-4 ">
                    <label for="inputFN" class="col-md-12 control-label align-right"><spring:message code="label.confirm.password"/>*</label>
                </div>
                <div class="col-md-6"><input type="password" class="form-control-general" maxlength="30"  value="" id="inputCPass" onkeypress="searchMyKeyEvent(event);" autofocus></div> 
                <div class="col-md-2">&nbsp;</div>  
                <div class="clearfix"></div>
            </div>
            <div class="single_form_item">
                <p><input type="button" class="btn btn-primary btn-sm" value="Reset" id="rstMyPaswId" onClick="resetMyPassword()" /></p>
            </div>
        </div>
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
			
			var src = "Tool";
			if (sessionStorage.getItem("src")) {
				src = sessionStorage.getItem("src");
			}
			// Enable the open my account details feature only if the user has logged in from the tool
			if (src == "Tool") {
			$("#fancybox-manual-my-account").click(function() {
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
			}
			
		});
	</script>
</html>