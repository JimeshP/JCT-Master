<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<head>
    <title><spring:message code="menu.sub.create.user" /></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <!-- Bootstrap -->
    
    <link href="../css/bootstrap.css" rel="stylesheet">
    <link href="../css/style.css" rel="stylesheet">
    <link rel="stylesheet" href="../css/alertify.core.css" />
	<link rel="stylesheet" href="../css/alertify.default.css" id="toggleCSS" />
	<meta name="viewport" content="width=device-width">
	<link rel="shortcut icon" href="/admin/img/crafting_ico.ico" />
    <!-- HTML5 Shim and Respond.js IE8 support of HTML5 elements and media queries -->
    <!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
      <script src="https://oss.maxcdn.com/libs/respond.js/1.3.0/respond.min.js"></script>
    <![endif]-->
	<script type="text/javascript" src="../lib/jquery.js"></script>
	<script type="text/javascript" src="../lib/spine.js"></script>
	<script type="text/javascript" src="../lib/ajax.js"></script>
	<!-- <script src="http://code.jquery.com/jquery-latest.min.js"></script> -->
	<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.8.3/jquery.min.js"></script>
	<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jqueryui/1.8.23/jquery-ui.min.js"></script>
	<script src="../lib/jquery-1.7.2.min.js"></script>
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
	<script src="../js/alertify.min.js"></script>
	<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css"> 
</head>
	<body>
	    <div class="main_warp">
	    
	    <!-- Header area start -->
	    <jsp:include page="header.jsp"/>
	    <!-- Header area end --> 	       
	        <div class="row">  
	            <div class="main-cont-wrapper" >
	            <div class="col-sm-2 pd-0">
	            	<!-- Menu area start -->
	    			<jsp:include page="sideMenuNew.jsp"/>
	   				 <!-- Menu area end -->   
	            </div>
	            
	            <div class="col-sm-10 col">
	            <div class= "heading_label">
	            	<ol class="breadcrumbAdmin">
           				<li class="first_nav"><spring:message code="reports.breadcrumb.manage.user.home"/> &nbsp;|</li>
           				<li class="first_nav"><spring:message code="menu.app.manage.user"/> &nbsp;|</li>
           				<li class="second_nav" id='updateId'><spring:message code="label.addRenew"/></li>
             		</ol>
	            </div>
	            <div class= "sub_contain"><spring:message code="label.addRenew"/></div>
	            <div class= "main_contain">
	            <div class= "inner_container" id=""><spring:message code="label.addRenew"/></div>	  
	            <form class="form-horizontal main_div" name="addNewUser" autocomplete="on" novalidate enctype="multipart/form-data" data-remote="true"> 
	            <div class="single_form_item"><br/></div>
	
	            
	   <!-- 1st Row Start -->
	            <div class="single_form_item addRenewDiv">
                     <div class="col-sm-3">
                     <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="label.select.user.type"/></label>
					 </div>
                     <div class="col-sm-3">
                   	
                     	<select class="form-control-general font_ipad" id="selectUserType" onchange="toggleRadio()">
	                   		<option class="form-control-general" value="0"><spring:message code="label.user.type.select.drpDwn"/></option>  
							<option class="form-control-general" value="1"><spring:message code="label.user.type.individual"/></option>      
							<option class="form-control-general" value="2"><spring:message code="label.user.type.facilitator"/></option>       
	                   	</select>
                     </div> 
					 					
					 <!-- Add New / Renew existing radio -->
					<div id="radioNewRenew" style="  margin-left: 20%;">
						<div class="col-sm-3">
							<label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="menu.sub.select.option"/></label>
						</div>
						<div class="col-sm-3 sub_text_label">
							<div class="col-sm-5  radio_manual">
								<input type="radio" name="actionRadio" id="newRadio" value="N" checked = "checked"><spring:message code="label.addRenew.addNew"/>
							</div>
							<div class="col-sm-5 radio_csv" id="renewRadioDiv">
								<input type="radio" name="actionRadio" id="renewRadio" value="R" ><spring:message code="label.addRenew.renewExisting"/>
							</div>
							<div class="col-sm-5 radio_csv" id="subRadioDiv" style="display: none;">
								<input type="radio" name="actionRadio" id="subRadio" value="R" ><spring:message code="label.addRenew.subExisting"/>
							</div>
						</div>
					</div> 
                   </div>      
		   			
		   		<div class="single_form_item">
                   <p>
					<input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Proceed" id="proceed" onclick="" />
				   </p>
                </div>	
		   
		         
                   <div class="single_form_item"><br/></div>        
		 <!-- </div> -->
		   

                  
	            </form>
	            

	                       </div>
	            </div>
	            
	                     
	            </div>	            
	            
				</div>
	    	</div>
	         <!-- Form area end -->
	    <!-- Footer area start -->
	    <jsp:include page="footer.jsp"/>
	    <!-- Footer area end --> 	          
		
		
<!-- <script src="https://code.jquery.com/jquery.js"></script> -->    
<!-- Include all compiled plugins (below), or include individual files as needed -->
<script src="../js/bootstrap.min.js"></script>
<script src="../js/bootstrap-filestyle.js"></script>
<script src="../js/common.js"></script>
<script src="../js/addRenewUsers.js"></script>
<script src="../js/jquery.tablesorter.js"></script>
<script type="text/javascript">
$(":file").filestyle();
</script>
<script type="text/javascript" src="../js/tooltip.js"></script>
	<script type="text/javascript" src="../js/popover.js"></script>
	<script>
	$(document).ready(function() {
	$("input").focus(function() {
	//alert("focus sss");
	  $(this).popover('show');
	  });
	  
	  $("input").click(function() {
	  //alert("clicked");
	  $(this).popover('toggle');
	  });

	 $("input").blur(function() {
	 //alert("sdfdhfjld");
	  $(this).popover('hide');
	  });
	 });
	</script>
	
	
	<!-- ALL ARE CORRECT -->
    <div class="modal fade" id="myModal" tabindex="-1" data-backdrop="static">
    <div class="modal-dialog2" >
    <div class="modal-content2">
    <div class="new_modal_header modal-header">
        &nbsp;      
    </div>
    	<div align="center">   
      		<span id="message"></span>
      	</div>
      	<div align="center">   
      		<nav class="alertify-buttons">
      			<button class="alertify-button alertify-button-ok" id="alertify-oks" onclick="saveUserAfterValidation()">OK</button>
      			<button class="alertify-button alertify-button-cancel" id="alertify-cancels" onclick="closeMyModel()">Cancel</button>
      		</nav>
      	</div>
      	<div class="new_modal_header modal-header">
        	&nbsp;      
    	</div>
      	<div align="center">   
      		<span id="tableData"></span>
      	</div>
    </div>
  	</div>
	</div>
	
	<!-- ERROR -->
	<div class="modal fade" id="myModal1" tabindex="-1" data-backdrop="static">
    <div class="modal-dialog2" >
    <div class="modal-content2">
    <div class="new_modal_header modal-header">
        &nbsp;      
    </div>
    	<div align="center" class="error_msg_text">   
      		<span id="message1"></span>
      	</div>
      	<div align="center" class="error_msg_text">   
      		<span id="tableData1"></span>
      	</div>
      	<div align="center">   
      		<nav class="alertify-buttons">
      			<button class="alertify-button alertify-button-ok" id="alertify-oks1" onclick="closeMyModel1()">OK</button>
      		</nav>
      	</div>
      	<br>
    </div>
  	</div>
	</div>
	
</body>
</html>