<!DOCTYPE html>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<html>
<head>
    <title><spring:message code="label.landing.page.title"/></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link rel="shortcut icon" href="/user/img/crafting_ico.ico" />
    <!-- Bootstrap -->
    <link href="../css/bootstrap.css" rel="stylesheet">
    <link href="../css/style.css" rel="stylesheet">
    <link href="https://s3-us-west-2.amazonaws.com/jobcrafting/css/commonStyle.css" rel="stylesheet">
    <link rel="stylesheet" href="../css/alertify.core.css" />
	<link rel="stylesheet" href="../css/alertify.default.css" id="toggleCSS" />
	<script src="../js/alertify.min.js"></script>
    <script type="text/javascript" src="../lib/jquery.js"></script>
    <script type="text/javascript" src="../lib/spine.js"></script>
	<script type="text/javascript" src="../lib/ajax.js"></script>
	
    <!-- HTML5 Shim and Respond.js IE8 support of HTML5 elements and media queries -->
    <!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
      <script src="https://oss.maxcdn.com/libs/respond.js/1.3.0/respond.min.js"></script>
    <![endif]-->

    </head>
<body>
    <div class="main_warp">
    	<!-- Header area start -->
        	<jsp:include page="loginHeader.jsp"/>
     	<!-- Header area end -->
       

        <!-- Form area start -->
        
        <div class="container-fluid">  
           <div class="container" >
            <div class="exp-container" >
            	<div id="expLogoDiv" class="exp-logo-div">
            		<div align="center"> <img src="../img/exp-logo.png" class="warningClass"> </div>
            	</div>
            	<p>&nbsp;</p>
            	
            	
            	<div id="expTextDiv" class="exp-text-div">
	            	 <p>Your subscription has expired on <span id="expDate" style="font-weight: bold;"></span></p>
	            	 <p>Please <a href="#" onclick="submitForm()">renew</a> your subscription to continue.</p>
            	
            	<form name="ajaxform" id="ajaxform" method="post">
        			<input type="hidden" id="custType" name="attributes[ind_customer_type_0]" value="Individual" />
  				    <input type="hidden" id="subsType" name="attributes[ind_subs_type_0]" value="Renew"/>
        			<input type="hidden" id="custEmail" name="attributes[ind_customer_email_0]" />
        			<input type="hidden" id="prdVarient" name="id" />
            		<input type="hidden" name="quantity" value="1" />
            		<!-- <input type="hidden" name="return_to" value="back" /> -->
            		<!-- <input type="submit" value="GOOOO"> -->
        		</form>
            	
            	</div>
            	
            	
            	
            </div>
           </div>
        </div>
         <!-- Form area end -->
        <!-- Footer area start -->
        <div class="row-fluid footer footer-wrapper footer_landing">
          <div class="container">
            <div class="copyright"><spring:message code="label.footer"/></div>
          </div>
        </div>
        <!-- Footer area end -->
    </div>

<!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->

<script src="https://code.jquery.com/jquery.js"></script>    
<script src="../js/bootstrap.min.js"></script>


<script src="../js/jquery.bxslider.min.js" type="text/javascript"></script>
<link href="../css/jquery.bxslider.css" rel="stylesheet" type="text/css" />
<!-- <link href="../css/bootstrap.min.css" rel="stylesheet" type="text/css" /> -->
<script src="../js/common.js"></script>
<script src="../js/accountRenew.js"></script>


<script type="text/javascript">

		/**
		 * Function add to disable browser back button
		 * while the page is loaded
		 * @param null
		 */
		window.location.hash="";
		window.location.hash="Again-No-back-button";//again because google chrome don't insert first hash into history
		window.onhashchange=function(){window.location.hash="";};

/* 
         $(document).ready(function () {           
             $('.bxslider').bxSlider({
                 mode: 'vertical',
                 slideMargin: 3,
                 auto:true
             });             
         }); */
    </script>
           


 
</body>

</html>