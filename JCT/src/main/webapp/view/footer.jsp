<!DOCTYPE html>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<html>

<head>
    <title>..: Footer :..</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <!-- Bootstrap -->
    <link href="../css/bootstrap.css" rel="stylesheet">
    <link href="../css/style.css" rel="stylesheet">
    <link href="../css/less-768.css" rel="stylesheet">
    <link href="../css/less-1280.css" rel="stylesheet">        
</head>
<body>
	<div class="row-fluid footer footer-wrapper">
	          <div class="container">
            <div class="copyright">
            	<spring:message code="label.footer"/><br>
            	<spring:message code="label.footer.permission"/>
            
            	<span style="float: right; cursor: pointer; margin-right: -2%" onclick="fetchTnstructions()"><spring:message code="label.jct.tool.version"/></span></div>
          </div>
	</div>
	<!-- Terms & Conditions modal box -->
	<div class="modal fade" id="tnCModal" tabindex="-1" data-backdrop="static">
	  <div class="modal-dialog" >
	    <div class="modal-content">
	      <div class="tcModal_header modal-header">
	        <button type="button" class="custom_close_btn" data-dismiss="modal"></button>      
	        <h5 class="modal-title tcModal_title" id="myModalLabel">Terms and Conditions</h5>
	      </div>
	      <br>
		   <div class="instruction_panel tcText_style" id="terms_condtn_div_id"></div>
	      <div class="modal-footer" style="padding: 5px 20px 8px;">
	      		<input type="submit" class="btn btn-primary btn-sm cstm-login-btn reset_pass tcClose-btn" value="Close" data-dismiss="modal">
	      </div>
	    </div>
	  </div>
	  
	</div>
	
	
	
	
	<a data-icon="A" id="top_link" href="#header-wrap" style="display: inline;">
	<span class="glyphicon glyphicon-open"></span></a>
	    <script>            
	    $(window).scroll(function() {
	        if ($(this).scrollTop()) {
	            $('#top_link').fadeIn();
	        } else {
	            $('#top_link').fadeOut();
	        }
	    });

	    $("#top_link").click(function () {
	       //1 second of animation time
	       //html works for FFX but not Chrome
	       //body works for Chrome but not FFX
	       //This strange selector seems to work universally
	       $("html, body").animate({scrollTop: 0}, 1000);
	    });
		</script>
</body>
</html>