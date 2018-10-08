<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<head>
    <title><spring:message code="menu.sub.create.instruction" /></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <!-- Bootstrap -->
    <link href="../css/bootstrap.css" rel="stylesheet">
    <link href="../css/style.css" rel="stylesheet">
    <link rel="stylesheet" href="../css/alertify.core.css" />
	<link rel="stylesheet" href="../css/alertify.default.css" id="toggleCSS" />
	<meta name="viewport" content="width=device-width">
   <link rel="shortcut icon" href="/admin/img/crafting_ico.ico" />
	<script type="text/javascript" src="../lib/jquery.js"></script>
	<script type="text/javascript" src="../lib/spine.js"></script>
	<script type="text/javascript" src="../lib/ajax.js"></script>
	<script src="http://code.jquery.com/jquery-latest.min.js"></script>
	<script>
	/* if (typeof jQuery == 'undefined') {
    	document.write(unescape("%3Cscript src='../js/latest_10.2_jquery.js' type='text/javascript'%3E%3C/script%3E"));
    } */
	</script>	
	<!-- <script src="http://code.jquery.com/ui/1.10.3/jquery-ui.js" type="text/javascript"></script> -->
	<!-- <script type="text/javascript">
	//<![CDATA[
        (window.jQuery && window.jQuery.ui && window.jQuery.ui.version === '1.10.3')||document.write('<script type="text/javascript" src="../js/jquery_ui.js"><\/script>');//]]>
    </script> -->	
    <script src="//cdn.ckeditor.com/4.5.1/standard/ckeditor.js"></script>
	<script src="../js/alertify.min.js"></script>
	<script type="text/javascript" src="../js/jquery-1.3.2.js"></script>
    <script type="text/javascript" src="../js/jquery-ui-1.7.2.custom.min.js"></script>
    <link rel="Stylesheet" type="text/css" href="../css/jqueryui/ui-lightness/jquery-ui-1.7.2.custom.css" />

    <script src="../js/ckeditor.js"></script>
	<link href="../css/sample.css" rel="stylesheet"> 


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
           				<li class="first_nav"><spring:message code="menu.cont.configs"/> &nbsp;|</li>
           				<li class="second_nav" id='updateId'><spring:message code="menu.sub.create.instruction"/></li>           				          				
             		</ol>
	            </div>
	            <div class= "sub_contain"><spring:message code="label.header.instruction"/></div>
	            <div class= "main_contain">	  
	             <div class="form_area_top">
	             <div class= "inner_container"><spring:message code="label.add.instruction"/></div> 
	            <form class="form-horizontal main_div" name="crtInstruction" autocomplete="on" novalidate enctype="multipart/form-data" data-remote="true"> 	            
	          	
	          	<!--  User Type -->
                 <div class="single_form_item">
                    <div class="col-sm-4 select_option">
                    <label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="label.select.user.type"/></label></div>
                    <div class="col-sm-5 select_option_radio">
                    	<select class="form-control-general" name="userType" id="userType" onchange="enableUsrProfile()">
                    		<option class="form-control-general" value="0"><spring:message code="label.user.type.select.drpDwn"/> </option>
                    		<option class="form-control-general" value="1"><spring:message code="label.user.type.select.drpDwn.ind"/> </option>
                    		<option class="form-control-general" value="3"><spring:message code="label.user.type.select.drpDwn.fac"/> </option>
                    	</select>
                    </div> 
                    <div class="col-sm-3 hidden-ipad">&nbsp;</div>  
                    <div class="clearfix"></div>
                </div>
	          	
	          	
	          	
	          	<!--  User Profile -->
                 <div class="single_form_item">
                    <div class="col-sm-4 select_option">
                    <label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="label.user.select.profile"/></label></div>
                    <div class="col-sm-5 select_option_radio">
                    	<select class="form-control-general" name="inputUserProfileInpt" id="inputUserProfileInpt" onchange="fetchTnC()" disabled>
                    		<option class="form-control-general" value="0"><spring:message code="label.user.select.profile_drpDwn"/> </option>
                    	</select>
                    </div> 
                    <div class="col-sm-3 hidden-ipad">&nbsp;</div>  
                    <div class="clearfix"></div>
                </div>
	          	
	   		<!-- t&c description text field -->
                 <div class="single_form_item" >
                   <div class="col-sm-4 select_option function_group">
                   <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="label.t&c.desc"/></label></div>
                   <div class="col-sm-6 select_option_radio"> 
                   	<div class="hero-unit">
						<textarea cols="80" id="editor1" name="editor1" rows="10"></textarea>
					</div>                   
                   </div> 
                   <div class="clearfix"></div>
                </div> 
                              
                 <div id ="tcAddDiv" >                 
                  <div class="single_form_item">
                     <p><input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Add/Update" id="tcAddBtnId" onclick="saveTC()" />                      
                     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                     <input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Reset" id="tcCancel" style="width: 100px;" onclick="clearAllData()" /></p>
                   </div>	
                  </div>
                 
             
                </form>
	            </div>	            
	         
	            </div>	                                 
	            </div>	            
	            
			</div>
	    	</div>
	    	</div>
	        <!-- Form area end -->
	    <!-- Footer area start -->
	  <jsp:include page="footer.jsp"/> 
	    <!-- Footer area end --> 	          
		
		
		
	<script src="https://code.jquery.com/jquery.js"></script>    
<!-- Include all compiled plugins (below), or include individual files as needed -->
<!-- <script src="../lib/bootstrap.min.js"></script> -->
<script src="../js/termsAndConditions.js"></script>
<script src="../js/common.js"></script>
<script src="../js/styles.js"></script>
    <script>
		CKEDITOR.replace( 'editor1', {
			/*
			 * Style sheet for the contents
			 */
			contentsCss: 'assets/outputxhtml/outputxhtml.css',

			removePlugins: 'elementspath, magicline',
			resize_enabled: false,
			toolbar: [
					[ '-', 'Bold', 'Italic', 'Underline', '-', 'BulletedList', 'NumberedList' ],
					[ 'IndentLeft', 'Outdent','Indent'],
					[ 'Font', 'FontSize' ]
				],


			
		});

	</script>
</body>
</html>