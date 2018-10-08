<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<head>
    <title><spring:message code="label.app.settings" /></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <!-- Bootstrap -->
    
    <link href="../css/bootstrap.css" rel="stylesheet">
    <link href="../css/style.css" rel="stylesheet">
	<!--  <link href="https://s3-us-west-2.amazonaws.com/jobcrafting/css/commonStyle.css" rel="stylesheet">-->
    <link rel="stylesheet" href="../css/alertify.core.css" />
	<link rel="stylesheet" href="../css/alertify.default.css" id="toggleCSS" />
	<meta name="viewport" content="width=device-width">
	<link rel="shortcut icon" href="/admin/img/crafting_ico.ico" />
	
	 
	
    
	<link href="../css/datepicker.css" rel="stylesheet">
    <link href="../css/bootstrap-multiselect.css" rel="stylesheet">

	<!-- Jcrop css inclusion -->
	<link rel="stylesheet" type="text/css" href="../css/jquery.Jcrop.min.css" />
	<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.7.2/jquery.min.js"></script>
	
	
    <!-- HTML5 Shim and Respond.js IE8 support of HTML5 elements and media queries -->
    <!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
      <script src="https://oss.maxcdn.com/libs/respond.js/1.3.0/respond.min.js"></script>
    <![endif]-->
	
	<!-- <script src="http://code.jquery.com/jquery-latest.min.js"></script> -->
	
	
	<!-- <script src="http://code.jquery.com/ui/1.10.3/jquery-ui.js" type="text/javascript"></script>
	<script type="text/javascript">
	//<![CDATA[
        (window.jQuery && window.jQuery.ui && window.jQuery.ui.version === '1.10.3')||document.write('<script type="text/javascript" src="../js/jquery_ui.js"><\/script>');//]]>
    </script>	 -->
	<script src="../js/alertify.min.js"></script>
	<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css">
	<script type="text/javascript" src="../lib/jquery.js"></script>
	<script type="text/javascript" src="../lib/spine.js"></script>
	<script type="text/javascript" src="../lib/ajax.js"></script>	
	<script src="http://code.jquery.com/jquery-latest.min.js"></script>
	
    <script type="text/javascript" src="../js/appSettings.js"></script>
	
	<!-- Color picker sof -->	
	   <script src="http://code.jquery.com/jquery-1.9.1.min.js"></script>
<script src="http://code.jquery.com/ui/1.10.2/jquery-ui.js"></script>
<!-- Jcrop Js plugin inclusion -->
<script src="../lib/jquery.Jcrop.min.js"></script>
	<link rel="stylesheet" href="../css/spectrum.css" />
	<script src="../js/spectrum.js"></script>
	<!-- Color picker eof -->
	 <script src="../js/bootstrap-filestyle.js"></script>
	 
	 
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
	            
	            <div class="col-sm-10">
	            <div class= "heading_label">
	            	<ol class="breadcrumbAdmin">
           				<li class="first_nav"><spring:message code="reports.breadcrumb.manage.user.home"/> &nbsp;|</li>
           				<li class="second_nav"><spring:message code="label.app.settings"/></li>
             		</ol>
	            </div>
	            <div class= "sub_contain"></div>
	            <div class= "main_contain">	  
	            <form class="form-horizontal" name="actionPlanRprt" autocomplete="on" novalidate enctype="multipart/form-data" data-remote="true"> 
	           <div class="single_form_item"><br/></div>
	           
	           <!-- Header Logo -->
	           <div class="single_form_item">
                   <div class="col-sm-4 ">
                       <label for="csvFileLbl" class="col-sm-12 control-label text_label_myAccnt left_text"><spring:message code="label.app.settings.header.logo"/></label></div>
                   <div class="col-sm-4">
						<input class="form-control-general-form-color-picker" type="text" id="filePath" disabled/></div>
				   <div class="col-sm-4"><a href="#" class="change-logo" data-toggle="modal" data-target="#myModal-frgtpass" id="changeLogo">Change Logo</a></div>
				   <div class="clearfix"></div>
               </div>
         
              
	           
	           <!-- Header Color -->
	           <div class="single_form_item">
                     <div class="col-sm-4 ">
                     <label for="csvFileLbl" class="col-sm-12 control-label text_label_myAccnt left_text"><spring:message code="label.app.settings.header.color" /></label>
                     </div>
                     <div class="col-sm-5">
                     <input class="form-control-general-form-color-picker" type="text" spellcheck="false" placeholder="Header Color" id="headerColorInptId" onkeypress="return false;"/>
                     </div>
                     
                    <div class="clearfix"></div>
               </div> 
			   
               
               <!-- Footer Color -->
	           <div class="single_form_item">
                     <div class="col-sm-4 ">
                     <label for="csvFileLbl" class="col-sm-12 control-label text_label_myAccnt left_text"><spring:message code="label.app.settings.footer.color" /></label>
                     </div>
                     <div class="col-sm-5">
                     <input class="form-control-general-form-color-picker" type="text" spellcheck="false" placeholder="Footer Color" id="footerColorInptId" onkeypress="return false;"/>
                     </div>
                     
                    <div class="clearfix"></div>
               </div>
               
               <!-- Sub Header Color -->
	           <div class="single_form_item">
                     <div class="col-sm-4 ">
                     <label for="csvFileLbl" class="col-sm-12 control-label text_label_myAccnt left_text"><spring:message code="label.app.settings.sub.header.color" /></label>
                     </div>
                     <div class="col-sm-5">
                     <input class="form-control-general-form-color-picker" type="text" placeholder="Sub Header Color" spellcheck="false" id="subHeaderColorInptId" onkeypress="return false;"/>
                     </div>                     
                    <div class="clearfix"></div>
               </div> 
               
               
               <!-- Instruction Panel Color -->
	           <div class="single_form_item">
                     <div class="col-sm-4 ">
                     <label for="csvFileLbl" class="col-sm-12 control-label text_label_myAccnt left_text"><spring:message code="label.app.settings.instruction.panel.color" /></label>
                     </div>
                     <div class="col-sm-5">
                     <input class="form-control-general-form-color-picker" type="text" placeholder="Instruction Panel Color" spellcheck="false" id="instructionPanelColorInptId" onkeypress="return false;"/>
                     </div>
                    <div class="clearfix"></div>
               </div>
               
               <!-- Instruction Panel Text Color -->
	           <div class="single_form_item">
                     <div class="col-sm-4 ">
                     <label for="csvFileLbl" class="col-sm-12 control-label text_label_myAccnt left_text"><spring:message code="label.app.settings.instruction.panel.text.color" /></label>
                     </div>
                     <div class="col-sm-5">
                     <input class="form-control-general-form-color-picker" type="text" placeholder="Instruction Panel Text" spellcheck="false" id="instructionPanelTextColorInptId" onkeypress="return false;"/>
                     </div>
                    <div class="clearfix"></div>
               </div> 
	           
	            
                  
                  <div class="single_form_item"><br/></div>
                   <div class="single_form_item"><br/></div>
                   <div class="single_form_item">
                       <p><input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Preview" id="Preview" onclick="preview()"/>
                       &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                       <input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Save" id="save" onclick="saveChanges()"/></p>
                   </div>
                   
                   <div class="single_form_item"><br/></div>
                   <div class="single_form_item"><br/></div>
                   <div class="single_form_item"><br/></div>
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
		
		
<!-- Include all compiled plugins (below), or include individual files as needed -->
<script src="../js/bootstrap.min.js"></script>


<script type="text/javascript" >

	$(":file").filestyle();

</script>

<div class="modal fade" id="myModal-frgtpass" tabindex="-1">
  <div class="modal-dialog">
    <div class="modal-content admin-modal-box">
      <div class="modal-header">
        <button type="button" class="close modal-close" data-dismiss="modal">&times;</button>
        <h4 class="modal-title modal-heading" id="myModalLabel"><spring:message code="lbl.change.logo" /></h4>
      </div>
	  <hr>
			<div class="modal-body">
                  
					<div class="single_form_item">
						 <div class="col-md-4 enter_email_lbl">
						 <label for="inputFN" class="col-md-12 control-label align-right"><spring:message code="lbl.browse.image" /></label></div>
						 <div class="col-md-8">
							<input type="file" id="imageFile" class="image-file" data-classButton="btn choose_file" accept="image/*" required>
							<div id="BrowserVisible" class="browser-visible"><input type="text" id="fileField" class="form-control-general-form-color-picker"/></div>
						 </div> 
						
						 <div class="clearfix"></div>
					</div>
					<div class="single_form_item" id="imgCaption">
					<hr>
						 <div class="col-md-6 enter_email_lbl">
						 <label for="inputFN" class="col-md-12 control-label align-centre"><spring:message code="lbl.source.image" /></label></div>
						 <div class="col-md-6 enter_email_lbl">
						 <label for="inputFN" class="col-md-12 control-label align-centre"><spring:message code="lbl.preview" /></label></div>						  
						 <div class="clearfix"></div>
					</div>
					<div class="single_form_item">
						 <div class="col-md-6 enter_email_lbl">
						 <div class="col-md-12" id="imageInput"></div></div>
						 
						 <div class="col-md-5">
						 <img class="col-md-12" id="imageOutput" /></div>					  
						 <div class="clearfix"></div>
					</div>
					
		
					
			</div>
		<hr>
		<div class="modal-footer">
			<input type="submit" class="pop-up-btn" value="Crop Selected" id="crop">
		</div>	  
    </div>
  </div>
</div>
</body>
</html>