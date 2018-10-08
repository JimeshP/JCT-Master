<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<head>
    <title><spring:message code="menu.sub.create.instruction" /></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <!-- Bootstrap -->
    <link href="/admin/css/bootstrap.css" rel="stylesheet">
    <link href="/admin/css/style.css" rel="stylesheet">
    <link rel="stylesheet" href="/admin/css/alertify.core.css" />
	<link rel="stylesheet" href="/admin/css/alertify.default.css" id="toggleCSS" />
	<meta name="viewport" content="width=device-width">
   <link rel="shortcut icon" href="/admin/img/crafting_ico.ico" />
	<script type="text/javascript" src="/admin/lib/jquery.js"></script>
	<script type="text/javascript" src="/admin/lib/spine.js"></script>
	<script type="text/javascript" src="/admin/lib/ajax.js"></script>
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
	<script src="/admin/js/alertify.min.js"></script>
	<script type="text/javascript" src="/admin/js/jquery-1.3.2.js"></script>
    <script type="text/javascript" src="/admin/js/jquery-ui-1.7.2.custom.min.js"></script>
    <link rel="Stylesheet" type="text/css" href="/admin/css/jqueryui/ui-lightness/jquery-ui-1.7.2.custom.css" />

    <script src="/admin/js/ckeditor.js"></script>
	<link href="/admin/css/sample.css" rel="stylesheet"> 


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
	             
	           <div class="form-horizontal main_div">
	           <form modelAttribute="uploadInstructionForm" method="POST" action="/admin/uploadInstructionVideo" enctype="multipart/form-data" onsubmit="return validateVideoFile()"> 	            
	          	<!--  User Profile -->
                  <div class="single_form_item">
                     <div class="col-sm-4 select_option function_group">
                     <label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="label.user.select.profile"/></label></div>
                     <div class="col-sm-5 select_option_radio">
                     	<select class="form-control-general" id="inputUserProfileInpt" name="inputUserProfileInpt" onchange="disablePageField(this)">
                     		<option class="form-control-general" value=""><spring:message code="label.user.select.profile_drpDwn"/> </option>
                     	</select>
                     </div> 
                     <div class="col-sm-3">&nbsp;</div>  
                     <div class="clearfix"></div>
                </div>  
                
                <!--  Page -->
                <div class="single_form_item">
                   <div class="col-sm-4 select_option function_group">
                   <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="label.select.related.page"/></label></div>
                   <div class="col-sm-5 select_option_radio">
                     	<select class="form-control-general" id="inputPageInpt" name="inputPageInpt" onchange="fetchInstructionData(this)">
                     		<option class="form-control-general" value="0"><spring:message code="label.select.page"/> </option>
                     		<option class="form-control-general" value="Creating Before Sketch">Creating Before Sketch</option>
                     		<option class="form-control-general" value="Mapping Yourself">Mapping Yourself</option>
                     		<option class="form-control-general" value="Creating After Diagram">Creating After Diagram</option>
                     	</select>
                     </div> 
                   <div class="col-sm-3">&nbsp;</div>  
                   <div class="clearfix"></div>
                </div>
                
                <div class="single_form_item">
                   <div class="col-sm-4 select_option function_group">
                   <label for="inputFN" class="col-sm-12 control-label text_label">Include video with instructions:</label></div>
                   <div class="col-sm-5 instruction_radio">
                     	<input type="checkbox" id="incVideo" name="incVideo" onclick="toogleCheckbox()">
                   </div> 
                    <div class="col-sm-3">&nbsp;</div>
                                
                   <div class="clearfix"></div>
                </div>
                
                
                <div class="single_form_item" id = "videoSectionDiv" style="display: none;">
                   <div class="col-sm-4 select_option function_group">   </div>            
                  
                  <div class="col-sm-6">
                   	<div id="videoSection" class="table_div_scrolls"></div>
					</div> 
					
                   <div class="clearfix"></div>
                </div>
                
                 <div class="single_form_item" id="videoDivId" style="display: none;">
	       		 	<div class="col-sm-4 ">
	                <label for="inputOC" class="col-sm-12 control-label text_label"><spring:message code="menu.sub.file.upld.video"/></label></div>	
	                <div class="col-sm-6 select_option_radio">
	                	<input type="file" class="filestyle" data-classButton="btn choose_file" name="file" id="filename" accept=".mp4,.ogv,.webm " onchange="getFileSize('filename')">
	                </div>  
	                <div class="clearfix"></div>  
	       	     </div>                       
                 
                 <div class="single_form_item" >
                   <div class="col-sm-4 select_option function_group">
                   <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="label.instructionText"/></label></div>
                   <div class="col-sm-6 select_option_radio"> 
                   	<div class="hero-unit">
						<!-- <textarea class="textarea" id='myTextAreaId' placeholder="Enter text ..." style="width: 530px; height: 300px"></textarea>  -->
						<textarea cols="80" id="editor1" name="editor1" rows="10"></textarea>
						
						</div>                   
                   </div> 
                   <div class="clearfix"></div>
                </div> 
                              
                 <div id ="instructionAddDiv" >                 
                  <div class="single_form_item">
                     <p><input type="submit" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Add" id="addInstructionBtnId" style="width: 100px;" />                      
                     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                     <input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Reset" id="cancelInstructionBtnId" style="width: 100px;" onclick="cancelInstruction()" /></p>
                   </div>	
                  </div>
                  
                   <div id ="instructionUpdateDiv" style=" display: none;">                 
                  <div class="single_form_item">
                     <p><input type="submit" class="btn_admin btn_admin-update btn_admin-sm search_btn" value="Update" id="updateInstructionBtnId" style="width: 100px;" />                      
                     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                     <input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Reset" id="cancelInstructionBtnId" style="width: 100px;"  onclick="cancelInstruction()" /></p>
                   </div>	
                  </div>
             	<input type="hidden" value="" id="hiddenFileName" name="hiddenFileName"/>
             	<input type="hidden" value="" id="hiddenProfileInput" name="hiddenProfileInput"/>
             	<input type="hidden" value="" id="hiddenProfileValInput" name="hiddenProfileValInput"/>
             	<input type="hidden" value="" id="hiddenPageInput" name="hiddenPageInput"/>
             	<input type="hidden" value="" id="createdBy" name="createdBy"/>
             	<!-- <div id="videoSection"></div> -->
                </form>
                </div>
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
<script src="/admin/js/instruction.js"></script>
<script src="/admin/js/common.js"></script>
<script src="/admin/js/styles.js"></script>
    <script>
		CKEDITOR.replace( 'editor1', {
			/*
			 * Style sheet for the contents
			 */
			contentsCss: '/admin/css/outputxhtml.css',

			removePlugins: 'elementspath, magicline',
			resize_enabled: false,
			toolbar: [
					[ '-', 'Bold', 'Italic', 'Underline', '-', 'BulletedList', 'NumberedList' ],
					[ 'IndentLeft', 'Outdent','Indent'],
					[ 'Font', 'FontSize' ]
				],


			
		});

	</script>
	<script src="/admin/js/bootstrap-filestyle.js"></script>
	<script type="text/javascript">
$(":file").filestyle();
</script>
<script type="text/javascript">

var msg = "${message}";
if(msg != "" && msg != null && msg != " ") {
	alertify.alert(msg);
	document.getElementById('videoSection').innerHTML = "<div align='center'><br /><br /><br />"+msg+"<div>";
}
</script>	
</body>
</html>