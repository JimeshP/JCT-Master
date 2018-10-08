<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@ taglib uri="http://java.sun.com/jsp/jstl/core" prefix="c" %>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<head>
    <title><spring:message code="menu.popup.ins" /></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <!-- Bootstrap -->
     <link rel="shortcut icon" href="/user/img/crafting_ico.ico" />
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
	if (typeof jQuery == 'undefined') {
    	document.write(unescape("%3Cscript src='/admin/js/latest_10.2_jquery.js' type='text/javascript'%3E%3C/script%3E"));
    }
	</script>	
	<!-- <script type="text/javascript" src="../js/disableRC.js"></script> -->
	<script src="http://code.jquery.com/ui/1.10.3/jquery-ui.js" type="text/javascript"></script>
	<script type="text/javascript">
	//<![CDATA[
        (window.jQuery && window.jQuery.ui && window.jQuery.ui.version === '1.10.3')||document.write('<script type="text/javascript" src="/admin/js/jquery_ui.js"><\/script>');//]]>
    </script>	
    <script src="//cdn.ckeditor.com/4.5.1/standard/ckeditor.js"></script>
	<script src="/admin/js/alertify.min.js"></script>
	<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css">
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
	            
	            <div class="col-sm-10">
	            <div class= "heading_label">
	            	<ol class="breadcrumbAdmin">
           				<li class="first_nav"><spring:message code="reports.breadcrumb.manage.user.home"/> &nbsp;|</li>
           				<li class="first_nav"><spring:message code="menu.cont.configs"/> &nbsp;|</li>
           				<li class="second_nav"><spring:message code="menu.popup.ins"/></li>
             		</ol> 
	            </div>
	             <div class= "sub_contain"><spring:message code="menu.popup.ins"/></div>
	            <div class= "main_contain">	  
	            <div class= "inner_container" id="headerId"><spring:message code="menu.popup.ins"/></div>	                 
	            
	            <div class="form-horizontal main_div">
	           <%--  <div class="single_form_item">
                     <div class="col-sm-4 select_option">
                     <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="menu.sub.select.option"/></label></div>
                     <div class="col-sm-8 sub_text_label select_option_radio">
                       <div class="col-sm-4 radio_text">
                          <input type="radio" name="insPopupFld" id="videoPopFldId" value="V" onclick="manageDiv('videoDivId', 'textInsId', 'Upload Video', 'onlyVideo')" checked = "checked"> <spring:message code="menu.sub.upld.video"/>
                       </div>
                       <div class="col-sm-4 radio_text">
                       	  <input type="radio" name="insPopupFld" id="InsPopFldId" value="I" onclick="manageDiv('textInsId', 'videoDivId', 'Text Instruction', 'onlyText')"> Text Instruction
                       </div>
                       <div class="col-sm-4 radio_text">
                       	  <input type="radio" name="insPopupFld" id="textVideoPopFldId" value="TV" onclick="manageDiv('videoDivId', 'textInsId', 'Upload Text And Video', 'textVideo')"> Text And Video 
                       </div>
                     </div>  
                     <div class="clearfix"></div>
                     </div> --%>
					 
					 <div id="videoDivId" style="display:block;">
		            <form modelAttribute="uploadForm" method="POST" action="/admin/uploadVideo" enctype="multipart/form-data" onsubmit="return validateVideoFile()">
	       				
	       				<!--  User Profile -->
		                  <div class="single_form_item">
		                     <div class="col-sm-4 select_option">
		                     <label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="label.user.select.profile"/></label></div>
		                     <div class="col-sm-5 select_option_radio">
		                     	<select class="form-control-general" name="inputUserProfileInpt" id="inputUserProfileInpt" onchange="disablePageFieldVideo(this)">
		                     		<option class="form-control-general" value="0"><spring:message code="label.user.select.profile_drpDwn"/> </option>
		                     	</select>
		                     </div> 
		                     <div class="col-sm-3 hidden-ipad">&nbsp;</div>  
		                     <div class="clearfix"></div>
		                 </div>

	       				<div class="single_form_item">
						<div class="col-sm-4 ">
		                     <label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="label.select.related.page"/></label>
		                </div>
		                <div class="col-sm-5 select_option_radio">
		                   	<select class="form-control-general" name="fileType" id="fileType" onchange="changeVideo(this)">
		                   		<option class="form-control-general" value="0"><spring:message code="label.select.video.type"/></option>  
								<option class="form-control-general" value="BS"><spring:message code="label.mappingYourself"/></option>      
								<option class="form-control-general" value="AS"><spring:message code="label.menu.afterdiagram"/></option>       
		                   	</select>
						</div> 
						<div class="col-sm-3">&nbsp;</div>
						<div class="clearfix"></div>
	                </div>
	                
	                
	                <%--  <div class="single_form_item">
	                   <div class="col-sm-4 select_option">
	                   <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="label.instruction.header"/></label></div>
	                   <div class="col-sm-6 select_option_radio"><input type="text" class="form-control-general-form-redefined" maxlength="80" 
	                   					value="" name="instructionHeaderText" id="instructionHeaderText" placeholder="Instruction Header"></div> 
	                   <div class="col-sm-2 hidden-ipad">&nbsp;</div>  
	                   <div class="clearfix"></div>
	               	 </div>   --%>
	                
	                
	                <div id="videoTextDivId">	                
	                 <div class="single_form_item" >
					   <div class="col-sm-4 select_option function_group">
					   <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="label.instructionText.before.video"/></label></div>
					   <div class="col-sm-6 select_option_radio"> 
					   <div class="hero-unit">
						   <textarea cols="40" id="inputText" name="inputTextBeforeVideo" rows="5"></textarea>
					   </div>                   
					   </div> 
					   <div class="clearfix"></div>
					</div> 				
	                </div>
	                
	                
	       				<div class="single_form_item">
	       					<div class="col-sm-4 ">
	                      	<label for="inputOC" class="col-sm-12 control-label text_label"><spring:message code="menu.sub.file.upld.video"/></label></div>	
	                       	<div class="col-sm-6 select_option_radio">
	                     		<input type="file" class="filestyle" data-classButton="btn choose_file" name="file" id="filename" accept=".mp4,.ogv,.webm " onchange="getFileSize('filename')">
	                     	</div>  
	                     	<div class="clearfix"></div>  
	       			 	</div>
	       			 	
	       			 	
	       			<div id="afterVideoTextDivId">	                
	                 <div class="single_form_item" >
					   <div class="col-sm-4 select_option function_group">
					   <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="label.instructionText.after.video"/></label></div>
					   <div class="col-sm-6 select_option_radio"> 
					   <div class="hero-unit">
						   <textarea cols="80" id="inputTextAfterVideo" name="inputTextAfterVideo" rows="10"></textarea>
					   </div>                   
					   </div> 
					   <div class="clearfix"></div>
					</div> 				
	                </div>
	       			 
	       			 	<div class="single_form_item">
	                    	 <p><input type="submit" id="onetDataSaveBtnId" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Save" />
							 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							 <input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Reset" id="resetuploadVDO" onclick="resetuploadVDOfield()"></p>
	                   	</div>
	                   	<input type="hidden" value="" id="hiddenFileName" name="hiddenFileName"/>
	                </form>  
					</div>
					
					<%-- <div id="textInsId" style="display:none;">
   			    <form method="POST">
	       				<!--  User Profile -->
		            <div class="single_form_item">
		                <div class="col-sm-4 select_option">
		                    <label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="label.user.select.profile"/></label></div>
		                    <!--  <div class="col-sm-5 select_option_radio"> -->
		                      <div class="col-sm-5 select_option">
		                     	<select class="form-control-general align_ipad_text" name="inputUserProfileInpt1" id="inputUserProfileInpt1" onchange="disablePageFieldText(this)">
		                     	<!--<select class="form-control-general align_ipad_text" name="inputUserProfileInpt1" id="inputUserProfileInpt1" onchange="changeTextIns(this)">  -->
		                     		<option class="form-control-general" value=""><spring:message code="label.user.select.profile_drpDwn"/> </option>
		                     	</select>
		                     </div> 
		                     <div class="col-sm-3 hidden-ipad">&nbsp;</div>  
		                <div class="clearfix"></div>
		            </div>
	                 
	                <div class="single_form_item">
						<div class="col-sm-4 ">
		                     <label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="label.select.related.page"/></label>
		                </div>
		                <div class="col-sm-5 select_option">
		                   	<select class="form-control-general align_ipad" name="fileType1" id="fileType1" onchange="changeTextIns(this)">
		                   		<option class="form-control-general" value="0"><spring:message code="label.select.video.type"/></option>  
								<option class="form-control-general" value="BS"><spring:message code="label.mappingYourself"/></option>      
								<option class="form-control-general" value="AS"><spring:message code="label.menu.afterdiagram"/></option>       
		                   	</select>
						</div> 
						<div class="col-sm-3">&nbsp;</div>
						<div class="clearfix"></div>
	                </div>
	                
	                
		            <div class="single_form_item">
	                   <div class="col-sm-4 select_option">
	                   <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="label.instruction.header"/></label></div>
	                   <div class="col-sm-6 select_option_radio"><input type="text" class="form-control-general-form-redefined" maxlength="80" value="" id="instHeaderId"></div> 
	                   <div class="col-sm-2 hidden-ipad">&nbsp;</div>  
	                   <div class="clearfix"></div>
	               	 </div>    
	                
	                <div class="single_form_item" >
					   <div class="col-sm-4 select_option function_group">
					   <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="label.instructionText"/></label></div>
					   <div class="col-sm-8 select_option_radio"> 
					   <div class="hero-unit">
						   <textarea cols="80" id="editor1" name="editor1" rows="10"></textarea>
					   </div>                   
					   </div> 
					   <div class="clearfix"></div>
					</div> 
                              
					<div id ="instructionAddDiv" >                 
					 <div class="single_form_item">
						 <p><input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Save" id="addInstructionBtnId" onclick="saveInstruction()" />                      
						 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						 <input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Reset" id="cancelInstructionBtnId" onclick="cancelInstruction()" /></p>
					 </div>	
					</div>
                 		
	       		</form>
   			    </div>	 --%>
				
                 </div>
	            
	            <div class="existing_data_div" style="display:none;" id="existing_data_div">    
	            <div class="single_form_item">
	            <div class="col-sm-5"><div  id="existing_list_Id" class= "existing_list"></div></div>
	            <div class="col-sm-6">
	            <div  id="remove_video_Id" class= "existing_list">
	            <input type="button" class="btn_admin btn_admin-sm remove_video_btn" value="Remove Video" id="removeVideoId" onclick="removeVideo()">
	            </div></div>
	            <div class="clearfix"></div>
	            </div>         
					
					<div id="videoSection" class="table_div_scrolls">
						
					</div> 
					<div id="instructionSection" class="table_div_scrolls">
						
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
 <script src="/admin/js/bootstrap.min.js"></script>
<script src="/admin/js/bootstrap-filestyle.js"></script>
<script src="/admin/js/common.js"></script>
<script src="/admin/js/popupInfoInstruction.js"></script>
<script src="/admin/js/jquery.tablesorter.js"></script>
<script src="/admin/js/equalHeight.js"></script>
<script type="text/javascript">
$(":file").filestyle();
</script>
<script type="text/javascript">
var msg = '<c:out value="${message}" />';
if(msg != "") {
	alertify.alert(msg);
	document.getElementById('videoSection').innerHTML = "<div align='center'><br /><br /><br />"+msg+"<div>";
}
</script>	
<script>	
	try {
		CKEDITOR.replace( 'inputText', {
			contentsCss: '/admin/view/assets/outputxhtml/outputxhtml.css',

			removePlugins: 'elementspath, magicline',
			resize_enabled: false,
			fontSize_sizes: '8/8px;9/9px;10/10px;11/11px;12/12px;14/14px;16/16px;18/18px;20/20px;22/22px;24/24px;26/26px;28/28px;36/36px;48/48px;54/54px;60/60px',
			toolbar: [
					[ '-', 'Bold', 'Italic', 'Underline', '-', 'BulletedList', 'NumberedList' ],
					[ 'IndentLeft', 'Outdent','Indent'],
					[ 'Font', 'FontSize' ]
				],


			
		});
	} catch(err) {
	    
	}
	try {
		CKEDITOR.replace( 'inputTextAfterVideo', {
			contentsCss: '/admin/view/assets/outputxhtml/outputxhtml.css',

			removePlugins: 'elementspath, magicline',
			resize_enabled: false,
			fontSize_sizes: '8/8px;9/9px;10/10px;11/11px;12/12px;14/14px;16/16px;18/18px;20/20px;22/22px;24/24px;26/26px;28/28px;36/36px;48/48px;54/54px;60/60px',
			toolbar: [
					[ '-', 'Bold', 'Italic', 'Underline', '-', 'BulletedList', 'NumberedList' ],
					[ 'IndentLeft', 'Outdent','Indent'],
					[ 'Font', 'FontSize' ]
				],


			
		});
	} catch(err) {
	    
	}
	</script>
</body>
</html>