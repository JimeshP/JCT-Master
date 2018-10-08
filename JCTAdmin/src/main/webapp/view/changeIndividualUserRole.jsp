<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<head>
    <title><spring:message code="menu.sub.change.role" /></title>
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
	<script src="http://code.jquery.com/jquery-latest.min.js"></script>
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
           				<li class="second_nav" id='updateId'><spring:message code="menu.sub.change.role"/></li>        				          				
             		</ol>
	            </div>
	            <br>
	            <div class= "sub_contain"></div>
	            <div class= "main_contain">	  
	             <div class="form_area_top">
	             <div class= "inner_container"><spring:message code="menu.sub.change.role"/></div> 
	            <form class="form-horizontal main_div main_div_new " name="addNewUser" autocomplete="on" novalidate enctype="multipart/form-data" data-remote="true"> 	            
	                  
	                  
	            <!-- Note -->     	            	            
	            <div class="single_form_item renew_note_area">
	            <div class="col-sm-1"></div>      
                   <div class="col-sm-10 renew_note">
                   <spring:message code="label.note.add.new.facilitator"/>
                   </div>
                    <div class="col-sm-1"></div>                  
                 </div>
                   
                 <!--  Change Role -->
				<div class="single_form_item">
                     <div class="col-sm-4 ">
                     <label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="menu.sub.select.option"/></label></div>
                     <div class="col-sm-5 sub_text_label ipad_alignment">
                     	<div class="col-sm-6 radio_text">
                            <input type="radio" name="preference" id="existingUserId" value="E" onclick="toogleSelectionDiv('E')" checked = "checked"> <spring:message code="label.existing"/>
                       	</div>
                       	<div class="col-sm-6 radio_text">
                       		<input type="radio" name="preference" id="newUserId" value="N" onclick="toogleSelectionDiv('N')"> <spring:message code="label.new"/>
                       	</div>
                     </div> 
                     <div class="col-sm-3">&nbsp;</div>  
                     <div class="clearfix"></div>
                </div>

				<!-- Existing User div-->
				<div id="existingUserDiv" class="single_form_item">
                     <div class="col-sm-4 ">
                     <label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="label.new.facilitator"/></label></div>
                     <div class="col-sm-5">
                     	<select class="form-control-general-form-redefined" id="existingUserInputId">
                     		<option class="form-control-general" value=""><spring:message code="label.user.select.user_drpDwn"/></option>  
						</select>
                     </div> 
                     <div class="col-sm-3">&nbsp;</div>  
                     <div class="clearfix"></div>
                </div>
                
				<!-- New User div-->
					<div id="newUserDiv" style="display: none;">
					<div class="single_form_item">
                     <div class="col-sm-4 ">
					 <label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="label.new.facilitator.email"/></label>
                     </div>
                     <div class="col-sm-5">
                     	<input id="newUserInputId" class="form-control-general-form-redefined form-control-general-form-redefined-white" type="text" autofocus="" placeholder="" value="" maxlength="60"></input>
                     </div> 
                     	<div class="col-sm-3">&nbsp;</div>
                     	<div class="clearfix"></div>
                   </div>
					</div>	
					
					<div id="existingUserAddDiv">
					<div class="single_form_item">
                     <p><input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Update" id="existingUserBtnId" onclick="updateUserToFacilitator()" /></p>
					</div>
					</div>

					<div id="newUserAddDiv" style="display: none;">
					<div class="single_form_item">
                     <p><input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Update" id="newUserBtnId" onclick="addNewFacilitator()" /></p>
					</div>
					</div>					
                </form>
	            </div>
	            	            
	             <div class="existing_data_div">	             
	             <div id="existingChangeRoleList" class= "existing_list" >
				 <table width='95%'>
				 	<tr>
				 		<td align="left">
				 			Existing Facilitator List
				 		</td>
				 		<td align="right">
				 			&nbsp
				 		</td>
				 	</tr>
				 </table>
				 </div>
	            <div id="existingUsersTableId">
	            	<div align="center">
	            		<!--<img src="../img/dataLoading.GIF" />
	            		<p> Loading Existing Users</p>-->
	            	</div>
	            </div> 
	            </div>
	            </div>
	                                 
	            </div>	            
	            
				</div>
	    	</div>
	    	<div class="loader_bg" style="display:none"></div>
	        <div class="loader" style="display:none"><img src="../img/Processing.gif" alt="loading"></div>
	    	</div>
	     <!-- Form area end -->
	    <!-- Footer area start -->
	    <jsp:include page="footer.jsp"/>
	    <!-- Footer area end --> 	          
		
		
<script src="https://code.jquery.com/jquery.js"></script>    
<!-- Include all compiled plugins (below), or include individual files as needed -->
<script src="../js/bootstrap.min.js"></script>
<script src="../js/common.js"></script>
<script src="../js/sortTable.js"></script>
<script src="../js/changeIndividualUserRole.js"></script>
</body>
</html>