<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<head>
    <title><spring:message code="menu.sub.renew.user" /></title>
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
           				<li class="second_nav" id='updateId'><spring:message code="menu.sub.renew.user"/></li>        				          				
             		</ol>
	            </div>
	            <br>
	            <div class= "sub_contain"></div>
	            <div class= "main_contain">	  
	             <div class="form_area_top">
	             <div class= "inner_container"><spring:message code="menu.sub.renew"/></div> 
	            <form class="form-horizontal main_div main_div_new" name="addNewUser" autocomplete="on" novalidate enctype="multipart/form-data" data-remote="true"> 
				 				  
				 <!--  Renew Subscription-->
	             <div class="single_form_item">
                     <div class="col-sm-3 ">
                     <label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="label.user.renew"/></label></div>
                     <div class="col-sm-2">
                     	<input id="totalRenewedUserId" class="form-control-general-form-redefined disable_style width_60 renewSubChange" type="text" autofocus="" onkeypress="return alphaNumericOnly(event)" placeholder="" value="" maxlength="50" disabled="true"></input>
                     </div>
					
					 <div class="col-sm-4 ">
                     <label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="label.user.renew.available"/></label></div>
                     <div class="col-sm-2">
                     	<input id="availRenewedUserId" class="form-control-general-form-redefined disable_style width_60 renewSubChange" type="text" autofocus="" onkeypress="return alphaNumericOnly(event)" placeholder="" value="" maxlength="50" disabled="true"></input>
                     </div>
                      <div class="col-sm-1"></div>
					 </div>
                     <div class="clearfix"></div> 
                     
                     <br> 
                     
               <%--   <div class="single_form_item renew_note_area">
                   <div class="col-sm-10 renew_note">
                   <spring:message code="label.note.renew.user"/>
                   </div>
                    <div class="col-sm-2"></div>
                   
                 </div>   --%>                                               
                </form>
	            </div>

	            
	            <!--  <div class="existing_data_div">	             
	             <div id="existingRenewList" class= "existing_list" >
				 	<table width="95%">
				 	<tbody><tr>
				 		<td align="left">
				 			Existing User List
				 		</td>
				 		<td align="right">
				 			<a class="activate_style" href="#" onclick="goToNextPage()">Bulk Renew</a>
				 		</td>
				 	</tr>
				 </tbody></table>
				 </div>
				 <div id="bulkRenewListId" class= "existing_list" >
				 	<a href='bulkAction.jsp' class='activate_style'>Bulk Renew</a>
				 </div>
				 
	            <div id="existingUsersTableId">
	            	<div align="center">
	            		<img src="../img/dataLoading.GIF" />
	            		<p> Loading Existing Users</p>
	            	</div>
	            </div> 
	            </div> -->
	            
	             <div class="existing_data_div" id='existingRenewList' style="display: none;">	   
	            <br>          
	             <div class= "existing_list">Existing User List</div>
	            	 <div id="mainPopulationDiv"></div>
	            	 <div id='buttonDivId' style="display: none;"></div>
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
<script src="../js/bootstrap-filestyle.js"></script>
<script src="../js/common.js"></script>
<script src="../js/renewSubscription.js"></script>
<script src="../js/bulkRenewFacilitator.js"></script>
<script src="../js/sortTable.js"></script>
<!--<script src="../js/createUserList.js"></script>-->
<script type="text/javascript">
$(":file").filestyle();
</script>

<div class="modal fade" id="myModal" tabindex="-1" data-backdrop="static">
  <div class="modal-dialog" >
    <div class="modal-content modal_width modal_height">
      <div class="new_modal_header modal-header">
        <button type="button" class="custom_close_btn" data-dismiss="modal"></button>      
        <h4 class="modal-title new_modal_title header_background" id="myModalLabel"></h4>
        <div class="" id="modelHeaderId"></div>
      </div>
         <!-- <div id="instruct_header" class="page_title_instruct"></div>    -->
             <!--  <div class="modal-body video_style" id="instruction_panel_video"></div>           
              <div class="instruction_panel text_style" id="instruction_panel_text"></div>
              <div class="instruction_panel text_style" id="instruction_panel_video_text">
                     <div class="modal-body video_style" id="video_instruction_id"></div>  -->
                     <div class="registered_div_header" id="info-div-header" style="display:none">Non Existing User(s)
                      </div>
                      <div class="registered_div" id="info-div" style="display:none"></div><br>
                    <div class="invalid_div_header" id="invalid_div_header" style="display:none">Invalid User(s)
                     <div class="invalid_div" id="invalid-div" style="display:none"></div>  
                      </div>
                     
                     <!-- <div class="instruction_panel text_style" id="text_instruction_id"></div>  -->
              </div>
      <div class="modal-footer">
      </div>
    </div>
  </div>
  
<div class="modal fade" id="bulkRenewModal" tabindex="-1" data-backdrop="static">
    <div class="modal-dialog_renew" >
    <div class="modal-content_renew">
    <div class="new_modal_header modal-header">
        &nbsp;      
    </div>
    	<div align="center" class="renew_user_list">   
      		<span id="renewUser"></span>
      	</div>  
      	<br>   	
      	<div align="center">   
      		<!-- <nav class="alertify-buttons">
      			<button class="alertify-button alertify-button-ok" id="alertify-oks1" onclick="closeMyModel1()">OK</button>
      		</nav> -->      		
      		<nav class="alertify-buttons">
      			<button class="alertify-button alertify-button-ok" id="alertify-oks" onclick="renewBulkUserAfterSuccess()" autofocus>OK</button>
      			<button class="alertify-button alertify-button-cancel" id="alertify-cancels" onclick="closeMyModel()">Cancel</button>
      		</nav>
      	</div>
      	<br>
      	<div align="center">   
      		<span id="tableData"></span>
      	</div>
      	<br>
    </div>
  	</div>
	</div>
</body>
</html>