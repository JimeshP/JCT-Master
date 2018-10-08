<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<head>
    <title><spring:message code="menu.sub.renewSubscribe.user" /></title>
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
	<!-- <script type="text/javascript" src="../js/disableRC.js"></script> -->
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
           				<li class="second_nav" id='updateId'><spring:message code="menu.sub.renewSubscribe.user"/></li>
             		</ol>
	            </div>
	            <div class= "sub_contain"><spring:message code="menu.sub.renewSubscribe.user"/></div>
	            <div class= "main_contain">
	            <div class= "inner_container" id=""><spring:message code="menu.sub.renewSubscribe.user"/></div>	
	            
	            
	            
	            <form class="form-horizontal main_div" autocomplete="on" novalidate enctype="multipart/form-data" data-remote="true" id="searchid" style="display: none; padding-bottom: 3%;">
	            	
	            	<div class="single_form_item"> 
	            	<div id="facilitatorIdDiv" style="display: none;">
	                     <div class="col-sm-3">
	                     <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="label.emailId"/></label>
						 </div>
	                     <div class="col-sm-3">
	                     	<input type="text" class="form-control-general" id="facilitatorIdSearch" />
	                     </div>
					</div>
                     <div class="col-sm-3">
                     	<input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Search" id="searchFaci" style="margin-top: 0%;"/>
					 </div>
                    						
	            	</div>
	            	<div class="single_form_item"><br/></div>
	            </form>
	            
	              
	            <form class="form-horizontal main_div" name="RenewSubscriptio" autocomplete="on" novalidate enctype="multipart/form-data" data-remote="true" id="formid" style="display: none;"> 
	            <div class="single_form_item"><br/></div>
	            
	   <!-- 1st Row Start -->
	             <div class="single_form_item"> 
	             	<!-- Renew Subscription -->                                                            
                     <div class="col-sm-3">
                     	<label for="inputFN" class="col-sm-12 control-label text_label" id="renewLabelId"><spring:message code="label.input.facilitator.renewSub"/></label>
					 </div>                   						
                     <div class="col-sm-3">
                     	<select class="form-control-general font_ipad" id="inputRenewSub" disabled>
             	        	<option class="form-control-general" value=""><spring:message code="label.input.facilitator.select.renewSub"/></option>
                     		<option class="form-control-general" value="RW"><spring:message code="label.select.facilitator.renew_drpDwn"/></option>
                     		<option class="form-control-general" value="AD"><spring:message code="label.select.facilitator.subscription_drpDwn"/></option>  
						</select>
                     </div>
                     
                     <!-- Payment mode -->         					 
					 <div class="col-sm-3">
                     <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="label.mode.payment"/></label>
					 </div>					 
                     <div class="col-sm-3">
                     	<select class="form-control-general font_ipad_two" onchange="tooglePaymentMode()" id="paymentMode">
	                    	<option class="form-control-general" value=""><spring:message code="label.select.mode.payment"/></option>
							<option class="form-control-general" value="CHK"><spring:message code="lbl.cheque.payment"/></option>
							<option class="form-control-general" value="CASH"><spring:message code="lbl.cheque.cash"/></option>    
							<option class="form-control-general" value="CMP"><spring:message code="lbl.cheque.complty"/></option>								
	                    </select>
                     </div> 
					 
                    <!-- <div class="col-sm-3">&nbsp;</div>  -->
                     <div class="clearfix"></div>
                </div>  
		<!-- 1st Row End -->	
        
		<!-- 2nd Row Start -->
				<div class="single_form_item">  
					<!-- Renew type: search user OR CSV upload -->                  						
				  	<div id="renewTypeRadio" >
                     <div class="col-sm-3">
					    <label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="menu.renew.select.option"/></label>
					 </div>
					  <div class="col-sm-3 sub_text_label">
                    	<div class="col-sm-5  radio_manual">
                            <input type="radio" name="preference" id="manual" value="M" onchange="toggleRenewSub('M')" checked = "checked"><spring:message code="label.searchIndi"/>
                       	</div>
                       	<div class="col-sm-5 radio_csv">
                       		<input type="radio" name="preference" id="csv" value="C" onchange="toggleRenewSub('C')"><spring:message code="label.csv"/>
                       	</div>
					 </div>
					</div>
					<div id="renewTypeRadioProxy" style="display: block" >
                     <div class="col-sm-6">
					  </div>					  
					</div>
					<div id="noOfUserFacilitatorDiv" style="display: none;">
						<div class="col-sm-3">
							<label for="inputUP" class="col-sm-12 control-label text_label">Number of Users:</label>
						</div>                   
						<div class="col-sm-3">						
							<input type="text" class="form-control-general" id="noOfUserFac" maxlength="3">
						<div class="clearfix"></div>	
						</div> 
					</div>
					<!-- CHECK amount -->
					<div id="chkAmountDiv" class="inputRightSide" style="display: none;">
						<div class="col-sm-3">
					    <label for="inputQtnNo" class="col-sm-12 control-label text_label"><spring:message code="label.chk.amount"/></label>
						</div>
                     <div class="col-sm-3">						
						<input type="text" class="form-control-general decimalPlace" value=""  id="totalAmount" maxlength = "7" data-container="body" data-toggle="popover" data-placement="top" data-content="Amount format should be 9999.99"/>
						<div class="clearfix"></div>	
                     </div> 
					</div> 
					<!-- CASH amount -->
					<DIV id="cashAmountDiv" class="inputRightSide" style="display: none;">
						<div class="col-sm-3">
					    <label for="inputQtnNo" class="col-sm-12 control-label text_label"><spring:message code="label.cash.amount"/></label>
						</div>
                     <div class="col-sm-3">						
						<input type="text" class="form-control-general decimalPlace" value=""  id="cashAmount" maxlength = "7" data-container="body" data-toggle="popover" data-placement="top" data-content="Amount format should be 9999.99"/>
						<div class="clearfix"></div>	
                     </div> 
					</DIV> 
					 
                     <!--<div class="col-sm-3">&nbsp;</div>  -->
                     <div class="clearfix"></div>
                </div>  
           <!-- 2nd Row End -->     
		   
		   <!-- 3rd Row Start -->
		    <div class="single_form_item">	
		    	<!-- Individual user's email -->				
                <DIV id="emalIndividualDiv">
					<div class="col-sm-3" >
		                <label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="label.emailId"/></label>
					</div>      	
		            <div class="col-sm-3">
		                 <input type="text" class="form-control-general" value="" id="userEmail">
						<div class="col-sm-3">&nbsp;</div>  
		                <div class="clearfix"></div>
		            </div>
	             </div>	 
	             <DIV id="emalIndividualDivProxy">
					<div class="col-sm-6" >		             
		            </div>
	             </div>
	               <!-- Facilitator cust ID -->
	             <div id="facilitatorIdDiv" style="display: none;">
                     <div class="col-sm-3">
                     <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="label.input.facilitator.id"/></label>
					 </div>
                     <div class="col-sm-3">
                     	<input type="text" class="form-control-general" id="facilitatorId"/>
                     </div>
				</div>
	             
		                 
               <!-- check no -->
               <DIV id="chkNo2" class="inputRightSide" style="display: none;">
					<div class="col-sm-3" >
	                   	<label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="label.number.cheque"/></label>
					</div>
	               	
	               	<div class="col-sm-3">
	                    	<input type="text" class="form-control-general" value="" id="chequeno" maxlength="15">
						<div class="col-sm-3">&nbsp;</div>  
	                   <div class="clearfix"></div>
	                   </div>
	             </div>				 
				<!-- CASH payment date -->
			  
	              <DIV id="paymentDtCashDiv" class="inputRightSide" style="display: none;">
						 <div class="col-sm-3">
	                     	<label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="label.date.payment"/></label>
	                     </div>
						 <div class="col-sm-3">
	                     	<input type="text" class="form-control-general datepickerPastDate" value="" id="paymentDtCash" readonly="readonly">
						</div> 
					</DIV> 
					<DIV id="freeProxyDiv">
						 <div class="col-sm-6">
	                     </div> 
					</DIV>          
	              
	             <div class="clearfix"></div>
                </div>  
		   
		   <!-- 3rd Row End -->
		   
		   <!-- 4th Row Start -->
				<div class="single_form_item">
		 
					 <div id="expiryDurationDiv">
	                     <div class="col-sm-3">
						    <label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="label.expDuration"/></label>
						 </div>                   
						 <div class="col-sm-3">						
						   <select class="form-control-general font_ipad" id="expiryDuration">
	             	        	<option class="form-control-general" value=""><spring:message code="label.choose.expire.date"/></option>
	                     		<option class="form-control-general" value="1"><spring:message code="label.choose.one.month"/></option>
	                     		<option class="form-control-general" value="2"><spring:message code="label.choose.two.month"/></option> 
	                     		<option class="form-control-general" value="3"><spring:message code="label.choose.three.month"/></option>
	                     		<option class="form-control-general" value="4"><spring:message code="label.choose.four.month"/></option>
	                     		<option class="form-control-general" value="5"><spring:message code="label.choose.five.month"/></option> 
	                     		<option class="form-control-general" value="6"><spring:message code="label.choose.six.month"/></option>
	                     		<option class="form-control-general" value="7"><spring:message code="label.choose.seven.month"/></option>
	                     		<option class="form-control-general" value="8"><spring:message code="label.choose.eight.month"/></option>
	                     		<option class="form-control-general" value="9"><spring:message code="label.choose.nine.month"/></option>
	                     		<option class="form-control-general" value="10"><spring:message code="label.choose.ten.month"/></option> 
	                     		<option class="form-control-general" value="11"><spring:message code="label.choose.eleven.month"/></option>
	                     		<option class="form-control-general" value="12"><spring:message code="label.choose.twelve.month"/></option>
						   </select>
							<div class="clearfix"></div>	
	                     </div> 
					</div> 				
							
				 
				<div id="noUserIndividual1" class="col-sm-6" style="display: none;">
				 </div> 
				 <!-- check payment date -->
					<DIV id="paymentDt2" class="inputRightSide" style="display: none;">
						 <div class="col-sm-3">
		                    	<label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="label.date.payment"/></label>
		                    </div>
						 <div class="col-sm-3">
		                    	<input type="text" class="form-control-general datepickerPastDate" value="" id="datepicker" readonly="readonly">
						</div> 
					</DIV>
                   <div class="clearfix"></div>
              </div>  
		   
		   <!-- 4th Row End -->
		 
			   
			<!-- 5th Row Start -->
				 <div class="single_form_item">
					<div id="csvFileArea1" style="display: none;">			
						<div class="single_form_item" style="display: none;" id="manualEntry">
                   		</div>
                   		
                   		<div id="fileUploadDiv">
                   			<div class="single_form_item">
	                     		<div class="col-sm-3 ">
	                     			<label for="csvFileLbl" class="col-sm-12 control-label text_label"><spring:message code="select.csv.file"/></label>
						 		</div>
			                    <div class="col-sm-5 new_width_csv">
			                    	<input type="file" class="filestyle" data-classButton="btn choose_file" name="filename" id="filename">
			                    </div> 
			                    <div class="col-sm-3"><a id='downloadCSVFmt' class='download_link_style download_ipad' onclick="downloadCSVFmt()" style="cursor: pointer;"><img src="../img/download.png" title="Download Sample CSV" /></a></div>  
	                   			<div class="clearfix"></div>
							</div>
                   		</div>
					</div>										
				</div> 
				
					
		   <!-- 5th Row End --> 
		   
		     <!-- 6th Row Start -->
		     <!--  <div class="single_form_item">
		      <div id="manualEntryArea1" style="display: none;">
		      		<div class="single_form_item" id="manualEntry">
                   	</div>
		      </div>
		      </div> -->		     
		      <!-- 6th Row End -->
		   
	             <div class="single_form_item">
                   <p>
					<input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Subscribe / Renew" id="addNewUserBtnId" onclick="subscribeOrRenewUser()" />
				   </p>
                 </div>	
                   <div class="single_form_item"><br/></div>
	            </form>
	      
	           
	           <div  class="existing_data_div">	             
	             <div id="existingDataListDiv" class= "existing_list" >			
					
				 </div>
				 <!-- existing user's details -->
	            <div id="existingUsersTableId" class="table_div_scroll" >
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
		
		
<!-- <script src="https://code.jquery.com/jquery.js"></script> -->    
<!-- Include all compiled plugins (below), or include individual files as needed -->
<script src="../js/bootstrap.min.js"></script>
<script src="../js/renew.js"></script>
<script src="../js/jquery.tablesorter.js"></script>
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
      			<button class="alertify-button alertify-button-ok" id="alertify-oks" onclick="renewCSVAfterValidation()">OK</button>
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