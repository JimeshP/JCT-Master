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
           				<li class="second_nav" id='updateId'><spring:message code="menu.sub.create.user"/></li>
             		</ol>
	            </div>
	            <div class= "sub_contain"><spring:message code="menu.sub.create.user"/></div>
	            <div class= "main_contain">
	            <div class= "inner_container" id=""><spring:message code="label.add.new.user"/></div>	  
	            <form class="form-horizontal main_div" name="addNewUser" autocomplete="on" novalidate enctype="multipart/form-data" data-remote="true"> 
	            <div class="single_form_item"><br/></div>
	
	            
	   <!-- 1st Row Start -->
	            <div class="single_form_item">
                     <div class="col-sm-3">
                     <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="label.select.user.type"/></label>
					 </div>
                     <div class="col-sm-3">
                   	
                     	<select class="form-control-general font_ipad" id="selectUserType" onchange="fieldsEnabled()" disabled>
	                   		<%-- <option class="form-control-general" value="0"><spring:message code="label.user.type.select.drpDwn"/></option>  
							<option class="form-control-general" value="1"><spring:message code="label.user.type.individual"/></option>      
							<option class="form-control-general" value="2"><spring:message code="label.user.type.facilitator"/></option> --%>       
	                   	</select>
                     </div> 
					 					
					 <!-- Expiry Duration Div (Individual User) -->
					 <div id="expiryDurationId" style="display: block;">
                    	<div class="col-sm-3" >
	                    	<label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="label.date.expiry"/></label>
	                	</div>
	                    <div class="col-sm-3">
                     	<select class="form-control-general font_ipad_two" id="expiryDuration">
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
                     </div>
                    </div>
                    
                    
                    <!-- Payment Mode Div (Facilitator) -->
                    <div id="paymentModeId" class="inputRightSide" style="display: none;">
                    <div class="col-sm-3">
                     <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="label.mode.payment"/></label>
					 </div>
                     <div class="col-sm-3">
                     	<select class="form-control-general font_ipad" onchange="tooglePaymentMode()" id="paymentMode1">
	                    	<option class="form-control-general" value=""><spring:message code="label.select.mode.payment"/></option>
							<option class="form-control-general" value="CHK"><spring:message code="lbl.cheque.payment"/></option>
							<option class="form-control-general" value="CASH"><spring:message code="lbl.cheque.cash"/></option>
							<option class="form-control-general" value="CMP"><spring:message code="lbl.cheque.complty"/></option>
	                    </select>
                     </div>
                    </div>
                    
                    <!-- <div class="col-sm-3">&nbsp;</div>  -->
                     <div class="clearfix"></div>
                </div>  
		<!-- 1st Row End -->	
            
        
		<!-- 2nd Row Start -->
			<div class="single_form_item">
				<div class="col-sm-3">
			    	<label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="label.user.select.group"/></label>
				</div>
                   <div class="col-sm-3">
                   	<select class="form-control-general font_ipad" id="inputUserGroupInpt" onChange="fetchExistingUsersByUserTypeAndGroup();">
                   		<option class="form-control-general" value=""><spring:message code="label.user.select.group_drpDwn"/></option>  
					</select>
                </div> 
                    
                 
                <!-- Payment Mode Div (Individual User) -->
                <div id="paymentModeIdIndividual" style="display: block;">
                  <div class="col-sm-3">
                   <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="label.mode.payment"/></label>
			 	</div>
                   <div class="col-sm-3">
                   	<select class="form-control-general font_ipad_two" onchange="tooglePaymentMode()" id="paymentModeInd">
                   	<option class="form-control-general" value=""><spring:message code="label.select.mode.payment"/></option>
					<option class="form-control-general" value="CHK"><spring:message code="lbl.cheque.payment"/></option>
					<option class="form-control-general" value="CASH"><spring:message code="lbl.cheque.cash"/></option>
					<option class="form-control-general" value="CMP"><spring:message code="lbl.cheque.complty"/></option>    	  	
                   </select>
                   </div>
                  </div>   
                  <!--  CHECK AMOUNT for FACILITATOR  - ARINDAM -->
                  <div id="chkAmountFaciDiv" class="inputRightSide" style="display:none">
                  <div class="col-sm-3">
                   <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="label.chk.amount"/></label>
			 	</div>
                   <div class="col-sm-3">
                   	<input type="text" class="form-control-general decimalPlace" maxlength="7" value="" id="chkAmountFaci" data-container="body" data-toggle="popover" data-placement="top" data-content="Amount format should be 9999.99"/>
                  </div> 
                  <div class="clearfix"></div>   
                     
                </div>	
                 <!--  CASH AMOUNT for FACILITATOR   -->
                  <div id="cashAmountFaciDiv" class="inputRightSide" style="display:none">
                  <div class="col-sm-3">
                   <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="label.cash.amount"/></label>
			 	</div>
                   <div class="col-sm-3">
                   	<input type="text" class="form-control-general decimalPlace" maxlength="7" value="" id="cashAmountFaci" data-container="body" data-toggle="popover" data-placement="top" data-content="Amount format should be 9999.99"/>
                  </div> 
                  <div class="clearfix"></div>   
                     
                </div> <div class="clearfix"></div>   
                </div>  
        <!-- 2nd Row End -->
           
		<!-- 3rd Row Start -->  
		    <div class="single_form_item">
		    	<div id="radioButtonDivIndividual" style="display: block;">
                     <div class="col-sm-3">
					    <label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="menu.sub.select.option"/></label>
					 </div>
					 <div class="col-sm-3 sub_text_label">
                    	<div class="col-sm-5  radio_manual">
                            <input type="radio" name="preference" id="manual" value="M" onclick="toogleSelectionDiv('M')" checked = "checked"><spring:message code="label.manual"/>
                       	</div>
                       	<div class="col-sm-5 radio_csv">
                       		<input type="radio" name="preference" id="csv" value="C" onclick="toogleSelectionDiv('C')"><spring:message code="label.csv"/>
                       	</div>
					</div>
				</div> 
				
				<!-- Check Amout Individual -->	
                <div id="chkAmmoutIndividual" class="inputRightSide" style="display: none;">
              		<div class="col-sm-3">
					   <label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="label.chk.amount"/></label>
						</div>
                     <div class="col-sm-3">
                     		<input type="text" class="form-control-general decimalPlace" value=""  id="totalAmount" maxlength="7" data-container="body" data-toggle="popover" data-placement="top" data-content="Amount format should be 9999.99"/>
					
                	<div class="clearfix"></div>
                   </div> 
                </div> 
                
                <!-- Cash Amout Individual -->	
                <div id="cashAmmoutIndividual" class="inputRightSide" style="display: none;">
              		<div class="col-sm-3">
					   <label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="label.cash.amount"/></label>
						</div>
                     <div class="col-sm-3">
                     		<input type="text" class="form-control-general decimalPlace"  maxlength="7" value=""  id="totalAmountIndi" data-container="body" data-toggle="popover" data-placement="top" data-content="Amount format should be 9999.99"/>
					
                	<div class="clearfix"></div>
                   </div> 
                </div> 
				
				<!-- Facilitator Name -->				
				<div id="facilitatorNameDiv" style="display: none;">
					<div class="col-sm-3 ">
                    	<label for="inputUP" class="col-sm-12 control-label text_label">
                    	<spring:message code="label.facilitator.email"/></label>
                    </div>
                    <div class="col-sm-3">
                    	<input type="text" class="form-control-general" value="" id="facEmail" />
                    </div>
                </div>  
                
                  <!-- Check Number Facilitator -->
             	<div id="chkNo1" class="inputRightSide" style="display: none;">
						<div class="col-sm-3" >
	                    	<label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="label.number.cheque"/></label>
	                	</div>
	                	<div class="col-sm-3">
	                     	<input type="text" class="form-control-general" value="" id="chequeNoFaci" maxlength="15">
	                    <div class="clearfix"></div>
	                    </div>
				</div>  
				
					<!-- Payment Date for Cash Individual -->
			 	<div id="paymentDtFaciCash" class="inputRightSide" style="display: none;">
					 <div class="col-sm-3">
                     	<label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="label.date.payment"/></label>
                     </div>
					 <div class="col-sm-3">
                     	<input type="text" class="form-control-general datepickerPastDate" value="" id="paymentDtFaciCashInpt" readonly="readonly">
					</div> 
				</div>            
                 
                               
               <div class="clearfix"></div>   
		    </div>
		  <!-- 3rd Row End -->		      
		   
		 <!-- 4th Row Start -->
				
		 <div class="single_form_item">
		 <div id="nosuser1" style="display: block;">
				<div class="col-sm-3">
			    <label for="inputQtnNo" class="col-sm-12 control-label text_label" id="noOfUserLabelId"><spring:message code="label.no.user"/></label>
				</div>
                   <div class="col-sm-3">
				<select class="form-control-general new_class_padding font_ipad" style="display: block;" id="inputNoOFUser" onchange="populateManualEntryFields()">
                   		<option class="form-control-general" value=""><spring:message code="label.select.noOfUser_drpDwn"/></option>   
                   		<option class="form-control-general" value="1"><spring:message code="label.choose.one"/></option>
                   		<option class="form-control-general" value="2"><spring:message code="label.choose.two"/></option>
                   		<option class="form-control-general" value="3"><spring:message code="label.choose.three"/></option>
                   		<option class="form-control-general" value="4"><spring:message code="label.choose.four"/></option>
                   		<option class="form-control-general" value="5"><spring:message code="label.choose.five"/></option>
                   		<option class="form-control-general" value="6"><spring:message code="label.choose.six"/></option>
                   		<option class="form-control-general" value="7"><spring:message code="label.choose.seven"/></option>
                   		<option class="form-control-general" value="8"><spring:message code="label.choose.eight"/></option>
                   		<option class="form-control-general" value="9"><spring:message code="label.choose.nine"/></option>
                   		<option class="form-control-general" value="10"><spring:message code="label.choose.ten"/></option>      
                   	</select>
			<!-- <input type="text" class="form-control-general" value="" style="display: none;" id="noOfUser1" maxlength="3"/> -->
			<div class="clearfix"></div>	
                   </div> 
			</div> 
		 
		   <!-- No of user (Facilitator) -->
		 	<div id="nosuser2" style="display: none;">
			<div class="col-sm-3 ">
                  	<label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="label.no.user"/></label>
                  </div>
                  <div class="col-sm-3">
                  	<input type="text" class="form-control-general" value=""  id="noOfUser2" maxlength="3"/>
			<!-- <div class="col-sm-3">&nbsp;</div>   -->
                  <div class="clearfix"></div>
                  </div>
             </div>
             
             
             <!-- Check Number Individual -->
             <div id="chkNo2" class="inputRightSide" style="display: none;">
						<div class="col-sm-3" >
	                    	<label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="label.number.cheque"/></label>
	                	</div>
	                	<div class="col-sm-3">	
	                     	<input type="text" class="form-control-general" value="" id="chequeno" maxlength="15">
	                    <div class="clearfix"></div>
	                    </div>
				</div>     
				
				<!-- Payment Date for Cash Individual -->
			 <div id="paymentDt1" class="inputRightSide" style="display: none;">
					 <div class="col-sm-3">
                     	<label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="label.date.payment"/></label>
                     </div>
					 <div class="col-sm-3">
                     	<input type="text" class="form-control-general datepickerPastDate" value="" id="datepickerCashIndiinput" readonly="readonly">
					</div> 
				</div>         
               <div class="clearfix"></div> 
               </div>  
               
            <!-- Facilitator First Name Row Started -->
			 	<div id="facilitatorFirstNameDiv" style="display: none;">
			 	<div class="single_form_item">    
					<div class="col-sm-3 ">
	                	 <label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="label.fac.first.name"/></label>
	             	</div>
	                <div class="col-sm-3">
	                  <input type="text" class="form-control-general" value=""  id="facFirstNameId" maxlength="30" onkeypress="return validateDescription(event)"/>	
	                </div>	                 	               
	                <div class="col-sm-3">&nbsp;</div>
	                 <div class="clearfix"></div>
	              </div>
	             </div>
		      <!-- Facilitator First Name Row Ended --> 
		      
		       <!-- Facilitator Last Name Row Started -->
			 	<div id="facilitatorLastNameDiv" style="display: none;">
			 	<div class="single_form_item">    					
	                <div class="col-sm-3">
	                	 <label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="label.fac.last.name"/></label>
	             	</div>
	                <div class="col-sm-3">
	                  <input type="text" class="form-control-general" value=""  id="facLastNameId" maxlength="30" onkeypress="return validateDescription(event)"/>	
	                </div>
	                <div class="col-sm-3">&nbsp;</div>
	                 <div class="clearfix"></div>
	              </div>
	             </div>
		      <!-- Facilitator Last Name Row Ended -->  
		      
           <div class="single_form_item">    
		    <div class="col-sm-3 ">&nbsp;</div>
		    <div class="col-sm-3 ">&nbsp;</div>
		    <!-- Payment Date for Check Individual -->
		    <div id="paymentDt2" class="inputRightSide" style="display: none;">
				<div class="col-sm-3">
	                <label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="label.date.payment"/></label>
	            </div>
				<div class="col-sm-3">
	                 <input type="text" class="form-control-general datepickerPastDate" id="datepicker" readonly="readonly">
				</div> 
			</div>
			<div class="clearfix"></div>
					
		    </div>  
		    <!-- 4th Row End --> 
		    
		    
		    		   
		   <!-- 5th Row Start -->
				 <div class="single_form_item">
					<div id="qorkingArea1" style="display: block;">			
						<div class="single_form_item" style="display: none;" id="manualEntry">
                   		</div>
                   		
                   		<div id="fileUploadDiv" style="display: none;">
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
				
					<div id="qorkingArea2" style="display: none;">
						<div id="facCheckDiv" class="facCheckDiv">
						<div class="single_form_item">                   	                   
                   			<input type="checkbox" class="" id="facCheck"> <span class="control-label text_label"><spring:message code="label.menu.faciAlsoUser"/></span>                   				
                   		</div>  
                   		<div class="clearfix"></div>
				  		<!-- </div> -->
						</div>
					</div>
					</div> 
		   <!-- 5th Row End -->            
		                                               
                 <div class="single_form_item">
                   <p>
					<input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" value="Create" id="addNewUserBtnId" onclick="saveUser()" />
				   </p>
                 </div>	
                   <div class="single_form_item"><br/></div>        
		 <!-- </div> -->
		   

                  
	            </form>
	            

	            
	            <div  class="existing_data_div">	             
	             <div id="existingDataListDiv" class= "existing_list" >
				 <table width='95%'>
				 	<tr>
				 		<td align="left">
				 			<spring:message code="label.existing.user.list"/>
				 		</td>
				 		<td align="right">
				 			<a href='bulkAction.jsp' class='activate_style'>Bulk Action</a>
				 		</td>
				 	</tr>
				 </table>
				 </div>
	            <div id="existingUsersTableId" class="table_div_scroll">
	           		<div align="center">
	            		<img src="../img/dataLoading.GIF" />
	            		<p> <spring:message code="label.existing.user"/></p>
	            	</div>
	            </div> 
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
		
		
<!-- <script src="https://code.jquery.com/jquery.js"></script> -->    
<!-- Include all compiled plugins (below), or include individual files as needed -->
<script src="../js/bootstrap.min.js"></script>
<script src="../js/bootstrap-filestyle.js"></script>
<script src="../js/common.js"></script>
<script src="../js/addUser.js"></script>
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