<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
<head>
    <title><spring:message code="menu.sub.create.surveyqtn.group" /></title>
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
	<script>
	if (typeof jQuery == 'undefined') {
    	document.write(unescape("%3Cscript src='../js/latest_10.2_jquery.js' type='text/javascript'%3E%3C/script%3E"));
    }
	</script>	
	<!--  <script type="text/javascript" src="../js/disableRC.js"></script>	 -->
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
	            <div class="col-sm-2 pd-0 col">
	            	<!-- Menu area start -->
	    			<jsp:include page="sideMenuNew.jsp"/>
	   				 <!-- Menu area end -->   
	            </div>
	            
	            <div class="col-sm-10">
	            <div class= "heading_label">
	            	<ol class="breadcrumbAdmin">
           				<li class="first_nav"><spring:message code="reports.breadcrumb.manage.user.home"/> &nbsp;|</li>
           				<li class="first_nav"><spring:message code="menu.sub.create.user"/> &nbsp;|</li>
           				<li class="second_nav" id='updateId'><spring:message code="menu.sub.create.surveyqtn.group"/></li>
             		</ol>
	            </div>
	            <div class= "sub_contain"><spring:message code="menu.sub.create.surveyqtn.group"/></div>
	            <div class= "main_contain">
	            <div class= "inner_container" id="addHeaderDiv"><spring:message code="menu.sub.add.surveyqtn"/></div>	  
	            <form class="form-horizontal main_div" name="crtRefQtn" autocomplete="on" novalidate enctype="multipart/form-data" data-remote="true" id="ugForm"> 
	            <div class="single_form_item"><br/></div>
	            
	            
	            <div class="single_form_item">
                     <div class="col-sm-4">
                     <label for="inputUP" class="col-sm-12 control-label text_label"><spring:message code="label.select.user.type"/></label></div>
                     <div class="col-sm-4">
                     	<select class="form-control-general selectSurvey" id="userType" onchange="fieldsDisabled()">
                     		<option class="form-control-general" value="0"><spring:message code="label.user"/> </option>
							<option class="form-control-general" value="1"><spring:message code="label.user.type.individual"/> </option>
							<option class="form-control-general" value="3"><spring:message code="label.user.type.facilitator"/> </option>
                     	</select>
                     </div> 
                     <div class="col-sm-3 hidden-ipad">&nbsp;</div>  
                     <div class="clearfix"></div>
                </div>  
                
                <!--  USER PROFILE -->
	            <div id="createUserGroupUpdate"> 
	            <div class="single_form_item">
                     <div class="col-sm-4 ">
                     <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="label.user.select.profile"/></label></div>
                     <div class="col-sm-4">
                     	<select class="form-control-general selectSurvey" id="inputUserProfileInpt" onchange="loadProfileData()" disabled>
                     		<option class="form-control-general" value="0"><spring:message code="label.user.select.profile_drpDwn"/> </option>
                     	</select>
                     </div> 
                     <div class="col-sm-3">&nbsp;</div>  
                     <div class="clearfix"></div>
                </div>  
                </div>
                
	            <div class="single_form_item">
                   <div class="col-sm-4">
                     <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="label.type.answer"/></label></div>
                     <div class="col-sm-8 sub_text_label select_option_radio sub_text_label_survey">
                       <div class="col-sm-2 radio_text radio_text_txt">
                          <input type="radio" name="radios" id="text" value="T" onclick="toogleSelectionDiv('T')" disabled><spring:message code="label.text"/>
                       </div>
                       <div class="col-sm-2 radio_text radio_text_radio">
                       	  <input type="radio" name="radios" id="rad" value="R" onclick="toogleSelectionDiv('R')" disabled> <spring:message code="label.radio"/>
                       </div>
					   <div class="col-sm-3 radio_text radio_text_dropdown">
                       	  <input type="radio" name="radios" id="drop" value="D" onclick="toogleSelectionDiv('D')" disabled> <spring:message code="label.dropdown"/>
                       </div>
					   <div class="col-sm-4 radio_text radio_text_checkbox">
                       	  <input type="radio" name="radios" id="check" value="C" onclick="toogleSelectionDiv('C')" disabled> <spring:message code="label.checkbox"/>
                       </div>
                     </div> 
                     <!--<div class="col-sm-2 hidden-ipad">&nbsp;</div>-->
                     <div class="clearfix"></div>
                </div>
                 <div id="radioqtn" style='display:none;'>
				  <div class="single_form_item">
                   <div class="col-sm-4 ">
                   <label for="inputFN" class="col-sm-12 control-label text_label"><spring:message code="label.question.main"/></label></div>
                   <div class="col-sm-4">
                   	<input type="text" class="form-control-general-form-redefined" style="width: 90%!important;" id="questions" maxlength="100" onkeypress='return validateDescription(event)'>
                   </div> 
                   <div class="col-md-2 mandatoryCheckBox">
						<label for='inputFN' class="text_label" style="margin-left: -13%;">Mandatory:</label>
						<input type="checkbox"  class="mandatoryChkBox" id="mandatoryChkbox" name="mandatory" value='Y'>
					</div>
                   <div class="col-sm-2 hidden-ipad">&nbsp;</div>
                     <div class="clearfix"></div>
                </div>
				 </div>
				 
				 <div id="textopt" style="display: none;">
				  <div class="single_form_item">
                     <div class="col-sm-4 ">
                     <label for="inputQtnNo" class="col-sm-12 control-label text_label"><spring:message code="label.number.question"/></label></div>
                     <div class="col-sm-4">
                     	<select class="form-control-general selectSurvey"  id="inputNoOFQues" onchange="populateFields()" disabled>
                     		<option class="form-control-general" value=""><spring:message code="label.select.number.question"/></option>   
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
                     		<option class="form-control-general" value="11"><spring:message code="label.choose.eleven"/></option>
                     		<option class="form-control-general" value="12"><spring:message code="label.choose.twelve"/></option>
                     		<option class="form-control-general" value="13"><spring:message code="label.choose.thirteen"/></option>
                     		<option class="form-control-general" value="14"><spring:message code="label.choose.fourteen"/></option>
                     		<option class="form-control-general" value="15"><spring:message code="label.choose.fifteen"/></option>
                     		<option class="form-control-general" value="16"><spring:message code="label.choose.sixteen"/></option>
                     		<option class="form-control-general" value="17"><spring:message code="label.choose.seventeen"/></option>
                     		<option class="form-control-general" value="18"><spring:message code="label.choose.eighteen"/></option>
                     		<option class="form-control-general" value="19"><spring:message code="label.choose.nineteen"/></option>
                     		<option class="form-control-general" value="20"><spring:message code="label.choose.twenty"/></option>
                     		<option class="form-control-general" value="21"><spring:message code="label.choose.twentyone"/></option>
                     		<option class="form-control-general" value="22"><spring:message code="label.choose.twentytwo"/></option>
                     		<option class="form-control-general" value="23"><spring:message code="label.choose.twentythree"/></option>
                     		<option class="form-control-general" value="24"><spring:message code="label.choose.twentyfour"/></option>
                     		<option class="form-control-general" value="25"><spring:message code="label.choose.twentyfive"/></option>
                     		<option class="form-control-general" value="26"><spring:message code="label.choose.twentysix"/></option>
                     		<option class="form-control-general" value="27"><spring:message code="label.choose.twentyseven"/></option>
                     		<option class="form-control-general" value="28"><spring:message code="label.choose.twentyeight"/></option>
                     		<option class="form-control-general" value="29"><spring:message code="label.choose.twentynine"/></option>
                     		<option class="form-control-general" value="30"><spring:message code="label.choose.thirty"/></option>
                     		<option class="form-control-general" value="31"><spring:message code="label.choose.thirtyone"/></option>
                     		<option class="form-control-general" value="32"><spring:message code="label.choose.thirtytwo"/></option>
                     		<option class="form-control-general" value="33"><spring:message code="label.choose.thirtythree"/></option>
                     		<option class="form-control-general" value="34"><spring:message code="label.choose.thirtyfour"/></option>
                     		<option class="form-control-general" value="35"><spring:message code="label.choose.thirtyfive"/></option>
                     		<option class="form-control-general" value="36"><spring:message code="label.choose.thirtysix"/></option>
                     		<option class="form-control-general" value="37"><spring:message code="label.choose.thirtyseven"/></option>
                     		<option class="form-control-general" value="38"><spring:message code="label.choose.thirtyeight"/></option>
                     		<option class="form-control-general" value="39"><spring:message code="label.choose.thirtynine"/></option>
                     		<option class="form-control-general" value="40"><spring:message code="label.choose.thirtyforty"/></option>
                     		<option class="form-control-general" value="41"><spring:message code="label.choose.fortyone"/></option>
                     		<option class="form-control-general" value="42"><spring:message code="label.choose.fortytwo"/></option>
                     		<option class="form-control-general" value="43"><spring:message code="label.choose.fortythree"/></option>
                     		<option class="form-control-general" value="44"><spring:message code="label.choose.fortyfour"/></option>
                     		<option class="form-control-general" value="45"><spring:message code="label.choose.fortyfive"/></option>
                     		<option class="form-control-general" value="46"><spring:message code="label.choose.fortysix"/></option>
                     		<option class="form-control-general" value="47"><spring:message code="label.choose.fortyseven"/></option>
                     		<option class="form-control-general" value="48"><spring:message code="label.choose.fortyeight"/></option>
                     		<option class="form-control-general" value="49"><spring:message code="label.choose.fortynine"/></option>
                     		<option class="form-control-general" value="50"><spring:message code="label.choose.fifty"/></option>                     		        		
                     	</select>
                     </div> 
                     <div class="col-sm-3">&nbsp;</div>  
                     <div class="clearfix"></div>
                   </div>
				   </div>
				   
				   <div class="single_form_item" style="display: none;" id="quesEntry">
                   </div>
                   
                   <div id='radiopt' style='display:none'>
				 <div class="single_form_item">
                     <div class="col-sm-4 ">
                     <label for="inputQtnNo" class="col-sm-12 control-label text_label"><spring:message code="label.number.options"/></label></div>
                     <div class="col-sm-4">
                     	<select class="form-control-general selectSurvey" id="inputNoOFUser" name="radioq" onchange="populateFieldsRadio();" disabled>
                     		<option class="form-control-general" id="select" value=""><spring:message code="label.select.options"/></option>   
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
                     		<option class="form-control-general" value="11"><spring:message code="label.choose.eleven"/></option>
                     		<option class="form-control-general" value="12"><spring:message code="label.choose.twelve"/></option>
                     		<option class="form-control-general" value="13"><spring:message code="label.choose.thirteen"/></option>
                     		<option class="form-control-general" value="14"><spring:message code="label.choose.fourteen"/></option>
                     		<option class="form-control-general" value="15"><spring:message code="label.choose.fifteen"/></option>
                     		<option class="form-control-general" value="16"><spring:message code="label.choose.sixteen"/></option>
                     		<option class="form-control-general" value="17"><spring:message code="label.choose.seventeen"/></option>
                     		<option class="form-control-general" value="18"><spring:message code="label.choose.eighteen"/></option>
                     		<option class="form-control-general" value="19"><spring:message code="label.choose.nineteen"/></option>
                     		<option class="form-control-general" value="20"><spring:message code="label.choose.twenty"/></option>
                     		<option class="form-control-general" value="21"><spring:message code="label.choose.twentyone"/></option>
                     		<option class="form-control-general" value="22"><spring:message code="label.choose.twentytwo"/></option>
                     		<option class="form-control-general" value="23"><spring:message code="label.choose.twentythree"/></option>
                     		<option class="form-control-general" value="24"><spring:message code="label.choose.twentyfour"/></option>
                     		<option class="form-control-general" value="25"><spring:message code="label.choose.twentyfive"/></option>
                     		<option class="form-control-general" value="26"><spring:message code="label.choose.twentysix"/></option>
                     		<option class="form-control-general" value="27"><spring:message code="label.choose.twentyseven"/></option>
                     		<option class="form-control-general" value="28"><spring:message code="label.choose.twentyeight"/></option>
                     		<option class="form-control-general" value="29"><spring:message code="label.choose.twentynine"/></option>
                     		<option class="form-control-general" value="30"><spring:message code="label.choose.thirty"/></option>
                     		<option class="form-control-general" value="31"><spring:message code="label.choose.thirtyone"/></option>
                     		<option class="form-control-general" value="32"><spring:message code="label.choose.thirtytwo"/></option>
                     		<option class="form-control-general" value="33"><spring:message code="label.choose.thirtythree"/></option>
                     		<option class="form-control-general" value="34"><spring:message code="label.choose.thirtyfour"/></option>
                     		<option class="form-control-general" value="35"><spring:message code="label.choose.thirtyfive"/></option>
                     		<option class="form-control-general" value="36"><spring:message code="label.choose.thirtysix"/></option>
                     		<option class="form-control-general" value="37"><spring:message code="label.choose.thirtyseven"/></option>
                     		<option class="form-control-general" value="38"><spring:message code="label.choose.thirtyeight"/></option>
                     		<option class="form-control-general" value="39"><spring:message code="label.choose.thirtynine"/></option>
                     		<option class="form-control-general" value="40"><spring:message code="label.choose.thirtyforty"/></option>
                     		<option class="form-control-general" value="41"><spring:message code="label.choose.fortyone"/></option>
                     		<option class="form-control-general" value="42"><spring:message code="label.choose.fortytwo"/></option>
                     		<option class="form-control-general" value="43"><spring:message code="label.choose.fortythree"/></option>
                     		<option class="form-control-general" value="44"><spring:message code="label.choose.fortyfour"/></option>
                     		<option class="form-control-general" value="45"><spring:message code="label.choose.fortyfive"/></option>
                     		<option class="form-control-general" value="46"><spring:message code="label.choose.fortysix"/></option>
                     		<option class="form-control-general" value="47"><spring:message code="label.choose.fortyseven"/></option>
                     		<option class="form-control-general" value="48"><spring:message code="label.choose.fortyeight"/></option>
                     		<option class="form-control-general" value="49"><spring:message code="label.choose.fortynine"/></option>
                     		<option class="form-control-general" value="50"><spring:message code="label.choose.fifty"/></option>
                     	</select>
                     </div> 
                     <div class="col-sm-3">&nbsp;</div>  
                     <div class="clearfix"></div>
                   </div>
				    <div class="single_form_item" style="display: none;" id="radioEntry">
                   </div>
				 </div>
				 <div class="single_form_item" style="display: none;" id="dropdwnEntry">
                   </div>
				 <div class="single_form_item" style="display: none;" id="checkboxEntry">
                   </div>
				 <div id="radiodiv" style='display:none;'>
				  <div class="single_form_item">
                   <div class="col-sm-4 ">
                   <label for="inputFN" class="col-sm-12 control-label text_label" id="labelqtn"><spring:message code="label.number.opt"/></label></div>
                   <div class="col-sm-2"><input type="text" class="form-control-general-form-redefined" value=""  id="options" onkeypress='return checkInput(event, this)'></div> 
                   <div class="col-sm-2 hidden-ipad">&nbsp;</div>
                     <div class="clearfix"></div>
                </div>
				</div>
				       
                                                               
                 <div class="single_form_item">
                   <p>
                   <input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn search_btn_survey" value="Save" id="saveConfig" onclick="saveSurveyQtns()" style="display: none; margin-left:44%"/>
					</p>
                 </div>	
                   
                  <div class="single_form_item" id="chgOrderDivSurvQtn">        
                   <div class="col-sm-5"></div>       
                   <div class="col-sm-4"><input type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn changeOrdBtn_Ipad" value="Change Order" id="chgOrderSurvQtnId" data-toggle="modal" data-target="#survQtnModal" onclick="changeOrderSurvQtn()"/></div>			
                 	<div class="col-sm-2"></div>
                 	 <div class="clearfix"></div>
                 </div>	
                  <div class="single_form_item"><br/></div>
	            </form>
	            
	        <!-- Modal -->
			<div class="modal fade" id="survQtnModal" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true">
			  <div class="modal-dialog">
			    <div class="modal-content">
			      <div class="modal-header">
			         <button type="button" class="custom_close_btn" data-dismiss="modal" aria-hidden="true"></button> 
			        
			        <div class="modal-header-desc" id="headerDescMainId">Change Survey Question Order</div>			       		         
			        <div class="pd-15"></div>
			        <!-- <div class="main-qtn-desc-area" id="mainQtnDescId"></div> -->
			        
			        <div id="mainAreaListId" class="main-qtn-area"></div>
			        <!-- <div id="subQtnListId" class="main-qtn-area"></div>  -->
			        

			      </div>
			      
			      <div class="modal-footer" id="attrSaveBtnId">
			        <!-- <button type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" data-dismiss="modal">Close</button> -->
			        <button type="button" class="btn_admin btn_admin-primary btn_admin-sm search_btn" onclick="saveSurveyQtnOrder()">Save</button>
			      </div>
			      			      		
			    </div>
			  </div>
			</div>
	            
	               <div class="existing_data_div">             
	             <div  id="existing_list_Id" class= "existing_list"><spring:message code="label.existing.survey.qtns"/></div>
	           		<div id="existingSurveyQtnsTableId" class="table_div_scroll">
	            		<div align="center">
	            			<img src="../img/dataLoading.GIF" />
	            			<p><spring:message code="msg.survey.qtn.loading"/></p>
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
		
		
		
		<script src="https://code.jquery.com/jquery.js"></script>    
<!-- Include all compiled plugins (below), or include individual files as needed -->
<script src="../js/bootstrap.min.js"></script>
<script src="../js/configureSurveyQtn.js"></script>
<script src="../js/jquery.tablesorter.js"></script>
<script src="../js/jquery.tablednd.js"></script>
<script type="text/javascript">


$(document).ready(function() {
    $("#table-1").tableDnD();
});


</script>
</body>
</html>