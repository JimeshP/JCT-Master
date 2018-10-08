<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<!DOCTYPE html>
<html>
	<body>
	    	<!-- FUNCTION GROUP -->
                   <div class="single_form_item">
                     <div class="col-sm-4 ">
                     <label for="inputFN" class="col-sm-12 control-label align-right"><spring:message code="label.functionGrp"/></label></div>
                     <div class="col-sm-5">
                     	<select class="form-control-general" id="inputFunctionGroup">
                     		<option class="form-control-general" value=""><spring:message code="label.select.functionGrp"/> </option>
                     	</select>
                     </div> 
                     <div class="col-sm-3">&nbsp;</div>  
                     <div class="clearfix"></div>
                   </div>
                 
                 <!-- JOB LEVEL -->
                   <div class="single_form_item">
                     <div class="col-sm-4 ">
                     <label for="inputFN" class="col-sm-12 control-label align-right"><spring:message code="label.jobLevel"/></label></div>
                     <div class="col-sm-5">
                     	<select class="form-control-general" id="inputJobLevel">
                     		<option class="form-control-general" value=""><spring:message code="label.select.jobLevel"/> </option>
                     	</select>
                     </div> 
                     <div class="col-sm-3">&nbsp;</div>  
                     <div class="clearfix"></div>
                   </div>
                   <div class="single_form_item">
                     <p><input type="button" class="btn btn-primary btn-sm" value="Search" id="searchBtnBS" onclick="populateBeforeSketchSelectedReport('newid')" /></p>
                   </div>	
                   <div class="single_form_item"><br/></div>           
</body>
</html>