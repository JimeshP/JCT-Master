<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>

<div class="row-fluid footer">
          <p class="pd-0">
          	<spring:message code="label.footer"/><br>
          	<spring:message code="label.footer.permission"/>	
          	<span id="tcHolder"  style="float: right; padding-right: 2%;cursor: pointer; display: none;" onclick="fetchTnstructions()">
          		<spring:message code="label.jct.tool.public.version"/>
          	</span>
          </p>
</div>
   <!-- Terms & Conditions modal box -->
	<div class="modal fade" id="tnCModal" tabindex="-1" data-backdrop="static">
	  <div class="modal-dialog" >
	    <div class="modal-content">
	      <div class="tcModal_header modal-header">
	        <button type="button" class="custom_close_btn" data-dismiss="modal"></button>      
	        <h5 class="modal-title tcModal_title" id="myModalLabel">Terms and Conditions</h5>
	      </div>
	      <br>
		   <div class="instruction_panel tcText_style" id="terms_condtn_div_id"></div>
		   <hr>
	      <div class="modal-footer" style="padding: 5px 20px 8px;">
	      		<input type="submit" class="pop-up-btn tcClose-btn" value="Close" data-dismiss="modal">
	      </div>
	    </div>
	  </div>	  
	</div>