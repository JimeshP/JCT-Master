<!DOCTYPE html>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<html>

<body> 	
	<!-- Instruction area start -->
					    <div class="instruction_area">
			    			<div class="row_custom instruction">
		       				<div id="panel" style="display: none;">
					          	<ul class="ins-options">
					          		<div id="instruction_data"> </div>
						        </ul>
					        </div>
				        		<p class="slide">
				        		<a class="btn-slide"  href="#" id="drp">
				        		<span class="bar_loop">&nbsp;</span>
				        		<span class="arrow_down"></span>Slide Panel
				        		</a>
				        		
				        		</p>
				       			<h4 class="ins-title"><spring:message code="label.instructions"/></h4>
				       			<h4 id="watchVideoId" class="ins-title-video">
				       				
				       			</h4>
				       			<span id="watchVideoId2" style="height: 40%;"/>
				      		<div class="clearfix"></div> 
		      			</div>
		      			<div class="clearfix"></div> 
		      			</div>
		      			<div class="clearfix"></div> 
        <!-- Instruction area end -->
        
<script src="../js/instructionBar.js"></script>
</body>
</html>