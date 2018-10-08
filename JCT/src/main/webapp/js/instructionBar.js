/**
 * Function add to populate the instruction description
 * while the page is loaded
 * @param null
 */
function populateInstruction() {
	var relatedPage= "";
	var profileId = sessionStorage.getItem("profileId");
	var pageSeq = sessionStorage.getItem("pageSequence");
	if(pageSeq == 1){
		if (sessionStorage.getItem("bsEdit")) {
			relatedPage = "Edit Before Sketch";
		} else {
			relatedPage = "Creating Before Sketch";
		}
	} else if(pageSeq == 3){
		relatedPage = "Mapping Yourself";
		} else if(pageSeq == 4){
		if (sessionStorage.getItem("bsEdit")) {
			relatedPage = "Edit After Diagram";
		} else {
			relatedPage = "Creating After Diagram";
		}
	}
	var instruction = Spine.Model.sub();
	instruction.configure("/user/navigation/populateInstruction", "profileId", "relatedPage");
	instruction.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new instruction({  
		profileId: profileId,
		relatedPage: relatedPage
	});
	modelPopulator.save(); //POST
	instruction.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	instruction.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		if(statusCode == 200) {
			populateInstructionData(obj.instructionDesc, obj.videoUrl);
		} else if(statusCode == 201) {		
			populateInstructionData(obj.instructionDesc, obj.videoUrl);
		}
	});
}

/**
 * Function add set the instruction description 
 * while page loaded
 * @param instructionDesc
 */
function populateInstructionData(instructionDesc, videoUrl){
	if (videoUrl != null) {
		document.getElementById("watchVideoId").style.display = "block";
		// Changes by Rajan to match different URL context
		//adding /user url  - HARDCODING
		videoUrl = "https://s3-us-west-2.amazonaws.com/jobcrafting/" + videoUrl;
		// Changes by Rajan end
		sessionStorage.setItem("videoLink", videoUrl);
		document.getElementById("watchVideoId").innerHTML = "<a href='#' class='instruction_link' onclick='showVideo()'> Watch Video </a>";
	}
	document.getElementById("instruction_data").innerHTML = instructionDesc;
}

function showVideo() {
	var videoLink = sessionStorage.getItem("videoLink");
	//alert('My Video: '+videoLink);
	$.fancybox.open({
		href : '#watchVideoId2',
		//type : 'iframe',
		closeClick : false,
		openEffect : 'elastic',
		openSpeed  : 150,
		closeEffect : 'elastic',
		closeSpeed  : 150,
		padding : 5,
		closeClick : true,
		modal: true,
		showCloseButton : true,
		afterShow : function() {
			$('.fancybox-skin').addClass('ins_video_box');
			$('.fancybox-skin').append('<a title="Close" class="fancybox-item fancybox-close" href="javascript:jQuery.fancybox.close();"></a>');
		}
	});
	var pageSeq = sessionStorage.getItem("pageSequence");
	var desc = "";
	if(pageSeq == 1){
		desc = "Before Sketch<br>";
	} else if (pageSeq == 3) {
		desc = "Mapping Yourself<br>";
	} else if (pageSeq == 4) {
		desc = "After Diagram<br>";
	}
	var html = "<div class='video-section' style='align: center; width:90%;' ><div class='ins-page-title'>"+desc+"<video class='ins-video-section' width='100%' poster='../img/frame.jpg' controls preload='auto'>";
	html = html + "<source src='"+videoLink+"' type='video/mp4'></video></div></div>";   
	document.getElementById("watchVideoId2").innerHTML = html;
}

/**
Change 20.01.14*/
function populatePopUp(){
	
	var relatedPage= "";
	var profileId = sessionStorage.getItem("profileId");
	var pageSeq = sessionStorage.getItem("pageSequence");
	if(pageSeq == "3"){
		relatedPage = "BS";
	}
	else if(pageSeq == "4"){
		relatedPage = "AS";
	}
	var instruction = Spine.Model.sub();
	instruction.configure("/user/navigation/populatePopUp", "profileId", "relatedPage");
	instruction.extend( Spine.Model.Ajax );
	//Populate the model with data to transfer
	var modelPopulator = new instruction({  
		profileId: profileId,
		relatedPage: relatedPage
	});
	modelPopulator.save(); //POST
	instruction.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	instruction.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		var statusCode = obj.statusCode;
		var type = obj.popupType;
		
		if(statusCode == 200) {
			$('#myModal').modal('show');						
			if((null != obj.instructionTextBeforeVideo) || (obj.instructionTextBeforeVideo != "")) {
				document.getElementById("text_ins_before_video_id").innerHTML = obj.instructionTextBeforeVideo;				
			}			
			if(obj.instructionVideo) {
				var vidString = "https://s3-us-west-2.amazonaws.com/jobcrafting/" + obj.instructionVideo;
				var extension = vidString.substring(vidString.length , vidString.length-3);
				var instVid = document.getElementById("video_instruction_id");
				var htmlString = "<video width='85%' controls='' preload='auto'>"
				htmlString = htmlString + "<source src="+vidString+" type='video/"+extension+"'>";
				htmlString = htmlString + "</video>";
				instVid.innerHTML = htmlString;
				
				/*var iframe = document.createElement('iframe');
				iframe.src = vidString;
				var instVid = document.getElementById("video_instruction_id");
				iframe.width = '460';
				iframe.height = '200'; 
				instVid.appendChild(iframe);*/

				$( ".custom_close_btn" ).click(function() {
					iframe.src = "";
					});
			} else {
				document.getElementById('video_instruction_id').setAttribute("style", "display:none");
			}			
			if((null != obj.instructionTextAfterVideo) || (obj.instructionTextAfterVideo != "")) {
				document.getElementById("text_ins_after_video_id").innerHTML = obj.instructionTextAfterVideo;				
			}			
		}
		else if(statusCode == 318){
			$('#myModal').modal('hide');
		}	
	});
}