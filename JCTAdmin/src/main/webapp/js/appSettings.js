var imgByte = null;
var imgTypeIsPic = "yes";
var imageWidthHeightRslt = "yes";

$(document).ready(function(){

	$('#changeLogo').click(function(){
		clearPreviousData();
		var inputFile = document.getElementById('imageFile');
			inputFile.addEventListener('click', function() {this.value = null;}, false);
			inputFile.addEventListener('change', readData, false);
		
		function readData(evt) {
			evt.stopPropagation();
			evt.preventDefault();
			$('#fileField').val($('#imageFile').val());
			$('#imgCaption').show();
			var file = evt.dataTransfer !== undefined ? evt.dataTransfer.files[0] : evt.target.files[0];
			var reader = new FileReader();
			reader.onload = (function(theFile) {
				return function(e) {
					var image = new Image();
					image.src = e.target.result;
					image.onload = function() {
						var canvas = document.createElement('canvas');
						var img1 = $('#imageInput img')[0];
						var myImage = new Image();
						myImage.src = e.target.result;
						
						canvas.width = 300;
						canvas.height = image.height * (300 / image.width);
						if (myImage.height < 53 && myImage.width < 300) {
							canvas.width = myImage.width;
							canvas.height = myImage.height;
						} else if (myImage.height < 53 && myImage.width > 300) {
							canvas.width = 300;
							canvas.height = myImage.height;
						} else if (myImage.height > 53 && myImage.width > 300) {
							canvas.width = 300;
							canvas.height = image.height * (300 / image.width);
						} else if (myImage.height > 53 && myImage.width < 300) {
							canvas.width = myImage.width;
							canvas.height = myImage.height;
						}
						
						var ctx = canvas.getContext('2d');
						ctx.drawImage(image, 0, 0, canvas.width, canvas.height);
						$('#imageInput').html(['<img src="', canvas.toDataURL(), '"/>'].join(''));
						var img = $('#imageInput img')[0];
						var canvas = document.createElement('canvas');
						$('#imageInput img').Jcrop({					//		Crop selector initializing using Jcrop plugin
							bgColor: 'black',
							bgOpacity: .5,
							setSelect: [0, 0, 130, 130],
							aspectRatio: 0,
							onSelect: imgSelect,
							onChange: imgSelect
						});
						function imgSelect(selection) {
							/*if (selection.h < 53 && selection.w < 203) {
							canvas.width = selection.w + 30;
							canvas.height = selection.h + 30;
							}*/
							canvas.width = selection.w;
							canvas.height= selection.h;
							var ctx = canvas.getContext('2d');
							
							/******** modified for Mozilla firefox's problem ********/
							
							//ctx.drawImage(img, selection.x, selection.y, selection.w, selection.h, 0, 0, canvas.width, canvas.height);
							if(selection.w <= 0 && selection.h <= 0){
								ctx.drawImage(img, selection.x, selection.y, 1, 1, 0, 0, canvas.width, canvas.height);	//	1 & 1 has given bcoz firefox can't draw canvas with height=width=0px
							} else {
								ctx.drawImage(img, selection.x, selection.y, selection.w, selection.h, 0, 0, canvas.width, canvas.height);	//	original					
							}
							
							$('#imageOutput').attr('src', canvas.toDataURL());		// setting src for output/preview image	box
							//$('#image_output').css('height',selection.h,'width',selection.w);
							$('#imageOutput').css('height',selection.h);
							$('#imageOutput').css('width',selection.w +29);
							//var outputMaxWidth = $(".jcrop-tracker:last").width();
							//$('#imageOutput').css('max-width',$(".jcrop-tracker:last").width());
							$('#imageOutput').addClass("image_output");
						}	
					}
				}
			})(file);
			reader.readAsDataURL(file);
		}
});

/**
 * funcion is used when click on the Crop-Image/ #crop button on the pop-up
 */
	$('#crop').click(function(){
		if($('#imageFile').val() == "" && $('#fileField').val() == ""){							// When no image is chosen
			alertify.alert('Please select an image file');	
		}
		else {
			$('#filePath').val($('#imageOutput').attr('src'));						
			$('#crop').attr('data-dismiss',"modal");				// After cropping hide the modal pop-up box
			//var testwindow = window.open($('#imageOutput').attr('src'),"mywindow","menubar=1,resizable=0,width=350,height=250");	//Pop-up in new window
			//testwindow.moveTo(800,500);		
			imgByte = $('#imageOutput').attr('src');
		}
	});
});

/**
 * function added to clear all previous data when modal appears
 */
function clearPreviousData(){
	if( $("#imageInput").length > 0 ){
		document.getElementById("imageInput").innerHTML = "";	
		$("#imageFile").val("");
		$("#fileField").val("");
		$('#imgCaption').hide();
		var attr = $('#imageOutput').attr("src");
		
		// For some browsers, `attr` is undefined; for others `attr` is false.  Check for both.
		if ( (typeof attr !== typeof undefined) && (attr !== false) ) {
			$('#imageOutput').removeAttr("src").removeAttr("style").removeClass("image_output");
		}
	}
	
}

/**
 * funcion is used to show the pic once the user selects his / her user profile pic
 */
window.onload = function(){
	if (window.File && window.FileList && window.FileReader) {
        var filesInput = document.getElementById("headerLogoInptId");
        if(null != filesInput) {
        filesInput.addEventListener("change", function(event){
            
            var files = event.target.files; //FileList object
            for (var i = 0; i< files.length; i++) {
                var file = files[0];
                //Only pics
				
                if(!file.type.match('image')) {
                	alertify.alert('Please select an image file');
                	imgTypeIsPic = "no";
                	return false;
                }
				imgTypeIsPic = "yes";
                var picReader = new FileReader();
                picReader.addEventListener("load",function(event){
                    var picFile = event.target;
                    imgByte = picFile.result;
					
					var url = URL.createObjectURL(file);
					var img = new Image;
					img.src = url;
					img.onload = function() {
						var width = img.height;
						var height = img.height;
						if (width > 203 || height > 34) {
							alertify.alert("The size of your image is "+width+" / "+height+". \n Max width and height allowed is: 203px and 34 px");
							imageWidthHeightRslt = "no";
							return false;
						} else {
							imageWidthHeightRslt = "yes";
						}
					};
					
                });
                picReader.readAsDataURL(file);
            }
        });
    }
    }
    else {
        console.log("Your browser does not support File API");
    }
	fetchColor();
};

/**
 * funcion is used to save the changes in database
 * @param null
 */
function saveChanges() {
	var headerColor = document.getElementById('headerColorInptId').value;
	var footerColor = document.getElementById('footerColorInptId').value;
	var subHeaderColor = document.getElementById('subHeaderColorInptId').value;
	var instructionPanelColor = document.getElementById('instructionPanelColorInptId').value;
	var instructionPanelTextColor = document.getElementById('instructionPanelTextColorInptId').value;
	/**
	 * ADDED FOR PUBLIC VERSION
	 */
	if(headerColor == ""){
		alertify.alert("Please enter Header Color");			//		if any field left blank
		return false;
	} else if(footerColor == ""){
		alertify.alert("Please enter Footer Color");
		return false;
	} else if( subHeaderColor == ""){
		alertify.alert("Please enter Sub Header Color");
		return false;
	} else if( instructionPanelColor == "" ){
		alertify.alert("Please enter Instruction Panel Color");
		return false;
	} else if( instructionPanelTextColor == "" ){
		alertify.alert("Please enter Instruction Panel Text Color");
		return false;
	}
	/****** ENDED ******/
	if (headerColor == "") {
		headerColor = "NC"; 				//No Change
	}
	if (footerColor == "") {
		footerColor = "NC"; 				//No Change
	}
	if (subHeaderColor == "") {
		subHeaderColor = "NC"; 				//No Change
	}
	if (instructionPanelColor == "") {
		instructionPanelColor = "NC"; 		//No Change
	}
	if (instructionPanelTextColor == "") {
		instructionPanelTextColor = "NC"; 	//No Change
	}
	if (null == imgByte) {
		imgByte = "NC";
	}
	if (imgTypeIsPic == "no") {
		alertify.alert('Please select an image file');
    	return false;
	}
	if (imageWidthHeightRslt == "no") {
		alertify.alert("Max width and height allowed for logo is: 203px and 34 px");
		return false;
	}
	var appSettingsStub = Spine.Model.sub();
	appSettingsStub.configure("/admin/appSettings/saveAppSettings", "headerColor", "footerColor", "subHeaderColor", "instructionPanelColor", "instructionPanelTextColor", "imgByte", "jctEmail");
	appSettingsStub.extend( Spine.Model.Ajax );
	var modelPopulator = new appSettingsStub({  
		headerColor: headerColor,
		footerColor: footerColor,
		subHeaderColor: subHeaderColor,
		instructionPanelColor: instructionPanelColor,
		instructionPanelTextColor: instructionPanelTextColor,
		imgByte: imgByte,
		jctEmail: sessionStorage.getItem("jctEmail")
	});
	modelPopulator.save(); //POST
	appSettingsStub.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	appSettingsStub.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		//var statusCode = obj.statusCode; - THIS IS TAKEN FOR FUTURE USE
		alertify.alert(obj.message);
		//document.getElementById('headerColorInptId').value = "";
		//document.getElementById('footerColorInptId').value = "";
		//document.getElementById('subHeaderColorInptId').value = "";
		//document.getElementById('instructionPanelColorInptId').value = "";
		//document.getElementById('instructionPanelTextColorInptId').value = "";
	});
}
/**
 * Function alows to preview the document view
 */
function preview() {
	var headerColor = document.getElementById('headerColorInptId').value;
	var footerColor = document.getElementById('footerColorInptId').value;
	var subHeaderColor = document.getElementById('subHeaderColorInptId').value;
	var instructionPanelColor = document.getElementById('instructionPanelColorInptId').value;
	var instructionPanelTextColor = document.getElementById('instructionPanelTextColorInptId').value;
	if(headerColor == ""){
		alertify.alert("Please enter Header Color");			//		if any field left blank
		return false;
	} else if(footerColor == ""){
		alertify.alert("Please enter Footer Color");
		return false;
	} else if( subHeaderColor == ""){
		alertify.alert("Please enter Sub Header Color");
		return false;
	} else if( instructionPanelColor == "" ){
		alertify.alert("Please enter Instruction Panel Color Color");
		return false;
	} else if( instructionPanelTextColor == "" ){
		alertify.alert("Please enter Instruction Panel Text Color");
		return false;
	}
	
	if (headerColor == "") {
		headerColor = "#006990";
	}
	if (subHeaderColor == "") {
		subHeaderColor = "#88cbdf";
	}
	if (footerColor == "") {
		footerColor = "#006990";
	}
	if (instructionPanelColor == "") {
		instructionPanelColor = "#810123";
	}
	if (instructionPanelTextColor == "") {
		instructionPanelTextColor = "#fff";
	}
	var imageText = "<h5 class=\"heading_main logo\">&nbsp;</h5>";
	if (null != imgByte) {
		imageText = "<h5 class=\"heading_main\"><img src='"+imgByte+"' /></h5>";
		//imageText = "<h5 class=\"heading_main\"><img width='166px;' height='62px;' src='"+imgByte+"' /></h5>";
	}
	var myWindow = window.open("", "MsgWindow", "width=1000, height=600, resizable=no, status=no, location=no");
    myWindow.document.write("<html>");
    myWindow.document.write("<head>");
    myWindow.document.write("<title>Preview Page</title>");    
    myWindow.document.write("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">");
    myWindow.document.write("<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\"/>");
    
    myWindow.document.write("<link rel=\"stylesheet\" type=\"text/css\" href=\"../css/style.css\" />");
    myWindow.document.write("<link rel=\"stylesheet\" type=\"text/css\" href=\"../css/bootstrap.css\" />");
    //myWindow.document.write("<style type='text/css'>");
    //myWindow.document.write("</style>"); 
    myWindow.document.write("</head>");
    myWindow.document.write("<body>");
    myWindow.document.write("<div class=\"main_warp\">");
    
    /*header top area start*/
    myWindow.document.write("<div class=\"container-fluid header\" style='background: "+headerColor+"; border: none; border-bottom: 3px solid "+headerColor+";'>");
    myWindow.document.write("<div class=\"header_wrap_area container\">");
    myWindow.document.write("<div class=\"row\">");
    myWindow.document.write("<div class=\"col-xs-12 col-md-4 logo_area\">");
    myWindow.document.write(imageText);
    myWindow.document.write("</div>");
    myWindow.document.write("<div class=\"col-xs-12 col-md-4 tool_title\">JOB CRAFTING EXERCISE</div>");
    myWindow.document.write("<div class=\"col-xs-12 col-md-3 welcome_area\"><a id=\"fancybox-manual-my-account\">My Account</a> &nbsp;&nbsp;| &nbsp;&nbsp;<a id=\"logMeOut\">Logout</a>");
    myWindow.document.write("</div>");   
    myWindow.document.write("<div class=\"clearfix\"></div>");
    myWindow.document.write("</div></div></div>");      
    /*header top area End*/

    // SOF menu area 
    myWindow.document.write("<nav class='navbar navbar-default'>");
    myWindow.document.write("<div class='container-fluid nav-header' style='background: "+subHeaderColor+"; border: none; border-bottom: 3px solid "+subHeaderColor+";'>");
    myWindow.document.write("<div class='header_wrap_area container'>");
    myWindow.document.write("<div class='navbar-header'> ");
    myWindow.document.write("<a class='navbar-brand' href='#' data-toggle='collapse' data-target='#bs-example-navbar-collapse-1'>Menu Bar</a>");
    myWindow.document.write("</div>"); 
    myWindow.document.write("<div class='row navbar-collapse collapse' id='bs-example-navbar-collapse-1'>");
    myWindow.document.write("<div class='col-md-3 col-xs-12 pull-left time_section'>");
    myWindow.document.write("</div></div></div></div></nav>");
    // EOF menu area
    
    /* SOF Content & Instruction BAR Area     */
	myWindow.document.write("<div id='header-wrap' class='container'>");	          	 
	myWindow.document.write("<div class='row'>"); 	            
	myWindow.document.write("<div class='form_area' style='min-height:409px;'>");
	myWindow.document.write("<div class='form_area_top col-md-12 prev_nxt_sec'>");
	myWindow.document.write("<div class='clearfix'></div>");
	myWindow.document.write("</div>");
	/* Instruction area start */
	myWindow.document.write("<div class='instruction_area'>");
	myWindow.document.write("<div class='row_custom instruction'>");
	myWindow.document.write("<div id='panel' style='background: "+instructionPanelColor+"; padding: 25px 0px 0px 25px; display: block;>");
	myWindow.document.write("<ul class='ins-options'>");
	myWindow.document.write("<div id='instruction_data'><li style='color:"+instructionPanelTextColor+"'>Your instruction text  1.</li><li style='color:"+instructionPanelTextColor+"'>Your instruction text  2. </li><li style='color:"+instructionPanelTextColor+"'>Your instruction text n.</li></div></ul>");
	myWindow.document.write("</div>");
	myWindow.document.write("<p class='slide preview_slide' style='background:none;border-top:24px solid "+instructionPanelColor+"' ><a class='btn-slide' style='position:absolute;border-left: 20px solid transparent;border-right: 20px solid transparent;border-top: 22px solid "+instructionPanelColor+";width:0;height:0;padding:0;left:20px;bottom:-27px;' href='#' id='drp'>Slide Panel</a><span style='width:40px; height:5px;display:block;left:20px;position:absolute;background:"+instructionPanelColor+"'>&nbsp;</span><span class='arrow_down'></span></p>");
	myWindow.document.write("<h4 class='ins-title'>Instructions:</h4>");
    myWindow.document.write("<div class='clearfix'></div>"); 
    myWindow.document.write("</div>");
    myWindow.document.write("<div class='clearfix'></div>"); 
    myWindow.document.write("</div>");
    myWindow.document.write("<div class='clearfix'></div>"); 
    /* Instruction area end */
    myWindow.document.write("<script src='../js/instructionBar.js'></script>");
    myWindow.document.write("<div class='clearfix'></div>");    
    myWindow.document.write("<div class='clearfix'></div>"); 
    myWindow.document.write("<div class='form_area_middle' style='text-align:center;'>Page Content</div>");
    myWindow.document.write("<div class='form_area_bottom'></div>");
    myWindow.document.write("</div>");
    myWindow.document.write("</div>");
    myWindow.document.write("</div>");
    /* EOF Content & Instruction BAR Area     */
    
   
    // Footer area start
    myWindow.document.write("<div class='footer footer-wrapper' style='background: "+footerColor+";'>");
    myWindow.document.write("<div class='container'>");
    myWindow.document.write("<div class='copyright'>Copyright Text Goes here.</div>");
    myWindow.document.write("</div>");
    myWindow.document.write("</div>");
    // Footer area End
    
    myWindow.document.write("</div>");
	myWindow.document.write("</body>");
    myWindow.document.write("</html>");
}


function fetchColor() {
	var headerColor = "#006990";
	var footerColor = "#006990";
	var subHeaderColor = "#88cbdf";
	var instructionPanelColor = "#810123";
	var instructionPanelTextColor = "#fff";
	
	var appSettingsStub = Spine.Model.sub();
	appSettingsStub.configure("/admin/appSettings/fetchColor", "none");
	appSettingsStub.extend( Spine.Model.Ajax );
	var modelPopulator = new appSettingsStub({  
		none: ""
	});
	modelPopulator.save(); //POST
	appSettingsStub.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server. Populating default colors.");
		colorPopulator(headerColor, footerColor, subHeaderColor, instructionPanelColor, instructionPanelTextColor);
		populateTextBox(headerColor, footerColor, subHeaderColor, instructionPanelColor, instructionPanelTextColor);
	});
	appSettingsStub.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		colorPopulator(obj.headerColor, obj.footerColor, obj.subHeaderColor, obj.instructionPanelColor, obj.instructionPanelTextColor);
		populateTextBox(obj.headerColor, obj.footerColor, obj.subHeaderColor, obj.instructionPanelColor, obj.instructionPanelTextColor);
		imgByte = obj.image;
	});
}

function colorPopulator(headerColor, footerColor, subHeaderColor, instructionPanelColor, instructionPanelTextColor) {
	$("#headerColorInptId").spectrum( {
	    color: headerColor,
	    className: "full-spectrum",
	    showInitial: true,
	    showPalette: true,
	    showSelectionPalette: true,
	    maxPaletteSize: 10,
	    preferredFormat: "hex",
	} );
	$("#headerColorInptId").show();
	
	$("#footerColorInptId").spectrum({
	    color: footerColor,
	    className: "full-spectrum",
	    showInitial: true,
	    showPalette: true,
	    showSelectionPalette: true,
	    maxPaletteSize: 10,
	    preferredFormat: "hex",
	});
	$("#footerColorInptId").show();
	
	$("#subHeaderColorInptId").spectrum({
	    color: subHeaderColor,
	    className: "full-spectrum",
	    showInitial: true,
	    showPalette: true,
	    showSelectionPalette: true,
	    maxPaletteSize: 10,
	    preferredFormat: "hex",
	});
	$("#subHeaderColorInptId").show();
	
	$("#instructionPanelColorInptId").spectrum({
	    color: instructionPanelColor,
	    className: "full-spectrum",
	    showInitial: true,
	    showPalette: true,
	    showSelectionPalette: true,
	    maxPaletteSize: 10,
	    preferredFormat: "hex",
	});
	$("#instructionPanelColorInptId").show();
	
	$("#instructionPanelTextColorInptId").spectrum({
	    color: instructionPanelTextColor,
	    className: "full-spectrum",
	    showInitial: true,
	    showPalette: true,
	    showSelectionPalette: true,
	    maxPaletteSize: 10,
	    preferredFormat: "hex",
	});
	$("#instructionPanelTextColorInptId").show();
}

function populateTextBox(headerColor, footerColor, subHeaderColor, instructionPanelColor, instructionPanelTextColor) {
	document.getElementById('headerColorInptId').value = headerColor;
	document.getElementById('footerColorInptId').value = footerColor;
	document.getElementById('subHeaderColorInptId').value = subHeaderColor;
	document.getElementById('instructionPanelColorInptId').value = instructionPanelColor;
	document.getElementById('instructionPanelTextColorInptId').value = instructionPanelTextColor;
}