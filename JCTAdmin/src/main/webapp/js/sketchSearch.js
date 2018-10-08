var firstRslt;
var maxRslt;
var imageText;
var nextLine;
/**
 * function is used search the before sketch diagram
 * for the specific job title.
 */
function searchBeforeSketch(){
	document.getElementById('headerDiv').innerHTML="";
	document.getElementById('imageDiv').style.display="block";
	document.getElementById('paginator').style.display="block";
	document.getElementById('imageEnlargeDiv').style.display="none";
	document.getElementById('imageDiv').innerHTML = "";
	document.getElementById('imageBS').innerHTML = "<img src='../img/Processing.gif' />";
	imageText = "<br />";
	nextLine = 0;
	$("#fancybox-manual-b").click(function() {
		$.fancybox.open({
			href : '#sketchSearchDiv',
			closeClick : false,
			openEffect : 'elastic',
			openSpeed  : 150,
			closeEffect : 'elastic',
			closeSpeed  : 150,
			padding : 5
		});
	});
	if(sessionStorage.getItem("firstResult")){
		sessionStorage.removeItem("firstResult");
	}
	if(sessionStorage.getItem("maxResults")){
		sessionStorage.removeItem("maxResults");
	}
	searchBSResult();
}
/**
 * function is used search the before sketch diagram
 * for the specific job title.
 */
function searchAfterSketch(){
	document.getElementById('ASheaderDiv').innerHTML="";
	document.getElementById('ASimageDiv').style.display="block";
	document.getElementById('ASpaginator').style.display="block";
	document.getElementById('ASimageEnlargeDiv').style.display="none";
	document.getElementById('ASimageDiv').innerHTML = "";
	document.getElementById('imageAS').innerHTML = "<img src='../img/Processing.gif' />";
	imageText = "<br />";
	nextLine = 0;
	$("#fancybox-manual-a").click(function() {
		$.fancybox.open({
			href : '#ASsketchSearchDiv',
			closeClick : false,
			openEffect : 'elastic',
			openSpeed  : 150,
			closeEffect : 'elastic',
			closeSpeed  : 150,
			padding : 5
		});
	});
	if(sessionStorage.getItem("firstResult")){
		sessionStorage.removeItem("firstResult");
	}
	if(sessionStorage.getItem("maxResults")){
		sessionStorage.removeItem("maxResults");
	}
	searchASResult();
}
/**
 * This function actually searches for the Before 
 * sketch Diagram.
 */
function searchBSResult(){
	//Create a model
	var beforeSketch = Spine.Model.sub();
	beforeSketch.configure("/user/search/searchBeforeSketch", "jobTitle", "firstResult", "maxResults");
	beforeSketch.extend( Spine.Model.Ajax );
	if(null == sessionStorage.getItem("firstResult")){
		firstRslt = 0;
	}else{
		firstRslt = parseInt(sessionStorage.getItem("maxResults")) + 1;
	}
	if(null == sessionStorage.getItem("maxResults")){
		maxRslt = 20;
	}else{
		maxRslt = parseInt(sessionStorage.getItem("maxResults")) + 20;
	}
	sessionStorage.setItem("firstResult", firstRslt);
	sessionStorage.setItem("maxResults", maxRslt);
	//Populate the model
	var populateModel = new beforeSketch({  
		jobTitle: sessionStorage.getItem("jobTitle"),
		firstResult: firstRslt,
		maxResults: maxRslt
	});
	populateModel.save(); //POST
	
	beforeSketch.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	
	beforeSketch.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		//console.log(obj);
		if(obj.statusCode == 200){
			var baseObj = obj.sketchStringList;
			var size = baseObj.length;
				if(size > 0){
				for(var index=0; index<size; index++){
					nextLine = nextLine + 1;
					imageText = imageText + "<img src="+baseObj[index]+" class='thumbnails' onclick='openImage("+baseObj[index]+")'/> &nbsp;&nbsp;&nbsp;&nbsp;";
					if(nextLine%3 == 0){
						imageText = imageText + "<br /><br /><br />";
					}
				}
				document.getElementById('headerDiv').innerHTML = "<br />Search Result For Level: "+sessionStorage.getItem("jobTitle");
				document.getElementById('imageBS').innerHTML = "";
				document.getElementById('imageDiv').innerHTML = imageText;
				//document.getElementById('contoller').innerHTML = "<< Previous &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Next >>";
			} else {
				document.getElementById('headerDiv').innerHTML = "<br />Search Result For Level: "+sessionStorage.getItem("jobTitle");
				//document.getElementById('imageBS').innerHTML = "";
				document.getElementById('imageBS').innerHTML  = "No result to display.";
			}
		}else{
			alert('handle the exception');
		}
	});
}

function searchASResult(){
	//Create a model
	var afterSketch = Spine.Model.sub();
	afterSketch.configure("/user/search/searchAfterSketch", "jobTitle", "firstResult", "maxResults");
	afterSketch.extend( Spine.Model.Ajax );
	if(null == sessionStorage.getItem("firstResult")){
		firstRslt = 0;
	}else{
		firstRslt = parseInt(sessionStorage.getItem("maxResults")) + 1;
	}
	if(null == sessionStorage.getItem("maxResults")){
		maxRslt = 20;
	}else{
		maxRslt = parseInt(sessionStorage.getItem("maxResults")) + 20;
	}
	sessionStorage.setItem("firstResult", firstRslt);
	sessionStorage.setItem("maxResults", maxRslt);
	//Populate the model
	var populateModel = new afterSketch({  
		jobTitle: sessionStorage.getItem("jobTitle"),
		firstResult: firstRslt,
		maxResults: maxRslt
	});
	populateModel.save(); //POST
	
	afterSketch.bind("ajaxError", function(record, xhr, settings, error) {
		alertify.alert("Unable to connect to the server.");
	});
	
	afterSketch.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		//console.log(obj);
		if(obj.statusCode == 200){
			var baseObj = obj.sketchStringList;
			var size = baseObj.length;
			if(size > 0){
				for(var index=0; index<size; index++){
					nextLine = nextLine + 1;
					imageText = imageText + "<img src="+baseObj[index]+" class='thumbnails' onclick='openImageAS("+baseObj[index]+")'/> &nbsp;&nbsp;&nbsp;&nbsp;";
					if(nextLine%3 == 0){
						imageText = imageText + "<br /><br /><br />";
					}
				}
				document.getElementById('ASheaderDiv').innerHTML = "<br />Search Result For Level: "+sessionStorage.getItem("jobTitle");
				document.getElementById('imageAS').innerHTML = "";//<img src="../img/Processing.gif" />
				document.getElementById('ASimageDiv').innerHTML = imageText;
				//document.getElementById('AScontoller').innerHTML = "<< Previous &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Next >>";
			} else {
				document.getElementById('ASheaderDiv').innerHTML = "<br />Search Result For Level: "+sessionStorage.getItem("jobTitle");
				//document.getElementById('imageAS').innerHTML = "";
				document.getElementById('imageAS').innerHTML  = "No result to display.";
			}
		}else{
			alert('handle the exception');
		}
	});
}

function openImage(img){
	document.getElementById('imageDiv').style.display="none";
	document.getElementById('paginator').style.display="none";
	document.getElementById('imageEnlargeDiv').style.display="block";
	document.getElementById('backButton').innerHTML = "<input type='button' class='popupBackButton' onClick='goBack()' value='Back'>";
	document.getElementById('mainPic').innerHTML = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br /><img src="+img+" class='thumbnailPopup'/>";
	//alert(img);
}

function openImageAS(img){
	document.getElementById('ASimageDiv').style.display="none";
	document.getElementById('ASpaginator').style.display="none";
	document.getElementById('ASimageEnlargeDiv').style.display="block";
	document.getElementById('ASbackButton').innerHTML = "<input type='button' class='popupBackButton' onClick='goBackAS()' value='Back'>";
	document.getElementById('ASmainPic').innerHTML = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br /><img src="+img+" class='thumbnailPopup'/>";
	//alert(img);
}
function goBack(){
	document.getElementById('imageDiv').style.display="block";
	document.getElementById('paginator').style.display="block";
	document.getElementById('imageEnlargeDiv').style.display="none";
}
function goBackAS(){
	document.getElementById('ASimageDiv').style.display="block";
	document.getElementById('ASpaginator').style.display="block";
	document.getElementById('ASimageEnlargeDiv').style.display="none";
}