var windowSize = $(window).width(); 
/**
 * Function call while refresh the page
 */
window.onunload = function(event) {
	return saveAfterSketchWhileRefresh();
};

window.setInterval(function(){
	updateTime();
}, 60000);

/**
 * Function add to populate the instruction description
 * while the page is loaded
 * @param null
 */ 
$(document).ready(function() {
	
	if (sessionStorage.getItem("asTimePage")) {
		sessionStorage.removeItem("asTimePage");
	}
	sessionStorage.setItem("asTimePage", "00:00:00");
	updateTimeAs();
	var insDisplay = sessionStorage.getItem("as2View");
	if(null == insDisplay || insDisplay == "" || insDisplay == "D"){
		
		populatePopUp();
	} else {
		$('#myModal').modal('hide');
	}
	populateInstruction();
	//ovalcontentEditable();	
	silentSave();
	toogleFocus();
 });

function updateTimeAs() {
	if (!sessionStorage.getItem("asTimePage")) {
		sessionTime = "00:00:00";
	} else {
		sessionTime = sessionStorage.getItem("asTimePage");
	}
	var totalTimeArr = sessionTime.split(":");
	var sec3 = parseInt(totalTimeArr[2]);
	sec3 = parseInt(sec3) + 1; 
	min3 = parseInt(totalTimeArr[1]);
	hour3 = parseInt(totalTimeArr[0]);
	if (parseInt(sec3) == 60) {
		sec3 = 0;
		min3 = parseInt(min3) + 1;
	} else {
		min3 = min3; 
	} if (parseInt(min3) == 60) {
		   min3 = 0; 
		   hour3 = parseInt(hour3) + 1; 
	}
	if (sec3<=9) { sec3 = "0" + sec3; }
	hour3 = ""+hour3;
	min3 = ""+min3;
	//console.log(hour3.length);
	if (""+hour3.length < 2) {
		hour3 = "0"+hour3;
	}
	if (""+min3.length < 2) {
		min3 = "0"+min3;
	}
	var totalTimeDisp = hour3+" Hrs "+min3+" Mins ";
	var totalTime3 = hour3+" : "+min3+" : "+sec3;
	sessionStorage.setItem("asTimePage", totalTime3);
	window.setTimeout("updateTimeAs();", 1000);
}

/**
 * Function add to disable browser back button
 * while the page is loaded
 * @param null
 */
window.location.hash="";
window.location.hash="Again-No-back-button";//again because google chrome don't insert first hash into history
window.onhashchange=function(){window.location.hash="";};

var clonedDivCount=0;
var trackerObjClone = null;
$('#pageWrapCopy').resizable();
/********** Added for instruction toggle ***********/
    $(".btn-slide").click(function(){
    $("#panel").slideToggle("slow");
    $(this).toggleClass("active"); 
    return false;
  });
    
/************** Element will be drag with in "#outderDiv" element ****************/
$(document).ready(function() {
	if (!sessionStorage.getItem("bsEdit")) {
	$('.after_sketch_task').draggable({
		cursor: 'move',       
		containment: '#outderDiv'    
	});  
	$('.oval_cloned_div').draggable({
	    cursor: 'move',       
	    containment: '#outderDiv'    
	 });
	$('.commonDrag').draggable({
	    cursor: 'move',       
	    containment: '#outderDiv'    
	 });
	}
});
	/************** To draw the Task Element of Before Sketch Page ****************/
var nextPageAfterSketch = 0; 
if(ClientSession.get("Next_page_after_sketch")!= null){
	var next_page_sketch = ClientSession.get("Next_page_after_sketch");
} else {
	var next_page_sketch = 0;
}
if(ClientSession.get("total_json_add_task")!=null && ClientSession.get("total_json_after_sketch_final") == null){
	var insDisplay = sessionStorage.getItem("as2View");
	if(null == insDisplay || insDisplay == "" || insDisplay == "D"){
		document.getElementById('panel').style.display = "block";
		document.getElementById('drp').setAttribute("class", "btn-slide active");
	} else {
		document.getElementById('panel').style.display = "none";
		document.getElementById('drp').setAttribute("class", "btn-slide");
	}
var jsonObj = JSON.parse(ClientSession.get("total_json_add_task"));
var count1 = jsonObj.length;
for (var i = 0; i<count1;i++){	
	var cnt = i+1;
	var divIda = jsonObj[i]["divId"];
	var divIdx = jsonObj[i]["divIdx"];
	var energyIda = jsonObj[i]["energyId"];
	var divStyle = jsonObj[i]["style"];
	if((divStyle == null) || (divStyle == "null") || (divStyle.search('block') != -1)){
		divStyle = "block;";
	}else{
		divStyle = "none;";
	}
	var energyValue = jsonObj[i]["energyValue"];
	var energyStyle = parseInt(energyValue)+5;
	var areaIda = jsonObj[i]["areaId"];
	var areaValue = jsonObj[i]["areaValue"];
	areaValue = areaValue.replace(/`/g, ",");
	areaValue = areaValue.replace(/-_-/g, "#");
	areaValue = areaValue.replace(/;_;/g, "?");
	var positionLefta = jsonObj[i]["positionLeft"];
	var positionTopa = jsonObj[i]["positionTop"];
	var positionWidth = jsonObj[i]["positionWidth"];
	var positionHeight = jsonObj[i]["positionHeight"];
	var additionalId = "additionalId_" + cnt;		
	
	var aspectRatio = jsonObj[i]["aspectRatio"];
	var deviceWidth = $("#outderDiv").width();
	var deviceHeight = $("#outderDiv").height();
	var positionLeftPercent = jsonObj[i]["positionLeftPercent"];
	var positionLeftVal = getActualValue(deviceWidth, positionLeftPercent);
	var positionTopPercent = jsonObj[i]["positionTopPercent"];
	var positionTopVal = getActualValue(deviceHeight, positionTopPercent);
	var positionWidthPercent = jsonObj[i]["positionWidthPercent"];
	var positionWidthVal = getActualValue(deviceWidth, positionWidthPercent);
	
	var positionLeftNew = jsonObj[i]["positionLeftNew"];
	var positionLeftNewVal = getActualValue(deviceWidth, positionLeftNew);
	var positionTopNew = jsonObj[i]["positionTopNew"];
	var positionTopNewVal = getActualValue(deviceWidth, positionTopNew);
	/**
	 * Added for lock and unlock feature
	 */
	var divLockId = jsonObj[i]["divLockId"];
	var lockImage = jsonObj[i]["lockImage"];
	
	/*var divLockId = "divImg_" + cnt+"_lock";
	var lockImage = "../img/imgUnlock.png";*/
	/******* END *******/
	
	if(navigator.userAgent.match(/Android/i) || navigator.userAgent.match(/iPhone/i)) {
		var $element2 = $('<div id="' + divIda +'" class="col-md-4 single_task_item_sketch after_sketch_task draggableResizable common_elem" style="position: absolute; top:'+positionTopNewVal+'px; left:'+positionLeftNewVal+'px; width:'+positionWidthVal+'px ;display:'+divStyle+'; height:'+positionWidthVal+'px;"><div class="col-md-12 task_item_top_sketch"><div class="col-md-3 col-xs-3 task" onclick=enableDraggableResizable("'+divIda+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 col-xs-5 energy" onclick=enableDraggableResizable("'+divIda+'") onblur="disableDraggableResizable()">Time/ Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyIda+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onchange="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm_sketch"><textarea class="col-md-12 col-xs-12 add_task_area_sketch" style="text-align: center;" id="'+areaIda+'" onkeypress="return disableKey(event)" type="text" name="areaName" maxlength="200" placeholder="Add Task Label Here" onBlur="disableDraggableResizable()">'+areaValue+'</textarea></div><div class="col-md-12 additional_area"><img class="person_icon" src="../img/personIcon.png"><textarea class="form-control-additional input-sm" id="'+additionalId+'" onkeypress="return disableKey(event)" name="additional" maxlength="120" placeholder="Add Relational Crafting Here" onblur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divIda+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divIda+'","'+divLockId+'")><img src="'+lockImage+'" alt="Lock" class="lock_unlock_img"/></div></div>');
	} else if(navigator.userAgent.match(/iPad/i) || navigator.userAgent.match(/IPad/i)) {
		var $element2 = $('<div id="' + divIda +'" class="col-md-4 single_task_item_sketch after_sketch_task draggableResizable common_elem" style="position: absolute; top:'+positionTopNewVal+'px; left:'+positionLeftNewVal+'px; width:'+positionWidthVal+'px ;display:'+divStyle+'; height:'+positionWidthVal+'px;"><div class="col-md-12 task_item_top_sketch"><div class="col-md-3 col-xs-3 task" onclick=enableDraggableResizable("'+divIda+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 col-xs-5 energy" onclick=enableDraggableResizable("'+divIda+'") onblur="disableDraggableResizable()">Time/Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyIda+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onchange="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm_sketch"><textarea class="col-md-12 col-xs-12 add_task_area_sketch" style="text-align: center;" id="'+areaIda+'" type="text" onkeypress="return disableKey(event)" name="areaName" maxlength="200" placeholder="Add Task Label Here" onBlur="disableDraggableResizable()">'+areaValue+'</textarea></div><div class="col-md-12 additional_area"><img class="person_icon" src="../img/personIcon.png"><textarea class="form-control-additional input-sm" id="'+additionalId+'" onkeypress="return disableKey(event)" name="additional" maxlength="120" placeholder="Add Relational Crafting Here" onblur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divIda+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divIda+'","'+divLockId+'")><img src="'+lockImage+'" alt="Lock" class="lock_unlock_img"/></div></div>');
	} else {
		var $element2 = $('<div id="' + divIda +'" class="col-md-4 single_task_item_sketch after_sketch_task draggableResizable common_elem" style="position: absolute; top:'+positionTopNewVal+'px; left:'+positionLeftNewVal+'px; width:'+positionWidthVal+'px ;display:'+divStyle+'; height:'+positionWidthVal+'px;"><div class="col-md-12 task_item_top_sketch"><div class="col-md-3 col-xs-3 task" onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 col-xs-5 energy">Time/Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyIda+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onchange="return getEnergyValue(this)" /></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm_sketch"><textarea class="col-md-12 col-xs-12 add_task_area_sketch" style="text-align: center;" id="'+areaIda+'" type="text" onkeypress="return disableKey(event)" name="areaName" maxlength="200" placeholder="Add Task Label Here">'+areaValue+'</textarea></div><div class="col-md-12 additional_area"><img class="person_icon" src="../img/personIcon.png"><textarea class="form-control-additional input-sm" id="'+additionalId+'" onkeypress="return disableKey(event)" name="additional" maxlength="120" placeholder="Add Relational Crafting Here"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divIda+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divIda+'","'+divLockId+'")><img src="'+lockImage+'" alt="Lock" class="lock_unlock_img"/></div></div>');
	}
	$("#outderDiv").append($element2);
	/***************** To Disable the paste functionality ISSUE# 186***************/
    $('#' +areaIda).bind("paste",function(e) {
        e.preventDefault();
    });
	$element2.draggable();	
	/*$element2.resizable({
	    handles: 'se'		    
	});*/	
	$element2.resizable({			
		//ghost: 'true',
		handles: 'se',
		aspectRatio: 2 / 2,
	});	
	generateJsonObj();
	
	$element2.resizable({
		stop:function(event,ui){
			var divId = $(this).attr('id');
			var finalWidth = $(this).width();     	
			calculateEnergy(finalWidth, divId);					
		}
	});	
	/**************** To reduce the Z-index value to fix the UI**********************/
	$(".single_task_item_sketch .ui-resizable-se").css({ 'z-index':'0'});
	setClassAgainstLockUnlockElememt(divIda, divLockId, lockImage);
}
fixDesignAS();
setTaskWhileLoad();
setOuterDivWhileDelete();
}

			
/**
 * Function add to calculate width difference
 * @param widthFinal, divIdFinal
 */
function calculateEnergy(widthFinal, divIdFinal){
	var status = null;
	var taskHeight = "";
	var taskWidth =  "";	
	var mainDivWidthPx = "";
	var outerMainDivWidth = "";
	var positionLeft = "";
	var mainDivWidth = $('#outderDiv').width();
	if( null == document.getElementById("pageWrapCopy").style.width || document.getElementById("pageWrapCopy").style.width == ""){
		if (navigator.userAgent.match(/iPhone|iPad|iPod/i)) { // added for issue #264
			mainDivWidthPx = "858px";
		} else {
			mainDivWidthPx = "1043px";
		}	
	 } else {
		mainDivWidthPx = document.getElementById("pageWrapCopy").style.width;
	 }	
	outerMainDivWidth = parseInt(mainDivWidthPx.replace("px", ''));	
	var initialDiv = JSON.parse(ClientSession.get("div_intial_value"));	
	divLength = initialDiv.length;		
	var divCnt = divIdFinal.substring(7);
	var differenceWidth = "";
	var positionWidthPercent = initialDiv[divCnt-1]["positionWidthPercent"];
	var initialWidth = getActualValue(mainDivWidth, positionWidthPercent);
	//var initialWidth = initialDiv[divCnt-1]["divWidthPixel"];
	var differenceValue = parseInt(widthFinal - initialWidth);
	if(mainDivWidth >1000) {
		differenceWidth = Math.round(parseInt(differenceValue/3));
	} else {		
		differenceWidth = getWidthDifference(parseInt(differenceValue/3), mainDivWidth);		
	}
	//differenceWidth = Math.round(parseInt(differenceValue/3));
	if(differenceWidth>0){
		status = "+";
	} else {
		differenceWidth = -(differenceWidth);
		status = "-";
	}	
	
	/*************** LOCK UNLOCK ******************/
	var unlockDivCount = $('#outderDiv').children('.single_task_item_sketch:visible').length;	//count the unlock div
	if(unlockDivCount == 1){
		calculateTotalEnergy();
	} else if(status=='+'){
		differenceWidth = getMaxEnergyValueToAdd(divIdFinal, differenceWidth, "RESIZE");
	}
	/**** END ****/
	
	taskHeight = $("#"+divIdFinal).height();
	taskWidth = $("#"+divIdFinal).width();	
	positionLeft = $("#"+divIdFinal).position().left;
	if(unlockDivCount == 1){
		//alertify.alert("Please unlock one task to resize");
		calculateTotalEnergy();
	} else if(differenceWidth == 0){
		calculateTotalEnergy();
	} else if(navigator.userAgent.match(/Android|iPhone/i)) {
		if(widthFinal > 225 || widthFinal < 70) {
			calculateTotalEnergy();
		} else if((taskHeight - taskWidth)>15 || (taskHeight - taskWidth)<-15){
			//calculateTotalEnergy();
		} else {
			calculateEnergyValue(differenceWidth,status,divIdFinal, "RESIZE");
		}
	} else if(navigator.userAgent.match(/iPad/i)) {
		if(widthFinal > 335 || widthFinal < 112) {
			calculateTotalEnergy();
		} else if((taskHeight - taskWidth)>15 || (taskHeight - taskWidth)<-15){
			//calculateTotalEnergy();
		} else {
			calculateEnergyValue(differenceWidth,status,divIdFinal, "RESIZE");
		}
	} else if(!(navigator.userAgent.match(/iPhone|iPad|iPod/i)) && !(navigator.userAgent.match(/Android/i)) ){
		if(widthFinal > 395 || widthFinal < 140) {
			calculateTotalEnergy();
		} else if((taskHeight - taskWidth)>15 || (taskHeight - taskWidth)<-15){
			//calculateTotalEnergy();
		} else {
			calculateEnergyValue(differenceWidth,status,divIdFinal, "RESIZE");
		}
	 } else {
		calculateEnergyValue(differenceWidth,status,divIdFinal, "RESIZE");
	}
}

/**
 * Function add to set the initial position if the total task is more thn 100 
 * @param widthFinal, divIdFinal
 */
function calculateTotalEnergy() {
	var mainDivWidth = $('#outderDiv').width();
	var initialDiv = JSON.parse(ClientSession.get("div_intial_value"));	
	divLength = initialDiv.length;		
	for (var i = 0; i<divLength;i++) {
		var divId = initialDiv[i]["divId"];
		var energyId = initialDiv[i]["energyId"];
		var energyValue = initialDiv[i]["energyValue"];
		//var divWidthPixel = initialDiv[i]["divWidthPixel"];	
		var positionWidthPercent = initialDiv[i]["positionWidthPercent"];
		var divWidthPixel = getActualValue(mainDivWidth, positionWidthPercent);		
		var elemEnergyId = document.getElementById(energyId);	
		var elemDivId = document.getElementById(divId);	
		elemEnergyId.value = energyValue;
		elemDivId.style.width = divWidthPixel+"px";
		elemDivId.style.height = divWidthPixel+"px";		
	}
	
	/*************************To fix the design ******************************/
	fixDesignAS();	
}


/**
 * Function add to calculate and set energy value and width 
 * @param energyValue
 * @param status
 * @param divIdFinal
 * @param action
 */
function calculateEnergyValue(energyValue,status,divIdFinal, action) {
	var tempValue = energyValue;
	 var check = true;
	/********** Calculate the total energy value *****************/	 	 
	 var enegryValTotal = 0;
	 var initialJsonVal = JSON.parse(ClientSession.get("div_intial_value"));	
	 var jsonLength = initialJsonVal.length;	
	 for (var i = 0; i<jsonLength;i++) {	
		 enegryValTotal = enegryValTotal + parseInt(initialJsonVal[i]["energyValue"]);
	 }	
	 var mainDivWidth = $('#outderDiv').width();
	/********** Called agsinst add task and when resize the boxes to increase the energy value *****************/
	if(status=='+'){
		var initialDiv = JSON.parse(ClientSession.get("div_intial_value"));	
		divLength = initialDiv.length;	
		var divCount = $('#outderDiv').children('.single_task_item_sketch:visible').length;
		var maxValue = 100 - ((divCount-1)*5);
		var divCountLoop = $('#outderDiv').children('.single_task_item_sketch').length;		
		totalDiv = divCount+1;	
		if(divCount<19){
			if(enegryValTotal == 99){
				for (var l=0;l<tempValue-1;){
					for (var i = 0; i<divLength;i++) {
						var divId = initialDiv[i]["divId"];
						var energyValue = initialDiv[i]["energyValue"];
						
						var positionWidthPercent = initialDiv[i]["positionWidthPercent"];
						var divWidthPixel = getActualValue(mainDivWidth, positionWidthPercent);					
						var style = initialDiv[i]["style"];	
						var lockImage = initialDiv[i]["lockImage"];	
							if (divId == divIdFinal){
								var energVal= parseInt(energyValue)+parseInt(tempValue);
								var elem = document.getElementById("energyId_" +(i+1));	
								var elemPixel = document.getElementById("divImg_" +(i+1));	
								elem.value = energVal;
								
								if(mainDivWidth >1000) {
									 widthValuePixel = Math.round(parseInt(divWidthPixel+(tempValue*3)));
								} else {
									var addValue = getWidthDifference(parseInt((tempValue*3)), mainDivWidth);	
									widthValuePixel = Math.round(parseInt(divWidthPixel+addValue));	
								}								
								elemPixel.style.width = widthValuePixel+"px";
								elemPixel.style.height = widthValuePixel+"px";
							}
						
						else{
							if((style == null) || (style == "null") || (style.search('block') != -1)){	
								if (lockImage.search('imgUnlock.png') != -1) {	
								if(energyValue>5){
									if(l  < tempValue){
										energyValue = energyValue-1;									
										initialDiv[i]["energyValue"] = energyValue;
										var elem = document.getElementById("energyId_" +(i+1));	
										var elemPixel = document.getElementById("divImg_" +(i+1));										
										elem.value = energyValue;
										var k = i;
										var jsonForEnergy = JSON.parse(ClientSession.get("div_intial_value"));
										var initialEnergyVal = jsonForEnergy[k]["energyValue"];
										
										if(mainDivWidth >1000) {
											 widthValuePixel = Math.round(parseInt(divWidthPixel-((initialEnergyVal - energyValue)*3)));
										} else {
											var addValue = getWidthDifference(parseInt(((initialEnergyVal - energyValue)*3)), mainDivWidth);												
											widthValuePixel = Math.round(parseInt(divWidthPixel- addValue));	
										}
										elemPixel.style.width = widthValuePixel+"px";
										elemPixel.style.height = widthValuePixel+"px";
										}
										l=l+1;
								}
							}
							}														
						}					
					}
				}
			} else {
				for (var l=0;l<tempValue;){
					for (var i = 0; i<divLength;i++) {
						var divId = initialDiv[i]["divId"];
						var energyValue = initialDiv[i]["energyValue"];				
						var positionWidthPercent = initialDiv[i]["positionWidthPercent"];
						var divWidthPixel = getActualValue(mainDivWidth, positionWidthPercent);
						var style = initialDiv[i]["style"];	
						var lockImage = initialDiv[i]["lockImage"];
							if (divId == divIdFinal){
								var energVal= parseInt(energyValue)+parseInt(tempValue);
								var elem = document.getElementById("energyId_" +(i+1));	
								var elemPixel = document.getElementById("divImg_" +(i+1));	
								elem.value = energVal;
								if (energVal > maxValue){									
									check = false;
									calculateTotalEnergy();
									return false;
								} else {
									elem.value = energVal;									
									if(mainDivWidth >1000) {
										widthValuePixel = Math.round(parseInt(divWidthPixel+(tempValue*3)));
									} else {
										var addValue = getWidthDifference(parseInt(tempValue*3), mainDivWidth);												
										widthValuePixel = Math.round(parseInt(divWidthPixel + addValue));											
									}
									elemPixel.style.width = widthValuePixel+"px";
									elemPixel.style.height = widthValuePixel+"px";
								}	
						}						
						else if (check == true) {
							if((style == null) || (style == "null") || (style.search('block') != -1)){
								if (lockImage.search('imgUnlock.png') != -1) {	
								if(energyValue>5){
									if(l  < tempValue){
										energyValue = energyValue-1;									
										initialDiv[i]["energyValue"] = energyValue;
										var elem = document.getElementById("energyId_" +(i+1));	
										var elemPixel = document.getElementById("divImg_" +(i+1));										
										elem.value = energyValue;
										var k = i;
										var jsonForEnergy = JSON.parse(ClientSession.get("div_intial_value"));
										var initialEnergyVal = jsonForEnergy[k]["energyValue"];
										
										if(mainDivWidth >1000) {
											widthValuePixel = Math.round(parseInt(divWidthPixel-((initialEnergyVal - energyValue)*3)));
										} else {
											var addValue = getWidthDifference(parseInt(((initialEnergyVal - energyValue)*3)), mainDivWidth);												
											widthValuePixel = Math.round(parseInt(divWidthPixel - addValue));	
										}										
										elemPixel.style.width = widthValuePixel+"px";
										elemPixel.style.height = widthValuePixel+"px";
										}
										l=l+1;
								}
							}
							}														
						}					
					}
				}
			}			
		} else if(divCount >= 19){
			if( action == "ADD"){
				for (var i = 0; i<divLength;i++) {
					var enrgyVal = 5;
					var elem = document.getElementById("energyId_" +(i+1));	
					var elemPixel = document.getElementById("divImg_" +(i+1));						
					if(mainDivWidth >1000) {
						widthValuePixel = 140;
					} else {
						widthValuePixel = getWidthDifference(140, mainDivWidth);									
					}
					elemPixel.style.width = widthValuePixel+"px";
					elemPixel.style.height = widthValuePixel+"px";
					elem.value = enrgyVal;				
					}
				} else {
					for (var l=0;l<tempValue;){						
						for (var i = 0; i<divLength;i++) {					
							var divId = initialDiv[i]["divId"];
							var energyValue = initialDiv[i]["energyValue"];							
							var positionWidthPercent = initialDiv[i]["positionWidthPercent"];
							var divWidthPixel = getActualValue(mainDivWidth, positionWidthPercent);					
							var style = initialDiv[i]["style"];	
							var lockImage = initialDiv[i]["lockImage"];
								if (divId == divIdFinal){
									var energVal= parseInt(energyValue)+parseInt(tempValue);
									var elem = document.getElementById("energyId_" +(i+1));	
									var elemPixel = document.getElementById("divImg_" +(i+1));			
									if (energVal > maxValue){									
										check = false;
										calculateTotalEnergy();
										return false;
									} else {
										elem.value = energVal;
										
										if(mainDivWidth >1000) {
											widthValuePixel = Math.round(parseInt(divWidthPixel+(tempValue*3)));
										} else {
											var addValue = getWidthDifference(parseInt(tempValue*3), mainDivWidth);												
											widthValuePixel = Math.round(parseInt(divWidthPixel + addValue));
										}
										elemPixel.style.width = widthValuePixel+"px";
										elemPixel.style.height = widthValuePixel+"px";
									}								
								}					
							else if (check == true){	
								if((style == null) || (style == "null") || (style.search('block') != -1)){	
									if (lockImage.search('imgUnlock.png') != -1) {
									if(energyValue>5){
										if(l  < tempValue){
											energyValue = energyValue-1;									
											initialDiv[i]["energyValue"] = energyValue;
											var elem = document.getElementById("energyId_" +(i+1));	
											var elemPixel = document.getElementById("divImg_" +(i+1));										
											elem.value = energyValue;
											var k = i;
											var jsonForEnergy = JSON.parse(ClientSession.get("div_intial_value"));
											var initialEnergyVal = jsonForEnergy[k]["energyValue"];
											
											if(mainDivWidth >1000) {
												widthValuePixel = Math.round(parseInt(divWidthPixel-((initialEnergyVal - energyValue)*3)));
											} else {
												var addValue = getWidthDifference(parseInt(((initialEnergyVal - energyValue)*3)), mainDivWidth);												
												widthValuePixel = Math.round(parseInt(divWidthPixel - addValue));											
											}
											elemPixel.style.width = widthValuePixel+"px";
											elemPixel.style.height = widthValuePixel+"px";
											}
											l=l+1;
									} 
								}
								}														
							}					
						}
					}
				}				
			}
	} 
	/********** Called agsinst delete task and when resize the boxes to decrease the energy value *****************/
	else {
		var initialDiv = JSON.parse(ClientSession.get("div_intial_value"));	
		divLength = initialDiv.length;	
		var divCount = $('#outderDiv').children('single_task_item_sketch:visible').length;
		var divCountLoop = $('#outderDiv').children('single_task_item_sketch').length;// total task block
		totalDiv = divCount+1;	
		for (var l=0;l<tempValue;){
			for (var i = 0; i<divLength;i++) {
				var divId = initialDiv[i]["divId"];
				var energyValue = initialDiv[i]["energyValue"];
				
				var positionWidthPercent = initialDiv[i]["positionWidthPercent"];
				var divWidthPixel = getActualValue(mainDivWidth, positionWidthPercent);	
				var style = initialDiv[i]["style"];	
				var lockImage = initialDiv[i]["lockImage"];
					if (divId == divIdFinal){								
						if(energyValue>5){							
							if (action == "DELETE" ) {
								var energVal= parseInt(energyValue)-parseInt(tempValue);							
								var elem = document.getElementById("energyId_" +(i+1));	
								var elemPixel = document.getElementById("divImg_" +(i+1));
								elem.value = energVal;
								if(mainDivWidth >1000) {
									widthValuePixel = Math.round(parseInt(divWidthPixel-(tempValue*3)));
								} else {
									var addValue = getWidthDifference(parseInt(tempValue*3), mainDivWidth);												
									widthValuePixel = Math.round(parseInt(divWidthPixel - addValue));							
								}
								elemPixel.style.height = widthValuePixel+"px";
							} else {
								var energVal= parseInt(energyValue)-parseInt(tempValue);	
								if (energVal >=5) {
									var elem = document.getElementById("energyId_" +(i+1));	
									var elemPixel = document.getElementById("divImg_" +(i+1));
									elem.value = energVal;
									if(mainDivWidth >1000) {
										widthValuePixel = Math.round(parseInt(divWidthPixel-(tempValue*3)));
									} else {
										var addValue = getWidthDifference(parseInt(tempValue*3), mainDivWidth);												
										widthValuePixel = Math.round(parseInt(divWidthPixel - addValue));							
									}				
									elemPixel.style.width = widthValuePixel+"px";
									elemPixel.style.height = widthValuePixel+"px";
								} else {
									calculateTotalEnergy();
								}							
							}														
						}	
					} else{							
						if((style == null) || (style == "null") || (style.search('block') != -1)){	
							if (lockImage.search('imgUnlock.png') != -1) {
							if(l  < tempValue){
								energyValue =  parseInt(energyValue)+1;
								initialDiv[i]["energyValue"] = energyValue;
								var elem = document.getElementById("energyId_" +(i+1));	
								var elemPixel = document.getElementById("divImg_" +(i+1));	
								elem.value = energyValue;								
								var k = i;
								var jsonForEnergy = JSON.parse(ClientSession.get("div_intial_value"));
								var initialEnergyVal = jsonForEnergy[k]["energyValue"];
								
								if(mainDivWidth >1000) {
									widthValuePixel = Math.round(parseInt(divWidthPixel+((energyValue - initialEnergyVal)*3)));
								} else {
									var addValue = getWidthDifference(parseInt(((energyValue - initialEnergyVal)*3)), mainDivWidth);												
									widthValuePixel = Math.round(parseInt(divWidthPixel + addValue));												
								}
								elemPixel.style.width = widthValuePixel+"px";
								elemPixel.style.height = widthValuePixel+"px";
								}
							l=l+1;
						}
						}																	
					}				
				}
			}	
	}
	fixDesignAS();
	adjustTaskWhileResize();
	setOuterDivWhileDelete();
	generateJsonObj();
}

/*************** Json Object of loaded tasks **************/
function generateJsonObj() {
	var cnt = 1;
	var divJson = [];
	var idx = 0;
	var divValue = document.getElementById("divImg_" + cnt);
	var mainDivWidth = $('#outderDiv').width();
	while(divValue != null) {
		var unitJson = {};		
			unitJson["divId"] = document.getElementById("divImg_" + cnt).id;
			unitJson["energyId"] = document.getElementById("energyId_"+ cnt).id;
			unitJson["energyValue"] = document.getElementById("energyId_"+ cnt).value;
			unitJson["divWidthPercent"] = document.getElementById("divImg_" + cnt).style.width;	
			unitJson["style"] = document.getElementById("divImg_" + cnt).getAttribute("style");
			var divWidthPercent= unitJson["divWidthPercent"];
			if(divWidthPercent.search("px") != -1) {
				var divWidthPixel = parseInt(divWidthPercent.replace("px",''));
			}
			if(divWidthPercent.search("%") != -1) {
				var divWidth = divWidthPercent.replace("%",'');
				var divWidthPixel =Math.round(parseInt(divWidth)*10.63);
			}			
			unitJson["divWidthPixel"] = divWidthPixel;
			unitJson["divHeight"] = document.getElementById("divImg_" + cnt).style.height;
			
			var positionWidthPixel = document.getElementById("divImg_" + cnt).style.width;	
			var positionWidth = parseInt(positionWidthPixel.replace("px", ''));
			unitJson["positionWidthPercent"] = getPercentageValue(mainDivWidth, positionWidth);
			unitJson["lockImage"] = document.getElementById("divImg_" + cnt+ "_lock").getElementsByTagName("img")[0].getAttribute("src");	
			divJson[idx++] = unitJson;
			divValue = document.getElementById("divImg_" + ++cnt);
	}
	sessionStorage.setItem("div_intial_value",JSON.stringify(divJson));
}



/**
 * Function add to calculate the Energy Value
 * @param null
 */
function calculateEnergySameValue() {
	var countEng = 0;
	var divCount = $('#outderDiv').children('.single_task_item_sketch:visible').length;
	var divCountLoop = $('#outderDiv').children('.single_task_item_sketch').length;
	totalDiv = divCount+1;
	var energyValue = Math.round(100/totalDiv);
	//var energyStyle = energyValue+5;
	var totalEnergy =  energyValue*totalDiv;
	
	for(var i=1; i<=divCountLoop;i++) {
		var divIdLock = "divImg_"+i+"_lock";
		var lockImage = document.getElementById(divIdLock).getElementsByTagName("img")[0].getAttribute("src");
		var style = document.getElementById("divImg_" +i).getAttribute("style");
		if((style.search('block') != -1)){
			if (lockImage.search('imgUnlock.png') != -1) {
			if(totalEnergy<101){				
				var elem = document.getElementById("energyId_" +i);		
				elem.value = energyValue;
			} else {
				energyValDiff = totalEnergy - 100;						
					var energyValDiff = totalEnergy -100;
						countEng++;
						if(countEng<=energyValDiff){							
							var elem = document.getElementById("energyId_" +i);		
							elem.value = energyValue-1;
						}  else {
							var elem = document.getElementById("energyId_" +i);		
							elem.value = energyValue;
						}							
					}	
				}
			}				
	}	
}

/**
 * Function to delete the task div
 * @param id
 */
function deleteDiv(id) {
	removeFocus();
	var divCount = $('#outderDiv').children('.after_sketch_task:visible').length;
	if(divCount>3){
		$("#" + id).children().find(".add_task_area_sketch").addClass("add_task_area_sketch_delete_focus");
		var unlockDivCount = $('#outderDiv').children('.single_task_item_sketch:visible').length;		
		if(unlockDivCount > 1) {
			alertify.set({ buttonReverse: true });
			alertify.set({ labels: {
			    ok     : "Yes",
			    cancel : "No"
			} });
			alertify.confirm("<img src='../img/confirm-icon.png'><br /><p>Are you sure you want to delete this task block?</p>", function (e) {
			    if (e) {	
			    	var divCnt = id.substring(7);	
			    	var deleteEnergyVal = document.getElementById("energyId_" +divCnt).value;	
			    	calculateEnergyValue(deleteEnergyVal, "-", id, "DELETE");
				    var div = document.getElementById(id);
				    if (div) {
				    	div.setAttribute("style", "display:none");			    	
				    }
				    //setOuterDivWhileDelete("decrease");
				    generateJsonObj();
			    } else {
			    	$("#" + id).children().find(".add_task_area_sketch").removeClass("add_task_area_sketch_delete_focus");
			    }
			});		
		} else {
			alertify.alert("<img src='../img/alert-icon.png'><br /><p>Cannot delete the task since rest all the tasks are locked.</p>");
			$('#alertify-ok').click(function() {
				$("#" + id).children().find(".add_task_area_sketch").removeClass("add_task_area_sketch_delete_focus");
			});		
			
			}	
	} else {
		alertify.alert("<img src='../img/alert-icon.png'><br /><p>Minimum three tasks are mandatory.</p>");
		$('#alertify-ok').click(function() {
			$("#" + id).children().find(".add_task_area_sketch").removeClass("add_task_area_sketch_delete_focus");
		});	
	}	
	
}


/**************** For add task functionality **********************/
var allCanvasArray = new Array();
var count = 1;
$('#addtask').click(function() {
	var unlockDivCount = $('#outderDiv').children('.single_task_item_sketch:visible').length;
	if(sessionStorage.getItem("bsEdit")){
		alertify.alert("<img src='../img/alert-icon.png'><br /><p>Cannot add tasks as you have already completed your after diagram.</p>");
		return false;
	} 
	/****************** check for lock and unlock task div ********************/
	else if (unlockDivCount == 0) {
		alertify.alert("<img src='../img/alert-icon.png'><br /><p>Cannot add tasks as your all tasks are blocked.</p>");
		return false;
	} else {
	var energyValue = null;
	var energyStyle = null;
	
	/****************** Check the if number of tasks are more than 20 ********************/
	var divCount = $('#outderDiv').children('.after_sketch_task:visible').length;	
	if(divCount>=20){
		alertify.alert("<img src='../img/alert-icon.png'><br /><p>You can have a maximum of 20 tasks.</p>");
		return false;
	}

	/****************** Check the Energy value of each Task ********************/
	
	var array = new Array();
	var divCount = $('#outderDiv').children('.after_sketch_task:visible').length;
	var divCountLoop = $('#outderDiv').children('.after_sketch_task').length;
	for(var i=1; i<=divCountLoop;i++){
		var style = document.getElementById("divImg_" +i).getAttribute("style");
		if((style.search('block') != -1)){
			var elem = document.getElementById("energyId_" +i);		
			array.push(elem.value);	
		}	
	}	
			
	var check =true;
	var firstelement = array[0];
	var element=null;
	var lockDivCount = $('#outderDiv').children('.lock_single_task_item_sketch:visible').length;
	for(var i = 0; i<array.length;i++){
		element = array[i];	
		if(lockDivCount > 0) {
			check = false;
			break;
		} else if(element !=firstelement) {
			check = false;
			break;
		} else { 
			check = true;	
		}		
	}		
	var mainDivWidth = $('#outderDiv').width();
	/****************** When the energy values are same ********************/
	if(check==true){			
		var existingDivId = $('#outderDiv').children('.after_sketch_task:visible').attr('id');	
		var divCnt = existingDivId.substring(7);	    			
		var previousWidth = document.getElementById(existingDivId).style.width;
		var previousEnergyVal = document.getElementById("energyId_" +divCnt).value;	
		var previousWidthValue = parseInt(previousWidth.replace("px",''));	
		var totalDiv = divCount+1;
		energyValue = Math.round(100/totalDiv);
		
		if(mainDivWidth >1000) {
			energyStyle = previousWidthValue - (parseInt(previousEnergyVal - energyValue)*3);	
		} else {
			var diffValue = getWidthDifference(parseInt(previousEnergyVal - energyValue)*3, mainDivWidth);	
			energyStyle = previousWidthValue - diffValue;
		}	
	    if(next_page_sketch == 0){
	    	count = ClientSession.get("total_count");
	    	var divId = "divImg_" + count;
			var divIdx = "divImg_" + count+"_x";		
			var energyId = "energyId_" + count;
			var areaId = "areaId_" + count;		
			var additionalId = "additionalId_" + count;
			var divIdLock = "divImg_" + count+"_lock";
			if(navigator.userAgent.match(/Android/i) || navigator.userAgent.match(/iPhone/i)) {
				var $element = $('<div id="' + divId +'" class="col-md-4 after_sketch_task single_task_item_sketch draggableResizable common_elem" style="position: absolute; display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top_sketch"><div class="col-md-3 col-xs-3 task" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 energy" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/ Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onchange="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm_sketch"><textarea class="well col-md-12 col-xs-12 add_task_area_sketch" style="text-align: center;" id="'+areaId+'" type="text" onkeypress="return disableKey(event)" name="areaName" value="value" maxlength="200" placeholder="Add Task Label Here" onBlur="disableDraggableResizable()"></textarea></div><div class="col-md-12 additional_area"><img class="person_icon" src="../img/personIcon.png"><textarea class="form-control-additional input-sm" id="'+additionalId+'" onkeypress="return disableKey(event)" name="additional" maxlength="120" placeholder="Add Relational Crafting Here" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divIdLock +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divIdLock+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');				
			} else if(navigator.userAgent.match(/iPad/i) || navigator.userAgent.match(/IPad/i)) {
				var $element = $('<div id="' + divId +'" class="col-md-4 after_sketch_task single_task_item_sketch draggableResizable common_elem" style="position: absolute; display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top_sketch"><div class="col-md-3 col-xs-3 task" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 energy" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onchange="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm_sketch"><textarea class="well col-md-12 col-xs-12 add_task_area_sketch" style="text-align: center;" id="'+areaId+'" type="text" onkeypress="return disableKey(event)" name="areaName" value="value" maxlength="200" placeholder="Add Task Label Here" onBlur="disableDraggableResizable()"></textarea></div><div class="col-md-12 additional_area"><img class="person_icon" src="../img/personIcon.png"><textarea class="form-control-additional input-sm" id="'+additionalId+'" onkeypress="return disableKey(event)" name="additional" maxlength="120" placeholder="Add Relational Crafting Here" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divIdLock +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divIdLock+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');				
			} else {
				var $element = $('<div id="' + divId +'" class="col-md-4 after_sketch_task single_task_item_sketch draggableResizable common_elem" style="position: absolute; display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top_sketch"><div class="col-md-3 col-xs-3 task">TASK</div><div class="col-md-6 col-xs-5 energy">Time/Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onchange="return getEnergyValue(this)" /></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm_sketch"><textarea class="well col-md-12 col-xs-12 add_task_area_sketch" style="text-align: center;" id="'+areaId+'" type="text" onkeypress="return disableKey(event)" name="areaName" value="value" maxlength="200" placeholder="Add Task Label Here"></textarea></div><div class="col-md-12 additional_area"><img class="person_icon" src="../img/personIcon.png"><textarea class="form-control-additional input-sm" id="'+additionalId+'" onkeypress="return disableKey(event)" name="additional" maxlength="120" placeholder="Add Relational Crafting Here"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divIdLock +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divIdLock+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');
			}		
	    } else {
			 count = ClientSession.get("total_count");
			 var divId = "divImg_" + count;
			 var divIdx = "divImg_" + count+"_x";		
			 var energyId = "energyId_" + count;
			 var areaId = "areaId_" + count;
			 var additionalId = "additionalId_" + count;
			 var divIdLock = "divImg_" + count+"_lock";
			 if(navigator.userAgent.match(/Android/i) || navigator.userAgent.match(/iPhone/i)) {
				 var $element = $('<div id="' + divId +'" class="col-md-4 after_sketch_task single_task_item_sketch draggableResizable common_elem" style="position: absolute; display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top_sketch" ><div class="col-md-3 col-xs-3 task" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 energy" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/ Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onchange="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm_sketch"><textarea class="well col-md-12 col-xs-12 add_task_area_sketch" style="text-align: center;" id="'+areaId+'" type="text" onkeypress="return disableKey(event)" name="areaName" value="value" maxlength="200" placeholder="Add Task Label Here" onBlur="disableDraggableResizable()"></textarea></div><div class="col-md-12 additional_area"><img class="person_icon" src="../img/personIcon.png"><textarea class="form-control-additional input-sm" id="'+additionalId+'" name="additional" maxlength="120" placeholder="Add Relational Crafting Here" onkeypress="return disableKey(event)" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divIdLock +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divIdLock+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');				
			 } else if(navigator.userAgent.match(/iPad/i) || navigator.userAgent.match(/IPad/i)) {
				 var $element = $('<div id="' + divId +'" class="col-md-4 after_sketch_task single_task_item_sketch draggableResizable common_elem" style="position: absolute; display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top_sketch"><div class="col-md-3 col-xs-3 task" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()>TASK</div><div class="col-md-6 col-xs-5 energy" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onchange="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm_sketch"><textarea class="well col-md-12 col-xs-12 add_task_area_sketch" style="text-align: center;" id="'+areaId+'" type="text" onkeypress="return disableKey(event)" name="areaName" value="value" maxlength="200" placeholder="Add Task Label Here" onBlur="disableDraggableResizable()"></textarea></div><div class="col-md-12 additional_area"><img class="person_icon" src="../img/personIcon.png"><textarea class="form-control-additional input-sm" id="'+additionalId+'" name="additional" maxlength="120" placeholder="Add Relational Crafting Here" onkeypress="return disableKey(event)" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divIdLock +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divIdLock+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');
			 } else {
				 var $element = $('<div id="' + divId +'" class="col-md-4 after_sketch_task single_task_item_sketch draggableResizable common_elem" style="position: absolute; display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top_sketch"><div class="col-md-3 col-xs-3 task">TASK</div><div class="col-md-6 col-xs-5 energy">Time/Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onchange="return getEnergyValue(this)" /></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm_sketch"><textarea class="well col-md-12 col-xs-12 add_task_area_sketch" style="text-align: center;" id="'+areaId+'" type="text" onkeypress="return disableKey(event)" name="areaName" value="value" maxlength="200" placeholder="Add Task Label Here"></textarea></div><div class="col-md-12 additional_area"><img class="person_icon" src="../img/personIcon.png"><textarea class="form-control-additional input-sm" id="'+additionalId+'" onkeypress="return disableKey(event)" name="additional" maxlength="120" placeholder="Add Relational Crafting Here"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divIdLock +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divIdLock+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');
			 }			 			   
	    }	
	  	 count++;	  	
	  	 ClientSession.set("total_count", count);
	  	 calculateEnergySameValue();
	  	 
	  	/************** To remove the focus to the new old task *************/
	  	 removeFocus();
	  	 
	     $("#outderDiv").append($element);	 
	     addTaskPosition(energyStyle, divId);
		 $element.draggable();		 
		 $element.resizable({			
				//ghost: 'true',
				handles: 'se',
				aspectRatio: 2 / 2,
		});
	     allCanvasArray.push($element);
	     $(".single_task_item_sketch").css({
			    width: energyStyle+"px",
			    height: energyStyle+"px"
			  }, 200);
	     $(".single_task_item_sketch").css({ 'z-index':'4'}); 
	     $('.single_task_item_sketch').draggable({
	 		cursor: 'move',       
	 		containment: '#outderDiv'    
	 	});  
	     /********* Function add to make the time/ energy field and task description area editable ******/
	     contentEditable();
	     /***************** To set the focus to the new Task ***************/
	     $("#" + divId).children().find(".add_task_area_sketch").addClass("add_task_area_sketch_focus");
	     $("#" + areaId).focus();
	     
	     /***************** To Disable the paste functionality ISSUE# 186***************/
	     $('#' +areaId).bind("paste",function(e) {
	        e.preventDefault();
	     });
		    
	     $('input[type=text], textarea').placeholder();
	     /*************************To fix the design ******************************/	     	     
	     fixDesignAS();
	     adjustTaskWhileResize();
	     
	   }
	
	/****************** When the energy values are different ********************/
	if(check==false){
		var status = "+";
		//var positionWidth = null;
		var totalDiv = divCount+1;
		var energyValToAdd = 0;	
		if(totalDiv==20){
			energyValue = 5;
			if(mainDivWidth >1000) {
				energyStyle = 140;
			} else {
				energyStyle = getWidthDifference(140, mainDivWidth);				
			}
			
			energyValToAdd = checkEnergyValWhileAdd();
			if(energyValToAdd < 5) {
				energyValToAdd = 333;
			}
		} else {
			energyValue = 10;			
			if(mainDivWidth >1000) {
				energyStyle = 155;
			} else {
				energyStyle = getWidthDifference(155, mainDivWidth);
			}
			energyValToAdd = checkEnergyValWhileAdd();
			if(totalDiv==22){ 
				if(energyValToAdd < 10) {
					energyValToAdd = 333;
				}
			} else {				
				if(energyValToAdd >= 5 && energyValToAdd < 10) {
					energyValue = 5;
					energyStyle = getWidthDifference(140, mainDivWidth);
				} else if (energyValToAdd < 5) {
					energyValToAdd = 333;
				}				
			}	
		}	
		if(energyValToAdd == 333 ){
			alertify.alert("<img src='../img/alert-icon.png'><br /><p>Task will not be added.</p>");
			return false;
		} 
		else if(next_page_sketch == 0){
			count = ClientSession.get("total_count");
		    var divId = "divImg_" + count;
			var divIdx = "divImg_" + count+"_x";		
			var energyId = "energyId_" + count;
			var areaId = "areaId_" + count;		
			var additionalId = "additionalId_" + count;
			var divIdLock = "divImg_" + count+"_lock";
			if(totalDiv==20){
				if(navigator.userAgent.match(/Android/i)) {
					var $element = $('<div id="' + divId +'" class="col-md-4 after_sketch_task single_task_item_sketch draggableResizable upto_9 common_elem" style="position: absolute; display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top_sketch"><div class="col-md-3 col-xs-3 task task_smaller" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6  energy energy_smaller" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/ Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch form-control-sketch_smaller input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onchange="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm_sketch"><textarea class="col-md-12 col-xs-12 add_task_area_sketch" style="text-align: center;" id="'+areaId+'" type="text" onkeypress="return disableKey(event)" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onBlur="disableDraggableResizable()"></textarea></div><div class="col-md-12 additional_area"><img class="person_icon" src="../img/personIcon.png"><textarea class="form-control-additional form-control-additional_smaller input-sm" id="'+additionalId+'" onkeypress="return disableKey(event)" name="additional" maxlength="120" placeholder="Add Relational Crafting Here" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divIdLock +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divIdLock+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');					
				}else if(navigator.userAgent.match(/iPhone/i)) {
					var $element = $('<div id="' + divId +'" class="col-md-4 after_sketch_task single_task_item_sketch draggableResizable upto_9 common_elem" style="position: absolute; display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top_sketch"><div class="col-md-3 col-xs-3 task task_smaller" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6  energy energy_smaller" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/ Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch form-control-sketch_smaller input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onchange="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm_sketch"><textarea class="col-md-12 col-xs-12 add_task_area_sketch" style="text-align: center;" id="'+areaId+'" type="text" onkeypress="return disableKey(event)" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onBlur="disableDraggableResizable()"></textarea></div><div class="col-md-12 additional_area"><img class="person_icon" src="../img/personIcon.png"><textarea class="form-control-additional form-control-additional_smaller input-sm" id="'+additionalId+'" onkeypress="return disableKey(event)" name="additional" maxlength="120" placeholder="Add Relational Crafting Here" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divIdLock +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divIdLock+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');
				}else if(navigator.userAgent.match(/iPad/i)){
					var $element = $('<div id="' + divId +'" class="col-md-4 after_sketch_task single_task_item_sketch draggableResizable upto_9 common_elem" style="position: absolute; display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top_sketch"><div class="col-md-3 col-xs-3 task task_smaller" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6  energy energy_smaller" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch form-control-sketch_smaller input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onchange="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm_sketch"><textarea class="col-md-12 col-xs-12 add_task_area_sketch" style="text-align: center;" id="'+areaId+'" type="text" onkeypress="return disableKey(event)" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onBlur="disableDraggableResizable()"></textarea></div><div class="col-md-12 additional_area"><img class="person_icon" src="../img/personIcon.png"><textarea class="form-control-additional form-control-additional_small input-sm" id="'+additionalId+'" onkeypress="return disableKey(event)" name="additional" maxlength="120" placeholder="Add Relational Crafting Here" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divIdLock +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divIdLock+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');
				}else {
					var $element = $('<div id="' + divId +'" class="col-md-4 after_sketch_task single_task_item_sketch draggableResizable uptoDesc_9 common_elem" style="position: absolute; display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top_sketch"><div class="col-md-3 col-xs-3 task task_smaller">TASK</div><div class="col-md-6  energy energy_smaller">Time/Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch form-control-sketch_smaller input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onchange="return getEnergyValue(this)" /></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm_sketch"><textarea class="col-md-12 col-xs-12 add_task_area_sketch" style="text-align: center;" id="'+areaId+'" type="text" onkeypress="return disableKey(event)" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200"></textarea></div><div class="col-md-12 additional_area"><img class="person_icon" src="../img/personIcon.png"><textarea class="form-control-additional form-control-additional_small input-sm" id="'+additionalId+'" onkeypress="return disableKey(event)" name="additional" maxlength="120" placeholder="Add Relational Crafting Here"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divIdLock +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divIdLock+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');
				}			
			} else {			 
				if(navigator.userAgent.match(/Android/i)) {
					var $element = $('<div id="' + divId +'" class="col-md-4 after_sketch_task single_task_item_sketch draggableResizable upto_10 common_elem" style="position: absolute; display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top_sketch"><div class="col-md-3 col-xs-3 task task_smaller" style="float: left;" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 energy energy_small" style="float: left;" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/ Energy: </div><div class="col-md-2 col-xs-2 energy_field" style="float: left;"><input type="text" class="form-control-sketch input-sm_custom form-control-sketch_small" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onchange="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent" style="float: left;">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm_sketch"><textarea class="col-md-12 col-xs-12 add_task_area_sketch" style="text-align: center;" id="'+areaId+'" type="text" onkeypress="return disableKey(event)" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onBlur="disableDraggableResizable()"></textarea></div><div class="col-md-12 additional_area"><img class="person_icon" src="../img/personIcon.png"><textarea class="form-control-additional input-sm form-control-additional_smaller" id="'+additionalId+'" onkeypress="return disableKey(event)" name="additional" maxlength="120" placeholder="Add Relational Crafting Here" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divIdLock +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divIdLock+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');					
				}else if(navigator.userAgent.match(/iPhone/i)) {
					var $element = $('<div id="' + divId +'" class="col-md-4 after_sketch_task single_task_item_sketch draggableResizable upto_10 common_elem" style="position: absolute; display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top_sketch"><div class="col-md-3 col-xs-3 task task_smaller" style="float: left;" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 energy energy_small" style="float: left;" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/ Energy: </div><div class="col-md-2 col-xs-2 energy_field" style="float: left;"><input type="text" class="form-control-sketch input-sm_custom form-control-sketch_small" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onchange="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent" style="float: left;">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm_sketch"><textarea class="col-md-12 col-xs-12 add_task_area_sketch" style="text-align: center;" id="'+areaId+'" type="text" onkeypress="return disableKey(event)" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onBlur="disableDraggableResizable()"></textarea></div><div class="col-md-12 additional_area"><img class="person_icon" src="../img/personIcon.png"><textarea class="form-control-additional input-sm form-control-additional_smaller" id="'+additionalId+'" onkeypress="return disableKey(event)" name="additional" maxlength="120" placeholder="Add Relational Crafting Here" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divIdLock +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divIdLock+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');
				}else if(navigator.userAgent.match(/iPad/i)) {
					var $element = $('<div id="' + divId +'" class="col-md-4 after_sketch_task single_task_item_sketch draggableResizable upto_10 common_elem" style="position: absolute; display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top_sketch"><div class="col-md-3 col-xs-3 task task_smaller" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6  energy energy_small" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch form-control-sketch_small input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onchange="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm_sketch"><textarea class="col-md-12 col-xs-12 add_task_area_sketch" style="text-align: center;" id="'+areaId+'" type="text" onkeypress="return disableKey(event)" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onBlur="disableDraggableResizable()"></textarea></div><div class="col-md-12 additional_area"><img class="person_icon" src="../img/personIcon.png"><textarea class="form-control-additional form-control-additional_small input-sm" id="'+additionalId+'" onkeypress="return disableKey(event)" name="additional" maxlength="120" placeholder="Add Relational Crafting Here" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divIdLock +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divIdLock+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');
				} else {
					var $element = $('<div id="' + divId +'" class="col-md-4 after_sketch_task single_task_item_sketch draggableResizable uptoDesc_10 common_elem" style="position: absolute; display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top_sketch"><div class="col-md-3 col-xs-3 task task_smaller">TASK</div><div class="col-md-6  energy energy_small">Time/Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch form-control-sketch_small input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onchange="return getEnergyValue(this)" /></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm_sketch"><textarea class="col-md-12 col-xs-12 add_task_area_sketch" style="text-align: center;" id="'+areaId+'" type="text" onkeypress="return disableKey(event)" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200"></textarea></div><div class="col-md-12 additional_area"><img class="person_icon" src="../img/personIcon.png"><textarea class="form-control-additional form-control-additional_small input-sm" id="'+additionalId+'" onkeypress="return disableKey(event)" name="additional" maxlength="120" placeholder="Add Relational Crafting Here"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divIdLock +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divIdLock+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');
				}	
			}	   	
		} else {
			count = ClientSession.get("total_count");
			var divId = "divImg_" + count;
			var divIdx = "divImg_" + count+"_x";		
			var energyId = "energyId_" + count; 
			var areaId = "areaId_" + count;
			var additionalId = "additionalId_" + count;
			var divIdLock = "divImg_" + count+"_lock";
			if(totalDiv==20){			  
				if(navigator.userAgent.match(/Android/i)) {
					var $element = $('<div id="' + divId +'" class="col-md-4 after_sketch_task single_task_item_sketch draggableResizable upto_9 common_elem" style="position: absolute; display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top_sketch"><div class="col-md-3 col-xs-3 task task_smaller" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6  energy energy_smaller" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/ Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch form-control-sketch_smaller input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onchange="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm_sketch"><textarea class="col-md-12 col-xs-12 add_task_area_sketch" style="text-align: center;" id="'+areaId+'" type="text" onkeypress="return disableKey(event)" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onBlur="disableDraggableResizable()"></textarea></div><div class="col-md-12 additional_area"><img class="person_icon" src="../img/personIcon.png"><textarea class="form-control-additional form-control-additional_smaller input-sm" id="'+additionalId+'" onkeypress="return disableKey(event)" name="additional" maxlength="120" placeholder="Add Relational Crafting Here" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divIdLock +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divIdLock+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');					
				}else if(navigator.userAgent.match(/iPhone/i)) {
					var $element = $('<div id="' + divId +'" class="col-md-4 after_sketch_task single_task_item_sketch draggableResizable upto_9 common_elem" style="position: absolute; display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top_sketch"><div class="col-md-3 col-xs-3 task task_smaller" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6  energy energy_smaller" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/ Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch form-control-sketch_smaller input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onchange="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm_sketch"><textarea class="col-md-12 col-xs-12 add_task_area_sketch" style="text-align: center;" id="'+areaId+'" type="text" onkeypress="return disableKey(event)" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onBlur="disableDraggableResizable()"></textarea></div><div class="col-md-12 additional_area"><img class="person_icon" src="../img/personIcon.png"><textarea class="form-control-additional form-control-additional_smaller input-sm" id="'+additionalId+'" onkeypress="return disableKey(event)" name="additional" maxlength="120" placeholder="Add Relational Crafting Here" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divIdLock +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divIdLock+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');
				}else if(navigator.userAgent.match(/iPad/i)){
					var $element = $('<div id="' + divId +'" class="col-md-4 after_sketch_task single_task_item_sketch draggableResizable upto_9 common_elem" style="position: absolute; display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top_sketch"><div class="col-md-3 col-xs-3 task task_smaller" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6  energy energy_smaller" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch form-control-sketch_smaller input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onchange="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm_sketch"><textarea class="col-md-12 col-xs-12 add_task_area_sketch" style="text-align: center;" id="'+areaId+'" type="text" onkeypress="return disableKey(event)" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onBlur="disableDraggableResizable()"></textarea></div><div class="col-md-12 additional_area"><img class="person_icon" src="../img/personIcon.png"><textarea class="form-control-additional form-control-additional_small input-sm" id="'+additionalId+'" onkeypress="return disableKey(event)" name="additional" maxlength="120" placeholder="Add Relational Crafting Here" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divIdLock +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divIdLock+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');
				}else {
					var $element = $('<div id="' + divId +'" class="col-md-4 after_sketch_task single_task_item_sketch draggableResizable uptoDesc_9 common_elem" style="position: absolute; display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top_sketch"><div class="col-md-3 col-xs-3 task task_smaller">TASK</div><div class="col-md-6  energy energy_smaller">Time/Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch form-control-sketch_smaller input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onchange="return getEnergyValue(this)" /></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm_sketch"><textarea class="col-md-12 col-xs-12 add_task_area_sketch" style="text-align: center;" id="'+areaId+'" type="text" onkeypress="return disableKey(event)" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200"></textarea></div><div class="col-md-12 additional_area"><img class="person_icon" src="../img/personIcon.png"><textarea class="form-control-additional form-control-additional_small input-sm" id="'+additionalId+'" onkeypress="return disableKey(event)" name="additional" maxlength="120" placeholder="Add Relational Crafting Here"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divIdLock +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divIdLock+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');
				}				
			} else {			  
				if(navigator.userAgent.match(/Android/i)) {
					var $element = $('<div id="' + divId +'" class="col-md-4 after_sketch_task single_task_item_sketch draggableResizable upto_10 common_elem" style="position: absolute; display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top_sketch"><div class="col-md-3 col-xs-3 task task_smaller" style="float: left;" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 energy energy_small" style="float: left;" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/ Energy: </div><div class="col-md-2 col-xs-2 energy_field " style="float: left;"><input type="text" class="form-control-sketch input-sm_custom form-control-sketch_small" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onchange="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent" style="float: left;">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm_sketch"><textarea class="col-md-12 col-xs-12 add_task_area_sketch" style="text-align: center;" id="'+areaId+'" type="text" onkeypress="return disableKey(event)" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onBlur="disableDraggableResizable()"></textarea></div><div class="col-md-12 additional_area"><img class="person_icon" src="../img/personIcon.png"><textarea class="form-control-additional input-sm form-control-additional_smaller" id="'+additionalId+'" onkeypress="return disableKey(event)" name="additional" maxlength="120" placeholder="Add Relational Crafting Here" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divIdLock +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divIdLock+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');					
				}else if(navigator.userAgent.match(/iPhone/i)) {
					var $element = $('<div id="' + divId +'" class="col-md-4 after_sketch_task single_task_item_sketch draggableResizable upto_10 common_elem" style="position: absolute; display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top_sketch"><div class="col-md-3 col-xs-3 task task_smaller" style="float: left;" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 energy energy_small" style="float: left;" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/ Energy: </div><div class="col-md-2 col-xs-2 energy_field" style="float: left;"><input type="text" class="form-control-sketch input-sm_custom form-control-sketch_small" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onchange="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent" style="float: left;">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm_sketch"><textarea class="col-md-12 col-xs-12 add_task_area_sketch" style="text-align: center;" id="'+areaId+'" type="text" onkeypress="return disableKey(event)" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onBlur="disableDraggableResizable()"></textarea></div><div class="col-md-12 additional_area"><img class="person_icon" src="../img/personIcon.png"><textarea class="form-control-additional input-sm form-control-additional_smaller" id="'+additionalId+'" onkeypress="return disableKey(event)" name="additional" maxlength="120" placeholder="Add Relational Crafting Here" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divIdLock +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divIdLock+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');
				}else if(navigator.userAgent.match(/iPad/i)) {
					var $element = $('<div id="' + divId +'" class="col-md-4 after_sketch_task single_task_item_sketch draggableResizable upto_10 common_elem" style="position: absolute; display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top_sketch"><div class="col-md-3 col-xs-3 task" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6  energy energy_small" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch form-control-sketch_small input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onchange="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm_sketch"><textarea class="col-md-12 col-xs-12 add_task_area_sketch" style="text-align: center;" id="'+areaId+'" type="text" onkeypress="return disableKey(event)" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onBlur="disableDraggableResizable()"></textarea></div><div class="col-md-12 additional_area"><img class="person_icon" src="../img/personIcon.png"><textarea class="form-control-additional form-control-additional_small input-sm" id="'+additionalId+'" onkeypress="return disableKey(event)" name="additional" maxlength="120" placeholder="Add Relational Crafting Here" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divIdLock +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divIdLock+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');
				} else {
					var $element = $('<div id="' + divId +'" class="col-md-4 after_sketch_task single_task_item_sketch draggableResizable uptoDesc_10 common_elem" style="position: absolute; display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top_sketch"><div class="col-md-3 col-xs-3 task task_smaller">TASK</div><div class="col-md-6  energy energy_small">Time/Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch form-control-sketch_small input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onchange="return getEnergyValue(this)" /></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm_sketch"><textarea class="col-md-12 col-xs-12 add_task_area_sketch" style="text-align: center;" id="'+areaId+'" type="text" onkeypress="return disableKey(event)" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200"></textarea></div><div class="col-md-12 additional_area"><img class="person_icon" src="../img/personIcon.png"><textarea class="form-control-additional form-control-additional_small input-sm" id="'+additionalId+'" onkeypress="return disableKey(event)" name="additional" maxlength="120" placeholder="Add Relational Crafting Here"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divIdLock +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divIdLock+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');
				}			
			}		
		}				
	  	count++;	  	
	  	ClientSession.set("total_count", count);	  	
	  	calculateEnergyValue(energyValue, status, null, "ADD");
	  	
	  	/************** To remove the focus to the new old task *************/
	  	removeFocus();
	  	
	    $("#outderDiv").append($element);
	    
	    addTaskPosition(energyStyle, divId);
	    $element.draggable();	    
	    $element.resizable({			
			//ghost: 'true',
			handles: 'se',
			aspectRatio: 2 / 2,
	    });
	    allCanvasArray.push($element);
	    $(".single_task_item_sketch").css({ 'z-index':'4'}); 
	    $('.single_task_item_sketch').draggable({
			cursor: 'move',       
			containment: '#outderDiv'    
		});  
	    fixDesignAS();
	    /***************** ISSUE# 207***************/
	    $("#" + divId).children(".ui-resizable-handle").css({'bottom' : '-15px'});
	    
	    /********* Function add to make the time/ energy field and task description area editable ******/
	    contentEditable();
	    /***************** To set the focus to the new Task ***************/
	    $("#" + divId).children().find(".add_task_area_sketch").addClass("add_task_area_sketch_focus");
	    $("#" + areaId).focus();
	    
	    /***************** To Disable the paste functionality ISSUE# 186***************/
	     $('#' +areaId).bind("paste",function(e) {
	        e.preventDefault();
	     });
	     
	    $('input[type=text], textarea').placeholder();
	}
		   
    /**************** Store the Json Object**********************/
    generateJsonObj();
        
    fixDivLength();
    /**************** to get the resized width value**********************/
  
    $(".single_task_item_sketch").resizable({
		stop:function(event,ui){			
			var divId = $(this).attr('id');		
	    	var finalWidth = $(this).width();   
	    	calculateEnergy(finalWidth, divId);		
		}
    });	
    $('.after_sketch_task').draggable({
        cursor: 'move',        // sets the cursor apperance
        containment: '#outderDiv'    // sets to can be dragged only within "#outderDiv" element
      });
    $("#outderDiv").css({'height':($("#pageWrapCopy").height()+'px')});
    /**************** To reduce the Z-index value to fix the UI**********************/
    $(".single_task_item_sketch .ui-resizable-se").css({ 'z-index':'0'});
	}
	if (windowSize <= 1024) {   		 	
	 	$(".after_sketch_task").on('touchstart', function(e){
    		enableDraggableResizable($(this).attr('id'));
	 	});
	}
	enableDraggableResizableOval();
	toogleFocus();
return false;
});



/************************** To Draw Strength Value Passion***************************/
var afterSktchJson = JSON.parse(ClientSession.get("total_json_obj"));
	if(afterSktchJson!=null){
		$('#strength_pageWrap').empty();
		$('#value_pageWrap').empty();
		$('#passion_pageWrap').empty();
		var cnt = afterSktchJson.length;
		for (var i = 0; i<cnt;i++){
		var divIda = afterSktchJson[i]["divId"];
		var textIda = afterSktchJson[i]["textId"];	
		var divValuea = afterSktchJson[i]["divValue"];		
		var positionLefta = afterSktchJson[i]["positionLeft"];
		var positionTopa = afterSktchJson[i]["positionTop"];	
		var divStyle = afterSktchJson[i]["style"];
		if((divStyle == null) || (divStyle == "null") || (divStyle.search('block') != -1)){
			divStyle = "block;";
		}else{
			divStyle = "none;";
		}
	if(divIda.search("divStrength_") != -1) {
	
	if(navigator.userAgent.match(/Android/i) || navigator.userAgent.match(/iPhone/i)) {
			var $elementNew1 = $('<div id="' + divIda +'" class=" commonDrop plus_sign_small" style="position: relative; display:'+divStyle+';"><img src="../img/strength.png" width="50" height="50" style="no-repeat 0 9px;"/><div class="removeService delete_element idGenClass" onclick=deleteElement(this)><img src="../img/cross-black.png" alt="Delete" /></div><div class="plus_sign_small_span" id="'+textIda+'"><p>'+divValuea+'</p></div>');
		} else if(navigator.userAgent.match(/iPad/i) || navigator.userAgent.match(/IPad/i)) {
			var $elementNew1 = $('<div id="' + divIda +'" class=" commonDrop plus_sign_small" style="position: relative; display:'+divStyle+';"><img src="../img/strength.png" width="80" height="80" style="no-repeat 0 9px;"/><div class="removeService delete_element idGenClass" onclick=deleteElement(this)><img src="../img/cross-black.png" alt="Delete" /></div><div class="plus_sign_small_span" id="'+textIda+'"><p>'+divValuea+'</p></div>');
		} else {
			var $elementNew1 = $('<div id="' + divIda +'" class=" commonDrop plus_sign_small" style="position: relative; display:'+divStyle+';"><img src="../img/strength.png" width="100" height="80" style="no-repeat 0 9px;"/><div class="removeService delete_element idGenClass" onclick=deleteElement(this)><img src="../img/cross-black.png" alt="Delete" /></div><div class="plus_sign_small_span" id="'+textIda+'"><p>'+divValuea+'</p></div>');
		}			
		$("#strength_pageWrap").append($elementNew1);
		$elementNew1.draggable();
	} else if(divIda.search("divValue_") != -1){	
		if(navigator.userAgent.match(/Android/i) || navigator.userAgent.match(/iPhone/i)) {
				var $elementNew2 = $('<div id="' + divIda +'" class="oval_small commonDrop" style="position: relative; display:'+divStyle+'"><img src="../img/value.png" width="55" height="45" style="no-repeat 0 9px;"/><div class="removeService delete_element idGenClass" onclick=deleteElement(this)><img src="../img/cross-black.png" alt="Delete" /></div><div class="oval_small_span" id="'+textIda+'"><p>'+divValuea+'</p></div>');
			} else if(navigator.userAgent.match(/iPad/i) || navigator.userAgent.match(/IPad/i)) {
				var $elementNew2 = $('<div id="' + divIda +'" class="oval_small commonDrop" style="position: relative; display:'+divStyle+'"><img src="../img/value.png" width="85" height="76" style="no-repeat 0 9px;"/><div class="removeService delete_element idGenClass" onclick=deleteElement(this)><img src="../img/cross-black.png" alt="Delete" /></div><div class="oval_small_span" id="'+textIda+'"><p>'+divValuea+'</p></div>');
			} else {
				var $elementNew2 = $('<div id="' + divIda +'" class="oval_small commonDrop" style="position: relative; display:'+divStyle+'"><img src="../img/value.png" width="105" height="85" style="no-repeat 0 9px;"/><div class="removeService delete_element idGenClass" onclick=deleteElement(this)><img src="../img/cross-black.png" alt="Delete" /></div><div class="oval_small_span" id="'+textIda+'"><p>'+divValuea+'</p></div>');
			}			
		$("#value_pageWrap").append($elementNew2);
		$elementNew2.draggable();
	} else if(divIda.search("divPassion_") != -1){
		if(navigator.userAgent.match(/Android/i) || navigator.userAgent.match(/iPhone/i)) {
				var $elementNew3 = $('<div id="' + divIda +'" class="up_triangle_small commonDrop" style="position: relative; display:'+divStyle+'"><img src="../img/passion.png" width="50" height="50" style="no-repeat 0 9px;"/><div class="removeService delete_element idGenClass" onclick=deleteElement(this)><img src="../img/cross-black.png" alt="Delete" /></div><div class="up_triangle_small_span" id="'+textIda+'"><p>'+divValuea+'</p></div>');
			} else if(navigator.userAgent.match(/iPad/i) || navigator.userAgent.match(/IPad/i)) {
				var $elementNew3 = $('<div id="' + divIda +'" class="up_triangle_small commonDrop" style="position: relative; display:'+divStyle+'"><img src="../img/passion.png" width="87" height="82" style="no-repeat 0 9px;"/><div class="removeService delete_element idGenClass" onclick=deleteElement(this)><img src="../img/cross-black.png" alt="Delete" /></div><div class="up_triangle_small_span" id="'+textIda+'"><p>'+divValuea+'</p></div>');
			} else {
				var $elementNew3 = $('<div id="' + divIda +'" class="up_triangle_small commonDrop" style="position: relative; display:'+divStyle+'"><img src="../img/passion.png" width="100" height="85" style="no-repeat 0 9px;"/><div class="removeService delete_element idGenClass" onclick=deleteElement(this)><img src="../img/cross-black.png" alt="Delete" /></div><div class="up_triangle_small_span" id="'+textIda+'"><p>'+divValuea+'</p></div>');
			}
		$("#passion_pageWrap").append($elementNew3);
		$elementNew3.draggable();
	}
	}  
	}
	
	/********  added for drag and drop functionality ***********/			
	$(document).ready(function () {
		var countOval = 0;
		 var countStrength = 0;
		 var countValue = 0;
		 var countPassion = 0;
	    $(".commonDrop ").draggable({
	        helper: 'clone',
			cursor: 'move',
			tolerance: 'fit',
			appendTo: '#outderDiv'
	    });	
	    if (!sessionStorage.getItem("bsEdit")) {
		$(".oval_draggable").draggable({
	        helper: 'clone',
			cursor: 'move',
			tolerance: 'fit',
			appendTo: '#outderDiv'
	    });		
	    }
				
	    $(".pageWrap_after_sketch").droppable({
	        accept: ".oval_draggable, .commonDrop",	     	       
	        drop: function(event, ui) {	  	      
	        	if (ui.helper.attr('class').search("oval_draggable") != -1){
	        		if ($('#outderDiv').children('.oval_cloned_div:visible').length == 3 ) { 
	        			alertify.alert("<img src='../img/alert-icon.png'><br /><p>You can have a maximum of 3 role frames.</p>");
	        			return false;	        			
	        		} else {
	        			if(ui.helper.attr('id') == undefined){       		
		        		var $clone = ui.helper.clone();
			            if (!$clone.is('.copied')) {
			                $(this).append($clone.addClass('copied').draggable({	                	
			                    containment: '.pageWrap_after_sketch'
			                }));
			            	 $(this).append($clone.addClass('copied'));
							//var elementStyle = $clone.attr('class');
			               // $("#outderDiv .speech-bubble").show();
							 $("#outderDiv .copied .speech-bubble").show();	
			                if($('#outderDiv').children('.oval_cloned_div').length ==0){
			                	countOval = 1;			                	
			                } else {
			                	countOval = $('#outderDiv').children('.oval_cloned_div').length+1;
			                }
			                $(this).append($(".copied").attr("id","cloneOval_"+countOval));
			                
//			              /  $(".idGenClass").attr("id","cloneOval_"+countOval+"_x");	
			                
			                /** ADDED FOR NEW ROLE FRAME **/
			                $(".copied .speech-bubble textarea").attr("id","roleFrame_"+countOval);	
			                /**** ENDED *****/
			                
			                $(".copied").removeClass('oval_draggable').addClass('oval_cloned_div');			               
			               // $(".copied").addClass('commonDelete');
			                $(".copied").children().removeClass('idGenClass');
			                $(".copied").addClass('common_elem');
			               // $(this).css("z-index",3);
			                $(".oval_cloned_div").css("z-index", 2);
			                if($('.copied').hasClass('horizontal_oval')){
			                	$(".copied").addClass('horizontal_oval_active');
			                	$(".copied").removeClass('horizontal_oval');
			                	fixPositionWhileDrop(this, countOval);
			                } 
			                $(".copied").removeClass('copied');		
			                var cloneId = document.getElementById("cloneOval_"+countOval).id;			                	              
			                $clone.resizable({
			                	
			                });
			                $("#" + cloneId).css({position: "absolute"});							
			                /** ADDED FOR NEW ROLE FRAME **/
			                enableSpeechBubble(cloneId);								
							//$("#" + cloneId).attr("onmouseup","reDragRoleFrame('"+cloneId+"')");
			                
			               /* $('.oval_cloned_div').draggable({
			                	  stop: function(event, ui) {
			                		  var cloneId = $(this).attr('id');	
			                		  reDragRoleFrame(cloneId);
			                	  }
			                });*/
			             
			                /*$(this).append($(".copied").attr("id","cloneOval_"+countOval));*/
			               var roleId = "role_" +cloneId;
			              /*  $('.oval_cloned_div').draggable({*/
			               	 $('#'+cloneId).draggable({	
			                	start: function(){
			                        if($(this).hasClass('oval_cloned_div')){
			                            var p = $(this).position();
			                            $(this).data('lastLeft',p.left);
			                            $(this).data('lastTop',p.top);
			                        }
			                    },
			                    stop: function(){
			                        if($(this).hasClass('oval_cloned_div')){
			                            $(this).removeData('lastLeft');
			                            $(this).removeData('lastTop');
			                        }
			                        reDragRoleFrame(cloneId);
			                    },
			                    drag: function(event, ui) {
			                        if($(this).hasClass('oval_cloned_div')){
			                            var p = $(this).position();
			                            $(this).data('lastLeft',p.left);
			                            $(this).data('lastTop',p.top);

			                            var dx = ui.position.left - $(this).data('lastLeft');
			                            var dy = ui.position.top - $(this).data('lastTop');
			                            $('#'+roleId).each(function(){
			                                var p = $(this).position();
			                                $(this).css({
			                                    left: (p.left + dx) + "px",
			                                    top: (p.top + dy) + "px"
			                                });
			                            });
			                        }
			                    }
			                });
							/**** ENDED *****/						
							$( "#" + cloneId ).resize(function() {
								checkResizeCordinates(cloneId);
							});
							toogleFocus();							
			            }
		        	}	
	        	}
	        	} else if (ui.helper.attr('class').search("plus_sign_small") != -1){
	        		if(sessionStorage.getItem("bsEdit")){
	        			alertify.alert("<img src='../img/alert-icon.png'><br /><p>Cannot add strength as you have already completed your after diagram.</p>");
	        			return false;
	        		} else if ($('#strength_pageWrap').children('.plus_sign_small:visible').length == 2 
	        				&& $('#outderDiv').children('.plus_sign_small:visible').length == 7 ) { 
	        			alertify.alert("<img src='../img/alert-icon.png'><br /><p>You can have a maximum of 3 same strengths.</p>");
	        			return false;	        			
	        		} else if ($('#strength_pageWrap').children('.plus_sign_small:visible').length == 3 
	        				&& $('#outderDiv').children('.plus_sign_small:visible').length == 10 ) { 
	        			alertify.alert("<img src='../img/alert-icon.png'><br /><p>You can have a maximum of 3 same strengths.</p>");
	        			return false;	        			
	        		} else if ($('#strength_pageWrap').children('.plus_sign_small:visible').length == 4 
	        				&& $('#outderDiv').children('.plus_sign_small:visible').length == 13 ) { 
	        			alertify.alert("<img src='../img/alert-icon.png'><br /><p>You can have a maximum of 3 same strengths.</p>");
	        			return false;	        			
	        		} else {
	        			if(ui.helper.attr('id') == undefined){	        		
			        		var $clone = ui.helper.clone();
				            if (!$clone.is('.copied')) {
				                $(this).append($clone.addClass('copied').draggable({	                	
				                    containment: '.pageWrap_after_sketch'			                    	 
				                })); 			                
				                countStrength = $('#outderDiv').children('.plus_sign_small').length;
								countStrength = countStrength - 1;
				                $(this).append($(".copied").attr("id","divStrengthClone_"+countStrength));
				                $(".idGenClass").attr("id","divStrengthClone_"+countStrength+"_x");	
				                $(".copied").addClass('commonDrag');
				                $(".copied").addClass('commonDelete');
				                $(".copied").children().removeClass('idGenClass');
				                $(".copied").addClass('common_elem');
				                $(".copied").removeClass('copied');
				                var leftPositionPx = document.getElementById("divStrengthClone_" + countStrength).style.left;
				                var topPositionPx = document.getElementById("divStrengthClone_" + countStrength).style.top;
				                var leftPosition = parseInt(leftPositionPx.replace("px",''));
				                var topPosition = parseInt(topPositionPx.replace("px",''));
				                var outerDivHeight = $('#outderDiv').height();				               
				                if(leftPosition<=0) {
				                	document.getElementById("divStrengthClone_" + countStrength).style.left = "10px";
				                } else if(leftPosition>=923) {
				                	document.getElementById("divStrengthClone_" + countStrength).style.left = "920px";
				                }
				                if(topPosition>=(outerDivHeight-100)) {
				                	document.getElementById("divStrengthClone_" + countStrength).style.top = (outerDivHeight - 100)+"px";
				                } else if(topPosition<0) {
				                	document.getElementById("divStrengthClone_" + countStrength).style.top = "0px";
				                }
				                var existingValue = "";
			        			var value = "";
			        			var count = 0;
			        			var outerLoop = $('#outderDiv').children('.plus_sign_small').length;
			        			for(var i = 1; i<outerLoop; i++){
			        				var style = document.getElementById("divStrengthClone_" +i).getAttribute("style");
			        				if((style.search('block') != -1)){
			        					value = $("#divStrengthClone_" + i + "  p").text();
					        			existingValue = $("#divStrengthClone_" + countStrength + "  p").text();
					        				if( value == existingValue) {
					        				  count++;
					        				}				
			        					}
			    				}
			        			if(count >= 4){			        				
			        				$("#divStrengthClone_" + countStrength).css("display", "none");
			        				alertify.alert("<img src='../img/alert-icon.png'><br /><p>You can have a maximum of 3 strengths.</p>");
			        				count == 0;
			        			}
				            }
			        	}	
	        		}
	        		
	        	}
	        	else if (ui.helper.attr('class').search("oval_small") != -1){
	        		if(sessionStorage.getItem("bsEdit")){
	        			alertify.alert("<img src='../img/alert-icon.png'><br /><p>Cannot add value as you have already completed your after diagram.</p>");
	        			return false;
	        		} else if ($('#value_pageWrap').children('.oval_small:visible').length == 2 
	        				&& $('#outderDiv').children('.oval_small:visible').length == 7 ) { 
	        			alertify.alert("<img src='../img/alert-icon.png'><br /><p>You can have a maximum of 3 same values.</p>");
	        			return false;	        			
	        		} else if ($('#value_pageWrap').children('.oval_small:visible').length == 3 
	        				&& $('#outderDiv').children('.oval_small:visible').length == 10 ) { 
	        			alertify.alert("<img src='../img/alert-icon.png'><br /><p>You can have a maximum of 3 same values.</p>");
	        			return false;	        			
	        		} else if ($('#value_pageWrap').children('.oval_small:visible').length == 4 
	        				&& $('#outderDiv').children('.oval_small:visible').length == 13 ) { 
	        			alertify.alert("<img src='../img/alert-icon.png'><br /><p>You can have a maximum of 3 same values.</p>");
	        			return false;	        			
	        		} else {
	        			if(ui.helper.attr('id') == undefined){	        		
		        		var $clone = ui.helper.clone();
			            if (!$clone.is('.copied')) {
			                $(this).append($clone.addClass('copied').draggable({	                	
			                    containment: '.pageWrap_after_sketch'
			                }));   			                
			                countValue = $('#outderDiv').children('.oval_small').length;
							countValue = countValue - 1;
			                $(this).append($(".copied").attr("id","divValueClone_"+countValue));
			                $(".idGenClass").attr("id","divValueClone_"+countValue+"_x");	
			                $(".copied").addClass('commonDrag');
			                $(".copied").addClass('commonDelete');
			                $(".copied").children().removeClass('idGenClass');
			                $(".copied").addClass('common_elem');
			                $(".copied").removeClass('copied');
			                var leftPositionPx = document.getElementById("divValueClone_" + countValue).style.left;
			                var topPositionPx = document.getElementById("divValueClone_" + countValue).style.top;
			                var leftPosition = parseInt(leftPositionPx.replace("px",''));
			                var topPosition = parseInt(topPositionPx.replace("px",''));
			                var outerDivHeight = $('#outderDiv').height();	
			                if(leftPosition<=0) {
			                	document.getElementById("divValueClone_" + countValue).style.left = "10px";
			                } else if(leftPosition>=921) {
			                	document.getElementById("divValueClone_" + countValue).style.left = "920px";
			                }
			                if(topPosition>=(outerDivHeight-100)) {
			                	document.getElementById("divValueClone_" + countValue).style.top = (outerDivHeight - 100)+"px";
			                } else if(topPosition<0) {
			                	document.getElementById("divValueClone_" + countValue).style.top = "0px";
			                }		                
			                var existingValue = "";
		        			var value = "";
		        			var count = 0;
		        			var outerLoop = $('#outderDiv').children('.oval_small').length;
		        			for(var i = 1; i<outerLoop; i++){
		        				var style = document.getElementById("divValueClone_" +i).getAttribute("style");
		        				if((style.search('block') != -1)){
		        					value = $("#divValueClone_" + i + "  p").text();
				        			existingValue = $("#divValueClone_" + countValue + "  p").text();				
				        				if( value == existingValue) {
				        				  count++;
				        				}				
		        					}
		    				}
		        			if(count >= 4){			        				
		        				$("#divValueClone_" + countValue).css("display", "none");
		        				alertify.alert("<img src='../img/alert-icon.png'><br /><p>You can have a maximum of 3 values.</p>");
		        				count == 0;
		        			}
			            }
	        			}
		        	}	
	        	} else if (ui.helper.attr('class').search("up_triangle_small") != -1){
	        		if(sessionStorage.getItem("bsEdit")){
	        			alertify.alert("<img src='../img/alert-icon.png'><br /><p>Cannot add passion as you have already completed your after diagram.</p>");
	        			return false;
	        		} else if ($('#passion_pageWrap').children('.up_triangle_small:visible').length == 2 
	        				&& $('#outderDiv').children('.up_triangle_small:visible').length == 7 ) { 
	        			alertify.alert("<img src='../img/alert-icon.png'><br /><p>You can have a maximum of 3 same passions.</p>");
	        			return false;	        			
	        		} else if ($('#passion_pageWrap').children('.up_triangle_small:visible').length == 3 
	        				&& $('#outderDiv').children('.up_triangle_small:visible').length == 10 ) { 
	        			alertify.alert("<img src='../img/alert-icon.png'><br /><p>You can have a maximum of 3 same passions.</p>");
	        			return false;	        			
	        		} else if ($('#passion_pageWrap').children('.up_triangle_small:visible').length == 4 
	        				&& $('#outderDiv').children('.up_triangle_small:visible').length == 13 ) { 
	        			alertify.alert("<img src='../img/alert-icon.png'><br /><p>You can have a maximum of 3 same passions.</p>");
	        			return false;	        			
	        		} else {
	        		if(ui.helper.attr('id') == undefined){	        		
		        		var $clone = ui.helper.clone();
			            if (!$clone.is('.copied')) {
			                $(this).append($clone.addClass('copied').draggable({	                	
			                    containment: '.pageWrap_after_sketch'
			                })); 			                
			                countPassion = $('#outderDiv').children('.up_triangle_small').length;
							countPassion = countPassion - 1;
			                $(this).append($(".copied").attr("id","divPassionClone_"+countPassion));
			                $(".idGenClass").attr("id","divPassionClone_"+countPassion+"_x");	
			                $(".copied").addClass('commonDrag');
			                $(".copied").addClass('commonDelete');
			                $(".copied").addClass('common_elem');
			                $(".copied").children().removeClass('idGenClass');
			                $(".copied").removeClass('copied');
			                var leftPositionPx = document.getElementById("divPassionClone_" + countPassion).style.left;
			                var topPositionPx = document.getElementById("divPassionClone_" + countPassion).style.top;
			                var leftPosition = parseInt(leftPositionPx.replace("px",''));
			                var topPosition = parseInt(topPositionPx.replace("px",''));
			                var outerDivHeight = $('#outderDiv').height();
			                if(leftPosition<=0) {
			                	document.getElementById("divPassionClone_" + countPassion).style.left = "10px";
			                } else if(leftPosition>=923) {
			                	document.getElementById("divPassionClone_" + countPassion).style.left = "920px";
			                }
			                if(topPosition>=(outerDivHeight-100)) {
			                	document.getElementById("divPassionClone_" + countPassion).style.top = (outerDivHeight - 100)+"px";
			                } else if(topPosition<0) {
			                	document.getElementById("divPassionClone_" + countPassion).style.top = "0px";
			                }
			                var existingValue = "";
		        			var value = "";
		        			var count = 0;
		        			var outerLoop = $('#outderDiv').children('.up_triangle_small').length;
		        			for(var i = 1; i<outerLoop; i++){
		        				var style = document.getElementById("divPassionClone_" +i).getAttribute("style");
		        				if((style.search('block') != -1)){
		        					value = $("#divPassionClone_" + i + "  p").text();
				        			existingValue = $("#divPassionClone_" + countPassion + "  p").text();
				        				if( value == existingValue) {
				        				  count++;
				        				}				
		        					}
		    				}
		        			if(count >= 4){			        				
		        				$("#divPassionClone_" + countPassion).css("display", "none");
		        				alertify.alert("<img src='../img/alert-icon.png'><br /><p>You can have a maximum of 3 passions.</p>");
		        				count == 0;
		        			}
			            }
	        		}
		        	}	
	        	}
	        	        		            
	        }		   	
	    });	    
	});
	  		
	


$(document).ready(function() {		
	//if(sessionStorage.getItem("isCompleted") == 0){
		if (!sessionStorage.getItem("bsEdit") || (null == sessionStorage.getItem("bsEdit"))) {
			/*$('.oval_cloned_div').draggable({
				  stop: function(event, ui) {
					  var cloneId = $(this).attr('id');	
					  reDragRoleFrame(cloneId);
				  }
			});*/
			/*  if($('#outderDiv').children('.oval_cloned_div').length ==0){
              	countOval = 1;			                	
              } else {
              	countOval = $('#outderDiv').children('.oval_cloned_div').length+1;
              }
			var cloneId = "cloneOval_"+countOval;	*/
			$('.oval_cloned_div').draggable({				
            	start: function(){
                    if($(this).hasClass('oval_cloned_div')){
                        var p = $(this).position();
                        $(this).data('lastLeft',p.left);
                        $(this).data('lastTop',p.top);
                    }
                },
                stop: function(){
                    if($(this).hasClass('oval_cloned_div')){
                        $(this).removeData('lastLeft');
                        $(this).removeData('lastTop');
                    }
                   var cloneId = $(this).attr('id');
                    reDragRoleFrame(cloneId);
                },
                drag: function(event, ui) {
                    if($(this).hasClass('oval_cloned_div')){
                        var p = $(this).position();
                        $(this).data('lastLeft',p.left);
                        $(this).data('lastTop',p.top);

                        var dx = ui.position.left - $(this).data('lastLeft');
                        var dy = ui.position.top - $(this).data('lastTop');
                        var cloneId = $(this).attr('id');
                        var roleId = "role_" +cloneId;
                        $('#'+roleId).each(function(){
                            var p = $(this).position();
                            $(this).css({
                                left: (p.left + dx) + "px",
                                top: (p.top + dy) + "px"
                            });
                        });
                    }
                }
            });
		}
	//}
});

	
/**
 * Function to store the json string  
 * while navigate to next page
 * @param null
 */
function goToNextPageAftSktch() {
	
	if (!sessionStorage.getItem("silentSave")) {
		$(".loader_bg").fadeIn();
		$(".loader").fadeIn();
	}
	
	getActualPosition();		
	nextPageAfterSketch = 1;
	ClientSession.set("Next_page_after_sketch", nextPageAfterSketch);	
	var cntTask = 1;
	var totalJsonAftSktch = [];
	var idx = 0;
	var divValueTask = document.getElementById("divImg_" + cntTask);
	
	var mainDivWidth = $("#outderDiv").width();
	var mainDivHeight = $("#outderDiv").height();	
    
	var gcd=calc(mainDivWidth,mainDivHeight);
  	var r1=mainDivWidth/gcd;
    var r2=mainDivHeight/gcd;
    var ratio=r1+":"+r2;
    
	while(divValueTask != null) {
		var unitJson = {};					
		unitJson["divId"] = document.getElementById("divImg_" + cntTask).id;
		unitJson["divIdx"] = document.getElementById("divImg_" + cntTask +"_x").id;
		unitJson["energyId"] = document.getElementById("energyId_"+ cntTask).id;
		unitJson["energyValue"] = document.getElementById("energyId_"+ cntTask).value;
		unitJson["style"] = document.getElementById("divImg_" + cntTask).getAttribute("style");
		unitJson["areaId"] = document.getElementById("areaId_"+ cntTask).id;
		unitJson["areaValue"] = document.getElementById("areaId_"+ cntTask).value.replace(/,/g, "`").replace(/#/g, "-_-").replace(/\?/g, ";_;");
		unitJson["additionalId"] = document.getElementById("additionalId_"+ cntTask).id;
		unitJson["additionalValue"] = document.getElementById("additionalId_"+ cntTask).value.replace(/,/g, "`").replace(/#/g, "-_-").replace(/\?/g, ";_;");	
		unitJson["positionLeft"] = document.getElementById("divImg_" + cntTask).style.left;
		unitJson["positionTop"] = document.getElementById("divImg_" + cntTask).style.top;
		unitJson["positionWidth"] = document.getElementById("divImg_" + cntTask).style.width;
		unitJson["positionHeight"] = document.getElementById("divImg_" + cntTask).style.height;
		
		var positionLeftPixel = document.getElementById("divImg_" + cntTask).style.left;	
		var positionLeft = parseInt(positionLeftPixel.replace("px", ''));
		unitJson["positionLeftPixel"] = positionLeftPixel;		
		unitJson["positionLeftPercent"] = getPercentageValue(mainDivWidth, positionLeft);
		var positionTopPixel = document.getElementById("divImg_" + cntTask).style.top;	
		var positionTop = parseInt(positionTopPixel.replace("px", ''));
		unitJson["positionTopPixel"] = positionTopPixel;
		unitJson["positionTopPercent"] = getPercentageValue(mainDivHeight, positionTop);
		var positionWidthPixel = document.getElementById("divImg_" + cntTask).style.width;	
		var positionWidth = parseInt(positionWidthPixel.replace("px", ''));
		unitJson["positionWidthPercent"] = getPercentageValue(mainDivWidth, positionWidth);
		unitJson["mainDivWidth"] = mainDivWidth;
		unitJson["mainDivHeight"] = mainDivHeight;
		unitJson["aspectRatio"] = ratio;
		
		/***** Added for lock and unlock feature *****/
		unitJson["divLockId"] = document.getElementById("divImg_" + cntTask +"_lock").id;
		unitJson["lockImage"] = document.getElementById("divImg_" + cntTask+ "_lock").getElementsByTagName("img")[0].getAttribute("src");
		
		totalJsonAftSktch[idx++] = unitJson;
		divValueTask = document.getElementById("divImg_" + ++cntTask);
	}
	
	
	var cntStrength = 1;
	var divStrength = document.getElementById("divStrengthClone_" + cntStrength);
	while(divStrength != null) {
			var unitJson = {};						
			unitJson["divId"] = document.getElementById("divStrengthClone_" + cntStrength).id;	
			var divStrengthId = unitJson["divId"];		
			var divStrengthVal = $("#" + divStrengthId + "  p").text();
			unitJson["divValue"] = divStrengthVal;		
			unitJson["style"] = document.getElementById("divStrengthClone_" + cntStrength).getAttribute("style");
			unitJson["positionLeft"] = document.getElementById("divStrengthClone_" + cntStrength).style.left;
			unitJson["positionTop"] = document.getElementById("divStrengthClone_" + cntStrength).style.top;	
			
			var positionLeftPixel = document.getElementById("divStrengthClone_" + cntStrength).style.left;	
			var positionLeft = parseInt(positionLeftPixel.replace("px", ''));
			unitJson["positionLeftPercent"] = getPercentageValue(mainDivWidth, positionLeft);
			var positionTopPixel = document.getElementById("divStrengthClone_" + cntStrength).style.top;	
			var positionTop = parseInt(positionTopPixel.replace("px", ''));
			unitJson["positionTopPercent"] = getPercentageValue(mainDivHeight, positionTop);
			unitJson["aspectRatio"] = ratio;
			totalJsonAftSktch[idx++] = unitJson;		
			divStrength = document.getElementById("divStrengthClone_" + ++cntStrength);		
	}	
	
	var cntValue = 1;
	var divValue = document.getElementById("divValueClone_" + cntValue);
	while(divValue != null) {
		var unitJson = {};		
		unitJson["divId"] = document.getElementById("divValueClone_" + cntValue).id;	
		var divValueId = unitJson["divId"];		
		var divValueVal = $("#" + divValueId + "  p").text();
		unitJson["divValue"] = divValueVal;	
		unitJson["style"] = document.getElementById("divValueClone_" + cntValue).getAttribute("style");		
		unitJson["positionLeft"] = document.getElementById("divValueClone_" + cntValue).style.left;
		unitJson["positionTop"] = document.getElementById("divValueClone_" + cntValue).style.top;	
		
		var positionLeftPixel = document.getElementById("divValueClone_" + cntValue).style.left;	
		var positionLeft = parseInt(positionLeftPixel.replace("px", ''));
		unitJson["positionLeftPercent"] = getPercentageValue(mainDivWidth, positionLeft);
		var positionTopPixel = document.getElementById("divValueClone_" + cntValue).style.top;	
		var positionTop = parseInt(positionTopPixel.replace("px", ''));
		unitJson["positionTopPercent"] = getPercentageValue(mainDivHeight, positionTop);
		unitJson["aspectRatio"] = ratio;
		totalJsonAftSktch[idx++] = unitJson;
		divValue = document.getElementById("divValueClone_" + ++cntValue);
	}

	var cntPassion = 1;
	var divPassion = document.getElementById("divPassionClone_" + cntPassion);
	while(divPassion != null) {
		var unitJson = {};		
		
		unitJson["divId"] = document.getElementById("divPassionClone_" + cntPassion).id;
		var divPassionId = unitJson["divId"];		
		var divPassionVal = $("#" + divPassionId + "  p").text();
		unitJson["divValue"] = divPassionVal;	
		unitJson["style"] = document.getElementById("divPassionClone_" + cntPassion).getAttribute("style");			
		unitJson["positionLeft"] = document.getElementById("divPassionClone_" + cntPassion).style.left;
		unitJson["positionTop"] = document.getElementById("divPassionClone_" + cntPassion).style.top;		
		
		var positionLeftPixel = document.getElementById("divPassionClone_" + cntPassion).style.left;	
		var positionLeft = parseInt(positionLeftPixel.replace("px", ''));
		unitJson["positionLeftPercent"] = getPercentageValue(mainDivWidth, positionLeft);
		var positionTopPixel = document.getElementById("divPassionClone_" + cntPassion).style.top;	
		var positionTop = parseInt(positionTopPixel.replace("px", ''));
		unitJson["positionTopPercent"] = getPercentageValue(mainDivHeight, positionTop);
		unitJson["aspectRatio"] = ratio;
		totalJsonAftSktch[idx++] = unitJson;
		divPassion = document.getElementById("divPassionClone_" + ++cntPassion);
	}
	
	var cntClone = 1;
	var divClone = document.getElementById("cloneOval_" + cntClone);
	while(divClone != null) {
		var unitJson = {};	
		
		unitJson["divId"] = document.getElementById("cloneOval_" + cntClone).id;
		var cloneId = unitJson["divId"];		
		//var cloneValue = $("#" + cloneId + " .speech-bubble textarea").val().replace(/#/g, "-_-").replace(/\?/g, ";_;");
		var cloneValue = $("#role_" + cloneId + " textarea").val().replace(/#/g, "-_-").replace(/\?/g, ";_;");
		unitJson["roleValue"] = cloneValue;			
		unitJson["positionLeft"] = document.getElementById("cloneOval_" + cntClone).style.left;
		unitJson["positionTop"] = document.getElementById("cloneOval_" + cntClone).style.top;	
		unitJson["positionWidth"] = document.getElementById("cloneOval_" + cntClone).style.width;
		unitJson["positionHeight"] = document.getElementById("cloneOval_" + cntClone).style.height;
		var position = $( "#cloneOval_" + cntClone ).position();
		unitJson["positionLeft"] = position.left;
		unitJson["positionTop"] = position.top;	
		unitJson["style"] = document.getElementById("cloneOval_" + cntClone).getAttribute("style");
		//unitJson["type"] = document.getElementById("cloneOval_" + cntClone).getAttribute("title");
		
		var positionLeftPixel = document.getElementById("cloneOval_" + cntClone).style.left;	
		var positionLeft = parseInt(positionLeftPixel.replace("px", ''));		
		unitJson["positionLeftPercent"] = getPercentageValue(mainDivWidth, positionLeft);
		var positionTopPixel = document.getElementById("cloneOval_" + cntClone).style.top;	
		var positionTop = parseInt(positionTopPixel.replace("px", ''));		
		unitJson["positionTopPercent"] = getPercentageValue(mainDivHeight, positionTop);
		var positionWidthPixel = document.getElementById("cloneOval_" + cntClone).style.width;	
		var positionWidth = parseInt(positionWidthPixel.replace("px", ''));
		unitJson["positionWidthPercent"] = getPercentageValue(mainDivWidth, positionWidth);
		var positionHeightPixel = document.getElementById("cloneOval_" + cntClone).style.height;	
		var positionHeight = parseInt(positionHeightPixel.replace("px", ''));
		unitJson["positionHeightPercent"] = getPercentageValue(mainDivWidth, positionHeight);
		unitJson["aspectRatio"] = ratio;
		
		unitJson["roleImage"] = document.getElementById(cloneId).getElementsByTagName("img")[0].getAttribute("src");
		var roleId = "role_" + cloneId;
		unitJson["roleId"] = roleId;
		
		var roleHeightPixel = document.getElementById(roleId).style.height;	
		var roleHeight = parseInt(roleHeightPixel.replace("px", ''));
		unitJson["roleHeightPercent"] = getPercentageValue(mainDivWidth, roleHeight);
		
		var roleWidthPixel = document.getElementById(roleId).style.width;	
		var roleWidth = parseInt(roleWidthPixel.replace("px", ''));
		unitJson["roleWidthPercent"] = getPercentageValue(mainDivWidth, roleWidth);
		
		var roleLeftPixel = document.getElementById(roleId).style.left;	
		var roleLeft = parseInt(roleLeftPixel.replace("px", ''));
		unitJson["roleLeftPercent"] = getPercentageValue(mainDivWidth, roleLeft);
		
		var roleTopPixel = document.getElementById(roleId).style.top;	
		var roleTop = parseInt(roleTopPixel.replace("px", ''));
		unitJson["roleTopPercent"] = getPercentageValue(mainDivHeight, roleTop);
		
		totalJsonAftSktch[idx++] = unitJson;
		divClone = document.getElementById("cloneOval_" + ++cntClone);
	}
	generatePositionJson();
	ClientSession.set("total_json_after_sketch_final",JSON.stringify(totalJsonAftSktch)); //Total Json to be stored in db
	
	if(sessionStorage.getItem("isLogout") == "N"){
		/*var resultTask = validateTask();
		var selection = validateSelection();*/
				
		var taskInRole = validateTaskInRole();
		
		var resultTask = true;
		var selection = true;
		
		var isRoleNameEmpty = validateEmptyRoleNames();	
		var isRoleNameValid = validateRoleNames();
		var roleCountIsValid = validateRoleCount();
		//var allTextValueProvided = validateTextArea();
		//var duplicateTaskDesc = validateDuplicateValue();
	
		var checkAttribute = validateAttribute();
		var allRelationCraftingValue = false;
		var allTextValueProvided = false;
		var duplicateTaskDesc = false;
		if (sessionStorage.getItem("silentSave")) {
			allRelationCraftingValue = true;
			allTextValueProvided = true;
			duplicateTaskDesc = true;
		} else {
			allTextValueProvided = validateTextArea();			
			if(allTextValueProvided == true) {
				duplicateTaskDesc = validateDuplicateValue();
			} else {
				duplicateTaskDesc = true;
			}
			if(allTextValueProvided == true && duplicateTaskDesc == true) {
				allRelationCraftingValue = validateRelationalCrafting();
			} else {
				allRelationCraftingValue = true;
			}						
		}		
		if (sessionStorage.getItem("silentSave")) {
			storeAfterSketchDiagramSilent(totalJsonAftSktch);
		} else {
			if(sessionStorage.getItem("src") != "Admin") {
			if((resultTask == true) && (selection == true) && (isRoleNameEmpty == true) && (isRoleNameValid == true) && 
					(roleCountIsValid == true) && (allTextValueProvided == true) && (duplicateTaskDesc == true) && 
					(checkAttribute == true) && (allRelationCraftingValue == true) && (taskInRole == true)){
				storeAfterSketchDiagram(totalJsonAftSktch);
			} else if(allTextValueProvided == false){
				alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please add task description.</p>");
				$(".loader_bg").fadeOut();
	    	    $(".loader").hide();
			} else if(duplicateTaskDesc == false){
				alertify.alert("<img src='../img/alert-icon.png'><br /><p>Task description should be different.</p>");
				$(".loader_bg").fadeOut();
	    	    $(".loader").hide();
			} else if(allRelationCraftingValue == false){
				alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please enter Relational Crafting for all tasks. If Relational Crafting is not relevant to a task, write 'N/A'. </p>");
				$(".loader_bg").fadeOut();
	    	    $(".loader").hide();
			} /*else if(((isRoleNameValid == true) && (isRoleNameEmpty == true)) && ((resultTask == false) && (checkAttribute == true) || (selection == false))){
				alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please put all the tasks and elements in role frame.</p>");
				$(".loader_bg").fadeOut();
	    	    $(".loader").hide();
			}*/ else if(((isRoleNameValid == true) && (isRoleNameEmpty == true)) && ((taskInRole == false) && (checkAttribute == true))){
				alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please put atleast one task in role frame.</p>");
				$(".loader_bg").fadeOut();
	    	    $(".loader").hide();
			} else if(isRoleNameEmpty == false){
				alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please provide role name.</p>");
				$(".loader_bg").fadeOut();
	    	    $(".loader").hide();
			} else if(isRoleNameValid == false){
				alertify.alert("<img src='../img/alert-icon.png'><br /><p>Role name should be between 1-120 characters.</p>");
				$(".loader_bg").fadeOut();
	    	    $(".loader").hide();
			} else if(roleCountIsValid == false){
				alertify.alert("<img src='../img/alert-icon.png'><br /><p>Minimum 1 role and maximum 3 roles are allowed.</p>");
				$(".loader_bg").fadeOut();
	    	    $(".loader").hide();
			} else if (checkAttribute == false) {
				alertify.alert("<img src='../img/alert-icon.png'><br /><p>Minimum 1 Value, Strength, Passion is required in your after diagram.</p>");
				$(".loader_bg").fadeOut();
	    	    $(".loader").hide();
			} 
		} else {	
			storeAfterSketchDiagram(totalJsonAftSktch);
		}		
		}
	} else {	
		storeAfterSketchDiagram(totalJsonAftSktch);
	}
	sessionStorage.setItem("next_page", "AS");
	/* sessionStorage.setItem("asView", "H"); */
}
/**
 * function checks if all visible tasks, and attributes are covered in the role.
 * If anything is left out or
 * @returns {Boolean}
 */
function validateTask(){
	var roleIdJson = sessionStorage.getItem("define_role");
	var arr = new Array();
	var json = JSON.parse(roleIdJson);
	for (var index = 0; index < json.length; index++) {
		var innerJson = JSON.stringify(json[index]);
		if (innerJson.search("block") != -1) {
			arr.push(innerJson);
		}
	}
	var taskResult = true;
	$('.after_sketch_task:visible','#outderDiv').each(function(){
		if (arr.toString().search($(this).attr('id')) == -1) {
			taskResult = false;
		}
	});
	if((taskResult == true)){
		return true;
	}else{
		return false;
	}
}
function validateSelection(){
	var roleIdJson = sessionStorage.getItem("define_role");
	var elementResult = true;
	$('.commonDrag:visible','#outderDiv').each(function(){
		if(roleIdJson.search($(this).attr('id')) == -1){
	    	elementResult = false;
	    }
	});
	if((elementResult == true)){
		return true;
	}else{
		return false;
	}
}

/**
 * function checks if atleast one task is present in the role frame.
 * @param null
 * @returns {Boolean}
 */
function validateTaskInRole(){
	var roleIdJson = sessionStorage.getItem("element_in_role");
	var elementResult = true;	
	$('.oval_cloned_div:visible','#outderDiv').each(function(){					
		if(roleIdJson.search($(this).attr('id')) == -1){
	    	elementResult = false;
	    } 
	});
	if((elementResult == true)){
		return true;
	}else{
		return false;
	}
}


/**
 * function checks the blank string for role frame
 * @returns {Boolean}
 */
function validateEmptyRoleNames(){
	var str="";
	var divCountLoop = $('#outderDiv').children('.oval_cloned_div').length;
	for(var i=1; i<=divCountLoop; i++){
			var style = document.getElementById("cloneOval_" +i).getAttribute("style");
			if((style.search('block') != -1)){
				//var taskDesc = $("#cloneOval_" + i + " .speech-bubble textarea").val().trim();	
				var taskDesc = $("#role_cloneOval_" + i + " textarea").val().trim();
				if(taskDesc.length == 0){
					$(".loader_bg").fadeOut();
					$(".loader").hide();
					str = str + "false";
				} 										
			}		
	}
	if(str.search('false')!= -1){
		return false;
	} else {
		return true;
	}
}
/**
 * function checks the text length for role frame
 * @returns {Boolean}
 */
function validateRoleNames(){	
	var divCountLoop = $('#outderDiv').children('.oval_cloned_div').length;
	for(var i=1; i<=divCountLoop; i++){
			var style = document.getElementById("cloneOval_" +i).getAttribute("style");
			if((style.search('block') != -1)){
				//var roleDesc = $("#cloneOval_" + i + " .speech-bubble textarea").val().trim();
				var roleDesc = $("#role_cloneOval_" + i + " textarea").val().trim();
				if(roleDesc.length <= 0){
					return false;
				} else {
					return true;
				}
			}		
	}		
}

function validateRoleCount(){
	clonedDivCount=0;
	var divCountLoop = $('#outderDiv').children('.oval_cloned_div').length;
	for(var i=1; i<=divCountLoop; i++){
			var style = document.getElementById("cloneOval_" +i).getAttribute("style");
			if((style.search('block') != -1)){
				clonedDivCount = clonedDivCount + 1;
			}		
	}
	if(clonedDivCount > 3){
		return false;
	} if(clonedDivCount == 0){
		return false;
	} else {
		return true;
	}
}

/******************* Null check for Task Description ***********************/
function validateTextArea(){
	var divCountLoop = $('#outderDiv').children('.after_sketch_task').length;	
	
	/******************* Set focus to the Task Area which does not contain value***********************/
	if(sessionStorage.getItem("isLogout")=="N"){
		removeFocus();
		for(var i=1; i<=divCountLoop;i++){
			var style = document.getElementById("divImg_" +i).getAttribute("style");
			if((style.search('block') != -1)){
				taskId = document.getElementById("divImg_"+i).id;
				taskDesc = document.getElementById("areaId_"+i).value.trim();	
				if(taskDesc == "") {
					$("#" + taskId).children().find(".add_task_area_sketch").addClass("add_task_area_sketch_focus");
				}
			}		
		}
	}
		
	if(sessionStorage.getItem("isLogout")=="N"){
		for(var i=1; i<=divCountLoop;i++){
			var style = document.getElementById("divImg_" +i).getAttribute("style");
			if((style.search('block') != -1)){
				taskDesc = document.getElementById("areaId_"+i).value.trim();
				if(taskDesc == "") {
					return false;
				}
			}		
		}
	}
	return true;
}

/******************* validate duplicate value for Task Description ***********************/
function validateDuplicateValue(){
	var descArray = new Array();	
	var divCountLoop = $('#outderDiv').children('.after_sketch_task').length;				
	if(sessionStorage.getItem("isLogout")=="N"){
		for(var i=1; i<=divCountLoop;i++){
			var style = document.getElementById("divImg_" +i).getAttribute("style");
			if((style.search('block') != -1)){
				taskDesc = document.getElementById("areaId_"+i).value;
				if(taskDesc != "") {
					descArray.push(taskDesc);
				}
			}		
		}
		for (var i = 0; i < descArray.length; i++) {
	        for (var j = 0; j < descArray.length; j++) {
	            if (i != j) {
	                if (descArray[i].trim().toUpperCase() == descArray[j].trim().toUpperCase()) {	                	
	                	for(var k=1; k<=divCountLoop;k++){
	            			var style = document.getElementById("divImg_" +k).getAttribute("style");
	            			if((style.search('block') != -1)){
	            				taskId = document.getElementById("divImg_"+k).id;
	            				taskDesc = document.getElementById("areaId_"+k).value;	
	            				if(taskDesc.trim().toUpperCase() == descArray[i].trim().toUpperCase()) {
	            					$("#" + taskId).children().find(".add_task_area_sketch").addClass("add_task_area_sketch_focus");
	            				}
	            			}		
	            		}	
	                	return false;
	                }
	            }
	        }
	    }
	}
	return true;
}

/**
 * Function add to calculate the total time
 * @param null
 */
function getTotalTimeSpent(){
	
	if (sessionStorage.getItem("asTimePage")) {
		var timeCounterSplit = sessionStorage.getItem("asTimePage").split(":");
		var hour = parseInt(timeCounterSplit[0]);
		var min = parseInt(timeCounterSplit[1]);
		var sec = parseInt(timeCounterSplit[2]);
		if(min != 0){
			min = min * 60;
		}
		if(hour != 0){
			hour = hour * 60 * 60;
		}
		var totalTimeInSec = hour+min+sec;
		return totalTimeInSec;
	} else {
		return 0;
	}
}

/**
 * Function add to disable the cross image,
 * to change the input field to output text
 * while capturing the image
 * @param null
 */
function closeAllCrossForBlockDiv() {
	var taskDesc = "";
	var additionalDesc = "";
	var totalDivCount = $('#outderDiv').children('.after_sketch_task').length;		// total task block count
	for(var index = 1; index <= totalDivCount; index++) {
		var elementStyle = document.getElementById("divImg_"+index).style.display;
		if ((elementStyle.search('block') != -1 )) { 				//BLOCK -- DISPLAYING
			document.getElementById("divImg_"+index+"_x").style.display = "none";	
			document.getElementById("divImg_"+index+"_lock").style.display = "none";  // remove the lock unlock image in screenshot
			
			$("#divImg_" +index).children().find(".add_task_area_sketch").addClass("task_area_bright");
			if(document.getElementById("energyId_"+index).value >=5 && document.getElementById("energyId_"+index).value <=9){					
				$("#divImg_" +index).children().find(".energy_field").children().addClass("as_small_remove_text");	
			} else {
				$("#divImg_" +index).children().find(".energy_field").children().addClass("as_remove_input_text");
			}	
			if(document.getElementById("energyId_"+index).value >=5 && document.getElementById("energyId_"+index).value <=20){	
				if(document.getElementById("additionalId_"+index).value != null && 
						document.getElementById("additionalId_"+index).value.trim() != "null" && 
						document.getElementById("additionalId_"+index).value.trim() != "" && document.getElementById("additionalId_"+index).value.trim() != "N/A"){	
					//$("#divImg_" +index).children().find(".form-control-additional").addClass("smaller_font");	
				}															
			} 
			if(document.getElementById("energyId_"+index).value >=5 && document.getElementById("energyId_"+index).value <= 9){							
				taskDesc = document.getElementById("areaId_"+index).value;
				additionalDesc = document.getElementById("additionalId_"+index).value;
				if(taskDesc.length <= 100) {
					$("#divImg_" +index).children().find(".add_task_area_sketch").addClass("smaller_task_font_5_1");	
				} else if (taskDesc.length > 100 && taskDesc.length <= 150) {
					$("#divImg_" +index).children().find(".add_task_area_sketch").addClass("smaller_task_font_5_2");	
				} else {
					$("#divImg_" +index).children().find(".add_task_area_sketch").addClass("smaller_task_font_5_3");	
				}
				if(additionalDesc.length <= 35) {
					$("#divImg_" +index).children().find(".form-control-additional").addClass("smaller_font_additional_5_1");	
				} else {
					$("#divImg_" +index).children().find(".form-control-additional").addClass("smaller_font_additional_5_2");	
				}
			} else if(document.getElementById("energyId_"+index).value >=10 && document.getElementById("energyId_"+index).value <= 14){	
				taskDesc = document.getElementById("areaId_"+index).value;
				additionalDesc = document.getElementById("additionalId_"+index).value;
				if(taskDesc.length <= 100) {
					$("#divImg_" +index).children().find(".add_task_area_sketch").addClass("smaller_task_font_10_1");	
				} else if (taskDesc.length > 100 && taskDesc.length <= 150) {
					$("#divImg_" +index).children().find(".add_task_area_sketch").addClass("smaller_task_font_10_2");	
				} else {
					$("#divImg_" +index).children().find(".add_task_area_sketch").addClass("smaller_task_font_10_3");	
				}
				if(additionalDesc.length <= 35) {
					$("#divImg_" +index).children().find(".form-control-additional").addClass("smaller_font_additional_10_1");	
				} else {
					$("#divImg_" +index).children().find(".form-control-additional").addClass("smaller_font_additional_10_2");	
				}
			} else if(document.getElementById("energyId_"+index).value >=15 && document.getElementById("energyId_"+index).value <=20){							
				$("#divImg_" +index).children().find(".form-control-additional").addClass("smaller_font_additional_15");
				taskDesc = document.getElementById("areaId_"+index).value;
				if(taskDesc.length <= 150) {
					$("#divImg_" +index).children().find(".add_task_area_sketch").addClass("smaller_task_font_15_1");	
				} else {
					$("#divImg_" +index).children().find(".add_task_area_sketch").addClass("smaller_task_font_15_2");	
				}
			} else if (document.getElementById("energyId_"+index).value >20 && document.getElementById("energyId_"+index).value <=29) {
				taskDesc = document.getElementById("areaId_"+index).value;
				if(taskDesc.length <= 150) {
					$("#divImg_" +index).children().find(".add_task_area_sketch").addClass("smaller_task_font_20_1");	
				} else {
					$("#divImg_" +index).children().find(".add_task_area_sketch").addClass("smaller_task_font_20_2");	
				}
				$("#divImg_" +index).children().find(".form-control-additional").addClass("smaller_font_additional_20");
			} else if (document.getElementById("energyId_"+index).value >=30 && document.getElementById("energyId_"+index).value <=40) {				
				taskDesc = document.getElementById("areaId_"+index).value;
				if(taskDesc.length <= 150) {
					$("#divImg_" +index).children().find(".add_task_area_sketch").addClass("smaller_task_font_30_1");	
				} else {
					$("#divImg_" +index).children().find(".add_task_area_sketch").addClass("smaller_task_font_30_2");	
				}
				$("#divImg_" +index).children().find(".form-control-additional").addClass("smaller_font_additional_30");	
			} else if (document.getElementById("energyId_"+index).value >40 && document.getElementById("energyId_"+index).value <=60) {
				//$("#divImg_" +index).children().find(".add_task_area_sketch").addClass("larger_task_font");	
				taskDesc = document.getElementById("areaId_"+index).value;
				if(taskDesc.length <= 150) {
					$("#divImg_" +index).children().find(".add_task_area_sketch").addClass("smaller_task_font_40_1");	
				} else {
					$("#divImg_" +index).children().find(".add_task_area_sketch").addClass("smaller_task_font_40_2");							
				}
				$("#divImg_" +index).children().find(".form-control-additional").addClass("smaller_font_additional_40");
			} else {
				$("#divImg_" +index).children().find(".add_task_area_sketch").addClass("larger_task_font");	
				$("#divImg_" +index).children().find(".form-control-additional").addClass("smaller_font_additional_60");
			}
			$(".oval_cloned_div").children().find(".speech_bubble_area").addClass("role_frame_font");	
			
			$("#divImg_" + index).removeClass("lock_single_task_item_sketch");
			$("#divImg_" + index).addClass("single_task_item_sketch");	
			$("#energyId_" + index).removeClass("energy_lock");
		}
	}
	
	$(".removeService").css({ display: "none" });
	$(".commonDrag").css("z-index", 4);
	$(".oval_cloned_div").css("z-index", 140);		
	
	/**** ADDED FOR ROLE FRAME SPEECH BUBBLE ****/
	var totalRoleCount = $('#outderDiv').children('.speech-bubble').length;
	for(var index = 1; index <= totalRoleCount; index++) {
		var roleValue = $("#role_cloneOval_" + index + " textarea").val().replace(/\(/g, " (").replace(/\)/g, ") ");
		document.getElementById("text_cloneOval_" +index).value = roleValue;
		
		/*$("#role_cloneOval_" + index).children(".ui-resizable-nw").css("display","none");
		$("#role_cloneOval_" + index).children(".ui-resizable-ne").css("display","none");
		$("#role_cloneOval_" + index).children(".ui-resizable-sw").css("display","none");
		$("#role_cloneOval_" + index).children(".ui-resizable-se").css("display","none");	 */		
	}
	$("#clickableElement1 img").removeAttr('src');
	$("#clickableElement2 img").removeAttr('src');
}

/**
 * Function add to enable the cross image
 * while capturing the image
 * @param null
 */
function openAllCrossForBlockDiv() {
	var totalDivCount = $('#outderDiv').children('.after_sketch_task').length;		// total task block count
	for(var index = 1; index <= totalDivCount; index++) {
		var elementStyle = document.getElementById("divImg_"+index).style.display;
		if ((elementStyle.search('block') != -1 )) { 				//BLOCK -- DISPLAYING
			document.getElementById("divImg_"+index+"_x").style.display = "block";
			document.getElementById("divImg_"+index+"_lock").style.display = "block"; 
			$(".single_task_item_sketch").children().find(".energy_field").children().removeClass("as_remove_input_text");
			$(".single_task_item_sketch").children().find(".energy_field").children().removeClass("as_small_remove_text");
			$("#energyId_" + index).removeClass("energy_lock");
			$("#divImg_" + index).removeClass("lock_single_task_item_sketch");
			$("#divImg_" + index).addClass("single_task_item_sketch");
			
			$("#additionalId_" + index).children().find(".form-control-additional").removeClass("small_font");
			$("#additionalId_" + index).children().find(".form-control-additional").removeClass("smaller_font_additional_common");
			$("#additionalId_" + index).children().find(".form-control-additional").removeClass("smaller_font_additional_5_1");
			$("#additionalId_" + index).children().find(".form-control-additional").removeClass("smaller_font_additional_5_2");
			$("#additionalId_" + index).children().find(".form-control-additional").removeClass("smaller_font_additional_10_1");
			$("#additionalId_" + index).children().find(".form-control-additional").removeClass("smaller_font_additional_10_2");
			$("#additionalId_" + index).children().find(".form-control-additional").removeClass("smaller_font_additional_15");
			$("#additionalId_" + index).children().find(".form-control-additional").removeClass("smaller_font_additional_20");
			$("#additionalId_" + index).children().find(".form-control-additional").removeClass("smaller_font_additional_30");
			$("#additionalId_" + index).children().find(".form-control-additional").removeClass("smaller_font_additional_40");
			
			
			$("#divImg_" +index).children().find(".add_task_area_sketch").removeClass("smaller_task_font_5_1");
			$("#divImg_" +index).children().find(".add_task_area_sketch").removeClass("smaller_task_font_5_2");
			$("#divImg_" +index).children().find(".add_task_area_sketch").removeClass("smaller_task_font_5_3");
			$("#divImg_" +index).children().find(".add_task_area_sketch").removeClass("smaller_task_font_10_1");
			$("#divImg_" +index).children().find(".add_task_area_sketch").removeClass("smaller_task_font_10_2");
			$("#divImg_" +index).children().find(".add_task_area_sketch").removeClass("smaller_task_font_10_3");
			$("#divImg_" +index).children().find(".add_task_area_sketch").removeClass("smaller_task_font_15_1");
			$("#divImg_" +index).children().find(".add_task_area_sketch").removeClass("smaller_task_font_15_2");

			$("#divImg_" +index).children().find(".add_task_area_sketch").removeClass("small_task_font");
			$("#divImg_" +index).children().find(".add_task_area_sketch").removeClass("large_task_font");
			$("#divImg_" +index).children().find(".add_task_area_sketch").removeClass("larger_task_font");
			$("#divImg_" +index).children().find(".add_task_area_sketch").removeClass("task_area_bright");
		}
	}
	/* control on display of oval */
	var totalOvalCount = $('#outderDiv').children('.oval_cloned_div').length;		// total oval count
	for(var index = 1; index <= totalOvalCount; index++) {	
		var elementStyleOval = document.getElementById("cloneOval_"+index).style.display;
		if ((elementStyleOval.search('block') != -1 )) {		//BLOCK -- DISPLAYING
			if (document.getElementById("cloneOval_"+index+"_x")) {
				document.getElementById("cloneOval_"+index+"_x").style.display = "block";
			}
		}
		$("#outderDiv .oval_cloned_div .removeService").css({opacity: "1"});		
	}
	/* control on display of passion */
	var totalPassionCount = $('#outderDiv').children('.up_triangle_small').length;		
	// total passion count
	for(var index = 1; index <= totalPassionCount; index++) {		
		var elementStylePassion = document.getElementById("divPassionClone_"+index).style.display;
		if ((elementStylePassion.search('block') != -1 )) { 	//BLOCK -- DISPLAYING
			document.getElementById("divPassionClone_"+index+"_x").style.display = "block";				
		}
	}	
	/* control on display of strength */
	var totalStrengthCount = $('#outderDiv').children('.plus_sign_small').length;		// total strength count
	for(var index = 1; index <= totalStrengthCount; index++) {		
		var elementStyleStrength = document.getElementById("divStrengthClone_"+index).style.display;
		if ((elementStyleStrength.search('block') != -1 )) { 				//BLOCK -- DISPLAYING
			document.getElementById("divStrengthClone_"+index+"_x").style.display = "block";				
		}
	}
	/* control on display of value */
	var totalValueCount = $('#outderDiv').children('.oval_small').length;		// total value count
	for(var index = 1; index <= totalValueCount; index++) {		
		var elementStyleValue = document.getElementById("divValueClone_"+index).style.display;
		if ((elementStyleValue.search('block') != -1 )) { 				//BLOCK -- DISPLAYING
			document.getElementById("divValueClone_"+index+"_x").style.display = "block";				
		}
	}
}

/**
 * function stores diagram data in the database
 * @param json
 */
function storeAfterSketchDiagram(json){
	getOuterDivHeight();
	sessionStorage.setItem("as2View","H");
	sessionStorage.setItem("pageSequence", 5);
	var timeSpent = getTotalTimeSpent();
	var outerDivHeight = document.getElementById('pageWrapCopy').style.height;
	var outerDivWidth = document.getElementById('pageWrapCopy').style.width;	
	var element = document.getElementById("outderDiv");
	closeAllCrossForBlockDiv();
	html2canvas(element, {
		onrendered: function(canvas) {
            var snapShotURL = canvas.toDataURL("image/png"); //get's image string
            openAllCrossForBlockDiv();
            
            var isEdit = "N";
            if (sessionStorage.getItem("bsEdit")) {
            	isEdit = "Y";
            }
            
            var afterSketchModel = Spine.Model.sub();
            afterSketchModel.configure("/user/saveAfterSketchDiagram", "createdBy", "jsonString", "jobReferenceString", 
            		"snapShot", "jobTitle", "roleJson", "divHeight", "divWidth", "totalCount", "divInitialVal", 
            		"timeSpents", "isCompleteds" , "profileId", "positionJson", "isEdits", "linkClicked", 
            		"rowId", "initialSave", "isPrevious", "jctUserId", "elementOutsideRoleJson");
            afterSketchModel.extend( Spine.Model.Ajax );
            var isCompleted = "N";
            if(sessionStorage.getItem("isLogout") == "Y"){
            	isCompleted = "Y";
            }
            //Populate the model with data to transfer
        	var modelPopulator = new afterSketchModel({  
        		createdBy: sessionStorage.getItem("jctEmail"), 
        		jsonString: json,
        		jobReferenceString: sessionStorage.getItem("jobReferenceNo"),
        		snapShot: snapShotURL,
        		jobTitle: sessionStorage.getItem("jobTitle"),
        		roleJson: jQuery.parseJSON(sessionStorage.getItem("define_role").replace(/\?/g, ";_;")),
        		divHeight: outerDivHeight,
        		divWidth: outerDivWidth,
        		totalCount: sessionStorage.getItem("total_count"),
        		divInitialVal: JSON.parse(sessionStorage.getItem("div_intial_value")),
        		timeSpents : timeSpent,
        		isCompleteds: isCompleted,
        		profileId: sessionStorage.getItem("profileId"),
        		positionJson: sessionStorage.getItem("position_json"),
        		isEdits: isEdit,
        		linkClicked: sessionStorage.getItem("linkClicked"),
        		rowId: sessionStorage.getItem("rowIdentity"),
        		initialSave: "N",
        		isPrevious: "N",
        		jctUserId: sessionStorage.getItem("jctUserId"),
        		elementOutsideRoleJson: jQuery.parseJSON(sessionStorage.getItem("element_outside_role").replace(/\?/g, ";_;"))
        	});
        	modelPopulator.save(); //POST
        	afterSketchModel.bind("ajaxError", function(record, xhr, settings, error) {
        		$(".loader_bg").fadeOut();
        	    $(".loader").hide();
        	});
        	
        	afterSketchModel.bind("ajaxSuccess", function(record, xhr, settings, error) {
        		var jsonStr = JSON.stringify(xhr);
    			var obj = jQuery.parseJSON(jsonStr);
    			if(sessionStorage.getItem("isLogout") == "N"){
    					removePresentSessionItems();
        				if (isEdit == "Y") {
        					sessionStorage.setItem("pageSequence", 8);
        					window.location = "finalPage.jsp";
        				} else {
    	    				sessionStorage.setItem("snapShotURLS", obj.snapShot);
    	    				sessionStorage.setItem("isCompleted", 1);
    	    				window.location = "actionPlan.jsp";
        				}
        			} else { //It is logout
        				//Create a model
        				var LogoutVO = Spine.Model.sub();
        				LogoutVO.configure("/user/auth/logout", "jobReferenceString", "rowId", "lastActivePage");
        				LogoutVO.extend( Spine.Model.Ajax );
        				//Populate the model with data to transder
        				var logoutModel = new LogoutVO({  
        					jobReferenceString: sessionStorage.getItem("jobReferenceNo"),
        					rowId: sessionStorage.getItem("rowIdentity"),
        					lastActivePage: "/user/view/afterSketch2.jsp"
        				});
        				logoutModel.save(); //POST
        				LogoutVO.bind("ajaxError", function(record, xhr, settings, error) {
        					$(".loader_bg").fadeOut();
        				    $(".loader").hide();
        					alertify.alert("Unable to connect to the server.");
        				});
        				
        				LogoutVO.bind("ajaxSuccess", function(record, xhr, settings, error) {
        					sessionStorage.clear();
        					window.location = "../signup-page.jsp";
        				});
        			}    			
        	});
            }
        });
}
	
/**
 * Function to create the json object
 * of the element which are with in a role
 * @param jsonObjOval
 * @param jsonObjElememt
 */
/*function findIdByRole(jsonObjOval, jsonObjElememt){
	var jsonObj = [];
	var counter = 0;
	var cntOval = jsonObjOval.length;
	var cntElement = jsonObjElememt.length;
	for (var i = 0; i<cntOval;i++){
		var unitJ = {};			
		var divOvalId = jsonObjOval[i]["divId"];	
		var cloneValue = jsonObjOval[i]["roleValue"];
		var leftOval = jsonObjOval[i]["LeftPos"];	
		var topOval = jsonObjOval[i]["TopPos"];		
		var cloneStyle = jsonObjOval[i]["displayDeleted"];
		//alert(jsonObjOval[i]["displayDeleted"]);
		var extremeLeftOval = jsonObjOval[i]["horizontalLeftPos"];
		var extremeTopOval = jsonObjOval[i]["verticalTopPos"];
		for (var j = 0; j<cntElement;j++){
			var divElementId = jsonObjElememt[j]["divId"];
			var divElementValue = jsonObjElememt[j]["divValue"];
			var leftElement = jsonObjElememt[j]["LeftPos"];	
			var topElement = jsonObjElememt[j]["TopPos"];		
			var extremeLeftElement = jsonObjElememt[j]["horizontalLeftPos"];
			var extremeTopElement = jsonObjElememt[j]["verticalTopPos"];			
			if(leftElement>leftOval && topElement>topOval && extremeLeftElement<extremeLeftOval && extremeTopElement<extremeTopOval){
				unitJ["RoleId"] = divOvalId;
				unitJ["RoleValue"] = cloneValue.replace(/,/g, "-_-");
				unitJ["ElementId"+j] = divElementId;
				unitJ["ElementValue"+j] = divElementValue;
				unitJ["CloneStyle"] = cloneStyle;
			}						
		}
		jsonObj[counter++] = unitJ;
		}
	ClientSession.set("define_role",JSON.stringify(jsonObj));  	
}*/

function findIdByRole(jsonObjOval, jsonObjElememt){
	var jsonObj = [];
	var counter = 0;
	var cntOval = jsonObjOval.length; //1
	var cntElement = jsonObjElememt.length; //3
	for (var i = 0; i<cntOval;i++){
		var unitJ = {};			
		var divOvalId = jsonObjOval[i]["divId"];	
		var cloneValue = jsonObjOval[i]["roleValue"];
		var leftOval = jsonObjOval[i]["LeftPos"];	
		var topOval = jsonObjOval[i]["TopPos"];		
		var cloneStyle = jsonObjOval[i]["displayDeleted"];
		//alert(jsonObjOval[i]["displayDeleted"]);
		var extremeLeftOval = jsonObjOval[i]["horizontalLeftPos"];
		var extremeTopOval = jsonObjOval[i]["verticalTopPos"];
		for (var j = 0; j<cntElement;j++){
			var divElementId = jsonObjElememt[j]["divId"];
			var divElementValue = jsonObjElememt[j]["divValue"];
			var leftElement = jsonObjElememt[j]["LeftPos"];	
			var topElement = jsonObjElememt[j]["TopPos"]+10;		
			var extremeLeftElement = jsonObjElememt[j]["horizontalLeftPos"];
			var extremeTopElement = jsonObjElememt[j]["verticalTopPos"];			
			if(leftElement>leftOval && topElement>topOval && extremeLeftElement<extremeLeftOval && extremeTopElement<extremeTopOval){
				unitJ["RoleId"] = divOvalId;
				unitJ["RoleValue"] = cloneValue.replace(/,/g, "-_-");
				unitJ["ElementId"+j] = divElementId;
				unitJ["ElementValue"+j] = divElementValue;
				unitJ["CloneStyle"] = cloneStyle;
			} 	
		}
		if (JSON.stringify(unitJ) != "{}")
			jsonObj[counter++] = unitJ;
		}
	ClientSession.set("define_role",JSON.stringify(jsonObj)); 
	
	/****
	 * ADDED FOR PUBLIC VERSION
	 */
	var jsonObj1 = [];
	var counter1 = 0;
	var myJson = JSON.stringify(jsonObj);
	for (var i = 0; i < jsonObj.length; i++) {
		var unitJson = {};		
		for (var j = 0; j<cntElement;j++){
			var divElementId = jsonObjElememt[j]["divId"];
			var divElementValue = jsonObjElememt[j]["divValue"];
			if (myJson.search(divElementId) == -1) {
				unitJson["RoleId"] = "N/A";
				unitJson["RoleValue"] = "N/A";
				unitJson["ElementId"+j] = divElementId;
				unitJson["ElementValue"+j] = divElementValue;
				unitJson["CloneStyle"] = "N/A";		
			}
		}
		
		jsonObj1[counter1++] = unitJson;
	}	
	ClientSession.set("element_outside_role",JSON.stringify(jsonObj1)); 
	/********** ENDED ***********/
	
	var jsonObj2 = [];
	var counter3 = 0;
	var cntOval = jsonObjOval.length; //1
	var cntElement = jsonObjElememt.length; //3
	for (var i = 0; i<cntOval;i++){
		var unitJ = {};			
		var divOvalId = jsonObjOval[i]["divId"];	
		var cloneValue = jsonObjOval[i]["roleValue"];
		var leftOval = jsonObjOval[i]["LeftPos"];	
		var topOval = jsonObjOval[i]["TopPos"];		
		var extremeLeftOval = jsonObjOval[i]["horizontalLeftPos"];
		var extremeTopOval = jsonObjOval[i]["verticalTopPos"];
		for (var j = 0; j<cntElement;j++){
			var divElementId = jsonObjElememt[j]["divId"];
			//var divElementValue = jsonObjElememt[j]["divValue"];
			var leftElement = jsonObjElememt[j]["LeftPos"];	
			var topElement = jsonObjElememt[j]["TopPos"]+10;		
			var extremeLeftElement = jsonObjElememt[j]["horizontalLeftPos"];
			var extremeTopElement = jsonObjElememt[j]["verticalTopPos"];	
			if(divElementId.search("divImg_") != -1) {			
				if(leftElement>leftOval && topElement>topOval && extremeLeftElement<extremeLeftOval && extremeTopElement<extremeTopOval){					
					unitJ["RoleId"] = divOvalId;
					unitJ["RoleValue"] = cloneValue.replace(/,/g, "-_-");
					unitJ["ElementId"+j] = divElementId;			
				} 
			}				
		}
		jsonObj2[counter3++] = unitJ;
		}
	ClientSession.set("element_in_role",JSON.stringify(jsonObj2)); 
}


/**
 * Function to create the json object  
 * to find the element Id associated with in a role
 * @param null
 */
function generatePositionJson() {	
	var jsonObjOval = [];
	var idxOval = 0;	
	var cntClone = 1;
	var divClone = document.getElementById("cloneOval_" + cntClone);
	while(divClone != null) {	
		var position = $( "#cloneOval_" + cntClone ).position();
		var width = $( "#cloneOval_" + cntClone ).width();
		var height = $( "#cloneOval_" + cntClone ).height();
		var horizontalLeftPos = position.left + width;
		var verticalTopPos = position.top + height;
		var unitJsonOval = {};		
		unitJsonOval["divId"] = document.getElementById("cloneOval_" + cntClone).id;
		var cloneId = unitJsonOval["divId"];		
		//var cloneValue = $("#" + cloneId + " .speech-bubble textarea").val();
		var cloneValue = $("#role_" + cloneId + " textarea").val();
		unitJsonOval["roleValue"] = cloneValue;
		unitJsonOval["LeftPos"] = position.left;		
		unitJsonOval["horizontalLeftPos"] = horizontalLeftPos;
		unitJsonOval["TopPos"] = position.top;
		unitJsonOval["verticalTopPos"] = verticalTopPos;
		unitJsonOval["displayDeleted"] = document.getElementById("cloneOval_" + cntClone).getAttribute("style");		
		jsonObjOval[idxOval++] = unitJsonOval;
		divClone = document.getElementById("cloneOval_" + ++cntClone);
	}
	
	var jsonObjElememt = [];
	var idx = 0;
	var cntTask = 1;	
	var divValueTask = document.getElementById("divImg_" + cntTask);
	while(divValueTask != null) {
		var position = $( "#divImg_" + cntTask ).position();
		var width = $( "#divImg_" + cntTask ).width();
		var height = $( "#divImg_" + cntTask ).height();
		var horizontalLeftPos = position.left + width;
		var verticalTopPos = position.top + height;
		var unitJson = {};		
		unitJson["divId"] = document.getElementById("divImg_" + cntTask).id;	
		unitJson["LeftPos"] = position.left;
		unitJson["horizontalLeftPos"] = horizontalLeftPos;
		unitJson["TopPos"] = position.top;
		unitJson["verticalTopPos"] = verticalTopPos;		
		jsonObjElememt[idx++] = unitJson;
		divValueTask = document.getElementById("divImg_" + ++cntTask);
	}
	
	
	var cntStrength = 1;
	var divStrength = document.getElementById("divStrengthClone_" + cntStrength);
	while(divStrength != null) {
			var position = $( "#divStrengthClone_" + cntStrength ).position();
			var width = $( "#divStrengthClone_" + cntStrength ).width();
			var height = $( "#divStrengthClone_" + cntStrength ).height();
			var horizontalLeftPos = position.left + width;
			var verticalTopPos = position.top + height;
			var unitJson = {};		
			unitJson["divId"] = document.getElementById("divStrengthClone_" + cntStrength).id;	
			unitJson["divValue"] = document.getElementById("divStrengthClone_" + cntStrength).value;	
			unitJson["LeftPos"] = position.left;
			unitJson["horizontalLeftPos"] = horizontalLeftPos;
			unitJson["TopPos"] = position.top;
			unitJson["verticalTopPos"] = verticalTopPos;		
			jsonObjElememt[idx++] = unitJson;			
			divStrength = document.getElementById("divStrengthClone_" + ++cntStrength);		
	}	
	
	var cntValue = 1;
	var divValue = document.getElementById("divValueClone_" + cntValue);
	while(divValue != null) {
		var position = $( "#divValueClone_" + cntValue ).position();
		var width = $( "#divValueClone_" + cntValue ).width();
		var height = $( "#divValueClone_" + cntValue ).height();
		var horizontalLeftPos = position.left + width;
		var verticalTopPos = position.top + height;
		var unitJson = {};		
		unitJson["divId"] = document.getElementById("divValueClone_" + cntValue).id;
		unitJson["divValue"] = document.getElementById("divValueClone_" + cntValue).value;
		unitJson["LeftPos"] = position.left;
		unitJson["horizontalLeftPos"] = horizontalLeftPos;
		unitJson["TopPos"] = position.top;
		unitJson["verticalTopPos"] = verticalTopPos;		
		jsonObjElememt[idx++] = unitJson;	
		divValue = document.getElementById("divValueClone_" + ++cntValue);
	}

	var cntPassion = 1;
	var divPassion = document.getElementById("divPassionClone_" + cntPassion);
	while(divPassion != null) {
		var position = $( "#divPassionClone_" + cntPassion ).position();
		var width = $( "#divPassionClone_" + cntPassion ).width();
		var height = $( "#divPassionClone_" + cntPassion ).height();
		var horizontalLeftPos = position.left + width;
		var verticalTopPos = position.top + height;
		var unitJson = {};		
		unitJson["divId"] = document.getElementById("divPassionClone_" + cntPassion).id;		
		unitJson["divValue"] = document.getElementById("divPassionClone_" + cntPassion).value;
		unitJson["LeftPos"] = position.left;
		unitJson["horizontalLeftPos"] = horizontalLeftPos;
		unitJson["TopPos"] = position.top;
		unitJson["verticalTopPos"] = verticalTopPos;
		jsonObjElememt[idx++] = unitJson;	
		divPassion = document.getElementById("divPassionClone_" + ++cntPassion);
	}	
	findIdByRole(jsonObjOval, jsonObjElememt);	
}


/**
 * function is used to navigate to the aftersketch1 page.
 */
function backToPreviousPage(){
	$(".loader_bg").fadeIn();
	$(".loader").fadeIn();
	getActualPosition();	
	nextPageAfterSketch = 1;
	ClientSession.set("Next_page_after_sketch", nextPageAfterSketch);	
	var cntTask = 1;
	var totalJsonAftSktch = [];
	var idx = 0;
	var divValueTask = document.getElementById("divImg_" + cntTask);
	
	var mainDivWidth = $("#outderDiv").width();
	var mainDivHeight = $("#outderDiv").height();	
    
	var gcd=calc(mainDivWidth,mainDivHeight);
  	var r1=mainDivWidth/gcd;
    var r2=mainDivHeight/gcd;
    var ratio=r1+":"+r2;
    
	while(divValueTask != null) {
		var unitJson = {};					
		unitJson["divId"] = document.getElementById("divImg_" + cntTask).id;
		unitJson["divIdx"] = document.getElementById("divImg_" + cntTask +"_x").id;
		unitJson["energyId"] = document.getElementById("energyId_"+ cntTask).id;
		unitJson["energyValue"] = document.getElementById("energyId_"+ cntTask).value;
		unitJson["style"] = document.getElementById("divImg_" + cntTask).getAttribute("style");
		unitJson["areaId"] = document.getElementById("areaId_"+ cntTask).id;
		unitJson["areaValue"] = document.getElementById("areaId_"+ cntTask).value.replace(/,/g, "`").replace(/#/g, "-_-").replace(/\?/g, ";_;");
		unitJson["additionalId"] = document.getElementById("additionalId_"+ cntTask).id;
		unitJson["additionalValue"] = document.getElementById("additionalId_"+ cntTask).value.replace(/,/g, "`").replace(/#/g, "-_-").replace(/\?/g, ";_;");		
		unitJson["positionLeft"] = document.getElementById("divImg_" + cntTask).style.left;
		unitJson["positionTop"] = document.getElementById("divImg_" + cntTask).style.top;
		unitJson["positionWidth"] = document.getElementById("divImg_" + cntTask).style.width;
		unitJson["positionHeight"] = document.getElementById("divImg_" + cntTask).style.height;
		
		var positionLeftPx = document.getElementById("divImg_" + cntTask).style.left;	
		var positionLeftPixel = positionLeftPx.replace("px", '');
		unitJson["positionLeftPixel"] = positionLeftPixel;
		unitJson["positionLeftPercent"] = getPercentageValue(mainDivWidth, positionLeftPixel);
		var positionTopPx = document.getElementById("divImg_" + cntTask).style.top;	
		var positionTopPixel = positionTopPx.replace("px", '');
		unitJson["positionTopPixel"] = positionTopPixel;
		unitJson["positionTopPercent"] = getPercentageValue(mainDivHeight, positionTopPixel);
		var positionWidthPixel = document.getElementById("divImg_" + cntTask).style.width;	
		var positionWidth = parseInt(positionWidthPixel.replace("px", ''));
		unitJson["positionWidthPercent"] = getPercentageValue(mainDivWidth, positionWidth);
		unitJson["mainDivWidth"] = mainDivWidth;
		unitJson["mainDivHeight"] = mainDivHeight;
		unitJson["aspectRatio"] = ratio;
		
		/***** Added for lock and unlock feature *****/
		unitJson["divLockId"] = document.getElementById("divImg_" + cntTask +"_lock").id;
		unitJson["lockImage"] = document.getElementById("divImg_" + cntTask+ "_lock").getElementsByTagName("img")[0].getAttribute("src");
		
		totalJsonAftSktch[idx++] = unitJson;
		divValueTask = document.getElementById("divImg_" + ++cntTask);
	}
	
	
	var cntStrength = 1;
	var divStrength = document.getElementById("divStrengthClone_" + cntStrength);
	while(divStrength != null) {
			var unitJson = {};						
			unitJson["divId"] = document.getElementById("divStrengthClone_" + cntStrength).id;	
			var divStrengthId = unitJson["divId"];		
			var divStrengthVal = $("#" + divStrengthId + "  p").text();
			unitJson["divValue"] = divStrengthVal;		
			unitJson["style"] = document.getElementById("divStrengthClone_" + cntStrength).getAttribute("style");
			unitJson["positionLeft"] = document.getElementById("divStrengthClone_" + cntStrength).style.left;
			unitJson["positionTop"] = document.getElementById("divStrengthClone_" + cntStrength).style.top;	
			
			var positionLeftPixel = document.getElementById("divStrengthClone_" + cntStrength).style.left;	
			var positionLeft = parseInt(positionLeftPixel.replace("px", ''));			
			unitJson["positionLeftPercent"] = getPercentageValue(mainDivWidth, positionLeft);
			var positionTopPixel = document.getElementById("divStrengthClone_" + cntStrength).style.top;	
			var positionTop = parseInt(positionTopPixel.replace("px", ''));
			unitJson["positionTopPercent"] = getPercentageValue(mainDivHeight, positionTop);
			unitJson["aspectRatio"] = ratio;
			
			totalJsonAftSktch[idx++] = unitJson;		
			divStrength = document.getElementById("divStrengthClone_" + ++cntStrength);		
	}	
	
	var cntValue = 1;
	var divValue = document.getElementById("divValueClone_" + cntValue);
	while(divValue != null) {
		var unitJson = {};		
		unitJson["divId"] = document.getElementById("divValueClone_" + cntValue).id;	
		var divValueId = unitJson["divId"];		
		var divValueVal = $("#" + divValueId + "  p").text();
		unitJson["divValue"] = divValueVal;	
		unitJson["style"] = document.getElementById("divValueClone_" + cntValue).getAttribute("style");		
		unitJson["positionLeft"] = document.getElementById("divValueClone_" + cntValue).style.left;
		unitJson["positionTop"] = document.getElementById("divValueClone_" + cntValue).style.top;	
		
		var positionLeftPixel = document.getElementById("divValueClone_" + cntValue).style.left;	
		var positionLeft = parseInt(positionLeftPixel.replace("px", ''));
		unitJson["positionLeftPercent"] = getPercentageValue(mainDivWidth, positionLeft);
		var positionTopPixel = document.getElementById("divValueClone_" + cntValue).style.top;	
		var positionTop = parseInt(positionTopPixel.replace("px", ''));
		unitJson["positionTopPercent"] = getPercentageValue(mainDivHeight, positionTop);
		unitJson["aspectRatio"] = ratio;
		
		totalJsonAftSktch[idx++] = unitJson;
		divValue = document.getElementById("divValueClone_" + ++cntValue);
	}

	var cntPassion = 1;
	var divPassion = document.getElementById("divPassionClone_" + cntPassion);
	while(divPassion != null) {
		var unitJson = {};		
		
		unitJson["divId"] = document.getElementById("divPassionClone_" + cntPassion).id;
		var divPassionId = unitJson["divId"];		
		var divPassionVal = $("#" + divPassionId + "  p").text();
		unitJson["divValue"] = divPassionVal;	
		unitJson["style"] = document.getElementById("divPassionClone_" + cntPassion).getAttribute("style");			
		unitJson["positionLeft"] = document.getElementById("divPassionClone_" + cntPassion).style.left;
		unitJson["positionTop"] = document.getElementById("divPassionClone_" + cntPassion).style.top;		
		var positionLeftPixel = document.getElementById("divPassionClone_" + cntPassion).style.left;
		
		var positionLeft = parseInt(positionLeftPixel.replace("px", ''));
		unitJson["positionLeftPercent"] = getPercentageValue(mainDivWidth, positionLeft);
		var positionTopPixel = document.getElementById("divPassionClone_" + cntPassion).style.top;	
		var positionTop = parseInt(positionTopPixel.replace("px", ''));
		unitJson["positionTopPercent"] = getPercentageValue(mainDivHeight, positionTop);
		unitJson["aspectRatio"] = ratio;
		
		totalJsonAftSktch[idx++] = unitJson;
		divPassion = document.getElementById("divPassionClone_" + ++cntPassion);
	}
	
	var cntClone = 1;
	var divClone = document.getElementById("cloneOval_" + cntClone);
	while(divClone != null) {
		var unitJson = {};	
		
		unitJson["divId"] = document.getElementById("cloneOval_" + cntClone).id;
		var cloneId = unitJson["divId"];		
		//var cloneValue = $("#" + cloneId + " .speech-bubble textarea").val().replace(/#/g, "-_-").replace(/\?/g, ";_;");
		var cloneValue = $("#role_" + cloneId + " textarea").val().replace(/#/g, "-_-").replace(/\?/g, ";_;");
		unitJson["roleValue"] = cloneValue;
		unitJson["positionLeft"] = document.getElementById("cloneOval_" + cntClone).style.left;
		unitJson["positionTop"] = document.getElementById("cloneOval_" + cntClone).style.top;	
		unitJson["positionWidth"] = document.getElementById("cloneOval_" + cntClone).style.width;
		unitJson["positionHeight"] = document.getElementById("cloneOval_" + cntClone).style.height;
		var position = $( "#cloneOval_" + cntClone ).position();
		unitJson["positionLeft"] = position.left;
		unitJson["positionTop"] = position.top;	
		unitJson["style"] = document.getElementById("cloneOval_" + cntClone).getAttribute("style");
	//	unitJson["type"] = document.getElementById("cloneOval_" + cntClone).getAttribute("title");
		var positionLeftPixel = document.getElementById("cloneOval_" + cntClone).style.left;	
		var positionLeft = parseInt(positionLeftPixel.replace("px", ''));	
		unitJson["positionLeftPercent"] = getPercentageValue(mainDivWidth, positionLeft);
		var positionTopPixel = document.getElementById("cloneOval_" + cntClone).style.top;	
		var positionTop = parseInt(positionTopPixel.replace("px", ''));
		unitJson["positionTopPercent"] = getPercentageValue(mainDivHeight, positionTop);
		var positionWidthPixel = document.getElementById("cloneOval_" + cntClone).style.width;	
		var positionWidth = parseInt(positionWidthPixel.replace("px", ''));
		unitJson["positionWidthPercent"] = getPercentageValue(mainDivWidth, positionWidth);
		var positionHeightPixel = document.getElementById("cloneOval_" + cntClone).style.height;	
		var positionHeight = parseInt(positionHeightPixel.replace("px", ''));
		unitJson["positionHeightPercent"] = getPercentageValue(mainDivWidth, positionHeight);
		unitJson["aspectRatio"] = ratio;
		
		unitJson["roleImage"] = document.getElementById(cloneId).getElementsByTagName("img")[0].getAttribute("src");
		var roleId = "role_" + cloneId;
		unitJson["roleId"] = roleId;
		
		/*var roleHeightPixel = document.getElementById(roleId).style.height;	
		var roleHeight = parseInt(roleHeightPixel.replace("px", ''));*/
		var roleHeight = $("#" + roleId).height();
		unitJson["roleHeightPercent"] = getPercentageValue(mainDivWidth, roleHeight);
		
		/*var roleWidthPixel = document.getElementById(roleId).style.width;	
		var roleWidth = parseInt(roleWidthPixel.replace("px", ''));*/
		
		var roleWidth = $("#" + roleId).width();
		unitJson["roleWidthPercent"] = getPercentageValue(mainDivWidth, roleWidth);
		
		var roleLeftPixel = document.getElementById(roleId).style.left;	
		var roleLeft = parseInt(roleLeftPixel.replace("px", ''));
		unitJson["roleLeftPercent"] = getPercentageValue(mainDivWidth, roleLeft);
		
		var roleTopPixel = document.getElementById(roleId).style.top;	
		var roleTop = parseInt(roleTopPixel.replace("px", ''));
		unitJson["roleTopPercent"] = getPercentageValue(mainDivHeight, roleTop);
		
	//	unitJson["resizeNE"] = $("#" + roleId).children(".ui-resizable-ne").css('display');
		//alert($("#" + roleId).children(".ui-resizable-ne").css('display'));
	//	unitJson["resizeNW"] = $("#" + roleId).children(".ui-resizable-nw").getAttribute("style");
	//	unitJson["resizeSE"] = $("#" + roleId).children(".ui-resizable-se").getAttribute("style");
	//	unitJson["resizeSW"] = $("#" + roleId).children(".ui-resizable-s
		totalJsonAftSktch[idx++] = unitJson;
		divClone = document.getElementById("cloneOval_" + ++cntClone);
	}
	generatePositionJson();
	

	sessionStorage.setItem("as2View","H");
	sessionStorage.setItem("pageSequence", 3);
	var timeSpent = getTotalTimeSpent();
	var outerDivHeight = document.getElementById('pageWrapCopy').style.height;
	var outerDivWidth = document.getElementById('pageWrapCopy').style.width;
	var element = document.getElementById("outderDiv");
	closeAllCrossForBlockDiv();
	html2canvas(element, {
		onrendered: function(canvas) {
            var snapShotURL = canvas.toDataURL("image/png"); //get's image string
            openAllCrossForBlockDiv();
            var isEdit = "N";
            if (sessionStorage.getItem("bsEdit")) {
            	isEdit = "Y";
            }
            var afterSketchModel = Spine.Model.sub();
            afterSketchModel.configure("/user/saveAfterSketchDiagram", "createdBy", "jsonString", "jobReferenceString", 
            		"snapShot", "jobTitle", "roleJson", "divHeight", "divWidth", "totalCount", "divInitialVal", 
            		"timeSpents", "isCompleteds", "profileId", "positionJson", "isEdits", "linkClicked", 
            		"rowId", "isPrevious", "initialSave", "jctUserId", "elementOutsideRoleJson");
            afterSketchModel.extend( Spine.Model.Ajax );
            var isCompleted = "N";
            if(sessionStorage.getItem("isLogout") == "Y"){
            	isCompleted = "Y";
            }
            //Populate the model with data to transfer
        	var modelPopulator = new afterSketchModel({ 
        		createdBy: sessionStorage.getItem("jctEmail"), 
        		jsonString: totalJsonAftSktch,
        		jobReferenceString: sessionStorage.getItem("jobReferenceNo"),
        		snapShot: snapShotURL,
        		jobTitle: sessionStorage.getItem("jobTitle"),
        		roleJson: jQuery.parseJSON(sessionStorage.getItem("define_role")),
        		divHeight: outerDivHeight,
        		divWidth: outerDivWidth,
        		totalCount: sessionStorage.getItem("total_count"),
        		divInitialVal: JSON.parse(sessionStorage.getItem("div_intial_value")),
        		timeSpents : timeSpent,
        		isCompleteds: "NA",
        		profileId: sessionStorage.getItem("profileId"),
        		positionJson: sessionStorage.getItem("position_json"),
        		isEdits: isEdit,
        		linkClicked: sessionStorage.getItem("linkClicked"),
        		rowId: sessionStorage.getItem("rowIdentity"),
        		isPrevious: "Y",
        		initialSave: "N",
        		jctUserId: sessionStorage.getItem("jctUserId"),
        		elementOutsideRoleJson: jQuery.parseJSON(sessionStorage.getItem("element_outside_role").replace(/\?/g, ";_;"))
        	});
        	modelPopulator.save(); //POST
        	
        	afterSketchModel.bind("ajaxError", function(record, xhr, settings, error) {
        		//alert(JSON.stringify(xhr));
        		$(".loader_bg").fadeOut();
        	    $(".loader").hide();
        	});
        	
        	afterSketchModel.bind("ajaxSuccess", function(record, xhr, settings, error) {
        		//removePresentSessionItems();
        		sessionStorage.setItem("as2View","H");
        		sessionStorage.setItem("pageSequence", 3);
        		if(null != sessionStorage.getItem("total_json_obj")){
        			var navigation = Spine.Model.sub();
        			navigation.configure("/user/navigation/goPrevious", "navigationSeq", "jobReferenceNos", "profileId", "linkClicked", "rowId");
        			navigation.extend( Spine.Model.Ajax );
        			//Populate the model
        			var replyModel = new navigation({
        				navigationSeq: "3",
        				jobReferenceNos: sessionStorage.getItem("jobReferenceNo"),
        				profileId: sessionStorage.getItem("profileId"),
        				linkClicked: sessionStorage.getItem("linkClicked"),
        				rowId: sessionStorage.getItem("rowIdentity")
        			});
        			replyModel.save(); //POST
        			
        			navigation.bind("ajaxError", function(record, xhr, settings, error) {
        				alertify.alert("Unable to connect to the server!");
        				return false;
        			});
        			
        			navigation.bind("ajaxSuccess", function(record, xhr, settings, error) {
        				var jsonStr = JSON.stringify(xhr);
        				var obj = jQuery.parseJSON(jsonStr);
        				if(sessionStorage.getItem("isLogout") == "N"){
        					sessionStorage.setItem("back_to_previous", 1);
        					sessionStorage.setItem("strength",JSON.stringify(obj.strengthMap));
        					sessionStorage.setItem("passion",JSON.stringify(obj.passionMap));
        					sessionStorage.setItem("value",JSON.stringify(obj.valueMap));
        					sessionStorage.setItem("checked_element", obj.afterSketchCheckedEle);
        					sessionStorage.setItem("total_json_obj", obj.afterSkPageOneTotalJson);
        					sessionStorage.setItem("strength_count", obj.strengthCount);
        					sessionStorage.setItem("passion_count", obj.passionCount);
        					sessionStorage.setItem("value_count", obj.valueCount);
        					sessionStorage.setItem("checked_passion", obj.checkedPassion);
        					sessionStorage.setItem("checked_strength", obj.checkedStrength);
        					sessionStorage.setItem("checked_value", obj.checkedValue);
        					sessionStorage.setItem("isCompleted", obj.isCompleted);
        					window.location = "afterSketch1.jsp";
        				}
        			});
        		} else {
        			window.location = "afterSketch1.jsp";
        		}

        	});
            }
        });	
}


//aftersketch
if(ClientSession.get("total_json_after_sketch_final") != null) {
	var insDisplay = sessionStorage.getItem("as2View");
	if(null == insDisplay || insDisplay == "" || insDisplay == "D"){
		document.getElementById('panel').style.display = "block";
		document.getElementById('drp').setAttribute("class", "btn-slide active");
	} else {
		document.getElementById('panel').style.display = "none";
		document.getElementById('drp').setAttribute("class", "btn-slide");
	}
	/******** Added to fix Issue#128 *********/
	if (sessionStorage.getItem("bsEdit")) {
		$("#addtask").addClass("disable_button");
		var jsonObjAftrSktc = JSON.parse(ClientSession.get("total_json_after_sketch_final"));				
		var jsonLength = jsonObjAftrSktc.length;
		//$('#outderDiv').empty();
		for (var i = 0; i<jsonLength;i++){	
		var divId = jsonObjAftrSktc[i]["divId"];
		if(divId.search("divImg_") != -1) {
			var position = "";
			var divStyleDiv = "";
			var divIdx = jsonObjAftrSktc[i]["divIdx"];
			var energyId = jsonObjAftrSktc[i]["energyId"];
			var energyValue = jsonObjAftrSktc[i]["energyValue"];
			var divStyle = jsonObjAftrSktc[i]["style"];
			if((divStyle == null) || (divStyle == "null") || (divStyle.search('block') != -1)){
				divStyleDiv = "block;";
			}else{
				divStyleDiv = "none;";
			}		
			/******* ISSUE 162 START*******/
			if((divStyle.search('absolute') != -1)){
				position = "absolute;";
			}else{
				position = "relative";
			}	
			/******* ISSUE 162 END*******/
			var areaId = jsonObjAftrSktc[i]["areaId"];
			var areaValue = jsonObjAftrSktc[i]["areaValue"];
			areaValue = areaValue.replace(/`/g, ",");
			areaValue = areaValue.replace(/-_-/g, "#");
			areaValue = areaValue.replace(/;_;/g, "?");
			var additionalId = jsonObjAftrSktc[i]["additionalId"];
			var additionalValue = jsonObjAftrSktc[i]["additionalValue"];
			additionalValue = additionalValue.replace(/`/g, ",");
			additionalValue = additionalValue.replace(/-_-/g, "#");
			additionalValue = additionalValue.replace(/;_;/g, "?");
			var positionLeft = jsonObjAftrSktc[i]["positionLeft"];
			var positionTop = jsonObjAftrSktc[i]["positionTop"];
			var positionWidth = jsonObjAftrSktc[i]["positionWidth"];
			var positionHeight = jsonObjAftrSktc[i]["positionHeight"];
			
			var aspectRatio = jsonObjAftrSktc[i]["aspectRatio"];
			var deviceWidth = $("#outderDiv").width();
			var deviceHeight = getActalHeight(aspectRatio, deviceWidth);
			$("#pageWrapCopy").css({'height': deviceHeight+'px'});
			var positionLeftPercent = jsonObjAftrSktc[i]["positionLeftPercent"];
			var positionLeftVal = getActualValue(deviceWidth, positionLeftPercent);
			var positionTopPercent = jsonObjAftrSktc[i]["positionTopPercent"];
			var positionTopVal = getActualValue(deviceHeight, positionTopPercent);
			var positionWidthPercent = jsonObjAftrSktc[i]["positionWidthPercent"];
			var positionWidthVal = getActualValue(deviceWidth, positionWidthPercent);
			
			/****  Added for lock and unlock feature *****/
			var divLockId = jsonObjAftrSktc[i]["divLockId"];
			var lockImage = jsonObjAftrSktc[i]["lockImage"];
			/******* END *******/
						
			if(navigator.userAgent.match(/Android/i) || navigator.userAgent.match(/iPhone/i)) {
				var $element = $('<div id="' + divId +'" class="col-md-4 after_sketch_task single_task_item_sketch draggableResizable common_elem" style="position:'+position+'; top:'+positionTopVal+'px; left:'+positionLeftVal+'px; width:'+positionWidthVal+'px ;display:'+divStyleDiv+'; height:'+positionWidthVal+'px;"><div class="col-md-12 task_item_top_sketch"><div class="col-md-3 col-xs-3 task">TASK</div><div class="col-md-6 col-xs-5 energy" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/ Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" readonly="true" /></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm_sketch"><textarea class="col-md-12 col-xs-12 add_task_area_sketch" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" maxlength="200" placeholder="Add Task Label Here" onkeypress="return disableKey(event)" onBlur="disableDraggableResizable()">'+areaValue+'</textarea></div><div class="col-md-12 additional_area"><img class="person_icon" src="../img/personIcon.png"><textarea class="form-control-additional input-sm" id="'+additionalId+'" onkeypress="return disableKey(event)" name="additional" maxlength="100" >'+additionalValue+'</textarea></div><div id="' + divIdx +'" class="cross"></div><div id="' + divLockId +'" class="lock_unlock none_display"><img src="'+lockImage+'" alt="Lock" /></div></div>');
			} else if(navigator.userAgent.match(/iPad/i) || navigator.userAgent.match(/IPad/i)) {
				var $element = $('<div id="' + divId +'" class="col-md-4 after_sketch_task single_task_item_sketch draggableResizable common_elem" style="position:'+position+'; top:'+positionTopVal+'px; left:'+positionLeftVal+'px; width:'+positionWidthVal+'px ;display:'+divStyleDiv+'; height:'+positionWidthVal+'px;"><div class="col-md-12 task_item_top_sketch"><div class="col-md-3 col-xs-3 task">TASK</div><div class="col-md-6 col-xs-5 energy" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" readonly="true" /></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm_sketch"><textarea class="col-md-12 col-xs-12 add_task_area_sketch" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)" onBlur="disableDraggableResizable()">'+areaValue+'</textarea></div><div class="col-md-12 additional_area"><img class="person_icon" src="../img/personIcon.png"><textarea class="form-control-additional input-sm" id="'+additionalId+'" onkeypress="return disableKey(event)" name="additional" maxlength="100" placeholder="Add Relational Crafting Here" onblur="disableDraggableResizable()">'+additionalValue+'</textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock none_display"><img src="'+lockImage+'" alt="Lock" /></div></div>');
			} else {
				var $element = $('<div id="' + divId +'" class="col-md-4 after_sketch_task single_task_item_sketch draggableResizable common_elem" style="position:'+position+'; top:'+positionTopVal+'px; left:'+positionLeftVal+'px; width:'+positionWidthVal+'px ;display:'+divStyleDiv+'; height:'+positionWidthVal+'px;"><div class="col-md-12 task_item_top_sketch"><div class="col-md-3 col-xs-3 task">TASK</div><div class="col-md-6 col-xs-5 energy">Time/Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" readonly="true" /></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm_sketch"><textarea class="col-md-12 col-xs-12 add_task_area_sketch" style="text-align: center;" id="'+areaId+'" type="text" onkeypress="return disableKey(event)" name="areaName" maxlength="200" placeholder="Add Task Label Here">'+areaValue+'</textarea></div><div class="col-md-12 additional_area"><img class="person_icon" src="../img/personIcon.png"><textarea class="form-control-additional input-sm" id="'+additionalId+'" onkeypress="return disableKey(event)" name="additional" maxlength="100" >'+additionalValue+'</textarea></div><div id="' + divIdx +'" class="cross"></div><div id="' + divLockId +'" class="lock_unlock none_display"><img src="'+lockImage+'" alt="Lock" /></div></div>');
			}			
			$("#outderDiv").append($element);
			if(energyValue >=5 && energyValue <=9){					
				$("#" +divId).children().find(".energy_field").children().addClass("as_small_remove_text");	
			} else {
				$("#" +divId).children().find(".energy_field").children().addClass("as_remove_input_text");
			}
			/***************** To Disable the paste functionality ISSUE# 186***************/
		    $('#' +areaId).bind("paste",function(e) {
		       e.preventDefault();
		    });
		    //setClassAgainstLockUnlockElememt(divId, divLockId, lockImage);		  					
		}
		
		if(divId.search("divStrengthClone_") != -1) {		
			var textId = jsonObjAftrSktc[i]["textId"];	
			var divValue = jsonObjAftrSktc[i]["divValue"];		
			var positionLeft = jsonObjAftrSktc[i]["positionLeft"];
			var positionTop = jsonObjAftrSktc[i]["positionTop"];	
			var divStyle = jsonObjAftrSktc[i]["style"];
			
			var aspectRatio = jsonObjAftrSktc[i]["aspectRatio"];
			var deviceWidth = $("#outderDiv").width();
			var deviceHeight = getActalHeight(aspectRatio, deviceWidth);
			var positionLeftPercent = jsonObjAftrSktc[i]["positionLeftPercent"];
			var positionLeftVal = getActualValue(deviceWidth, positionLeftPercent);
			var positionTopPercent = jsonObjAftrSktc[i]["positionTopPercent"];
			var positionTopVal = getActualValue(deviceHeight, positionTopPercent);
			
			var deleteElementId = divId+"_x";     	
			if((divStyle == null) || (divStyle == "null") || (divStyle.search('block') != -1)){
				divStyle = "block;";
			}else{
				divStyle = "none;";
			}
			if(navigator.userAgent.match(/Android/i) || navigator.userAgent.match(/iPhone/i)) {
				var $element = $('<div id="' + divId +'" class="plus_sign_small commonDrag commonDelete common_elem" style="position: absolute; top:'+positionTopVal+'px; left:'+positionLeftVal+'px; display:'+divStyle+'"><img src="../img/strength.png" width="50" height="50" style="no-repeat 0 9px;"><div id="' + deleteElementId +'" class="removeService delete_element"></div><div class="plus_sign_small_span" id="'+textId+'"><p>'+divValue+'</p></div>');
			} else if(navigator.userAgent.match(/iPad/i) || navigator.userAgent.match(/IPad/i)) {
				var $element = $('<div id="' + divId +'" class="plus_sign_small commonDrag commonDelete common_elem" style="position: absolute; top:'+positionTopVal+'px; left:'+positionLeftVal+'px; display:'+divStyle+'"><img src="../img/strength.png" width="80" height="80" style="no-repeat 0 9px;"><div id="' + deleteElementId +'" class="removeService delete_element"></div><div class="plus_sign_small_span" id="'+textId+'"><p>'+divValue+'</p></div>');
			} else {
				var $element = $('<div id="' + divId +'" class="plus_sign_small commonDrag commonDelete common_elem" style="position: absolute; top:'+positionTopVal+'px; left:'+positionLeftVal+'px; display:'+divStyle+'"><img src="../img/strength.png" width="100" height="80" style="no-repeat 0 9px;"><div id="' + deleteElementId +'" class="removeService delete_element"></div><div class="plus_sign_small_span" id="'+textId+'"><p>'+divValue+'</p></div>');
			}								
			$("#outderDiv").append($element);
			$(".plus_sign_small").css({ cursor: "default" });
		}
		
		if(divId.search("divValueClone_") != -1) {
			var textId = jsonObjAftrSktc[i]["textId"];	
			var divValue = jsonObjAftrSktc[i]["divValue"];		
			var positionLeft = jsonObjAftrSktc[i]["positionLeft"];
			var positionTop = jsonObjAftrSktc[i]["positionTop"];	
			var divStyle = jsonObjAftrSktc[i]["style"];
			
			var aspectRatio = jsonObjAftrSktc[i]["aspectRatio"];
			var deviceWidth = $("#outderDiv").width();
			var deviceHeight = getActalHeight(aspectRatio, deviceWidth);
			var positionLeftPercent = jsonObjAftrSktc[i]["positionLeftPercent"];
			var positionLeftVal = getActualValue(deviceWidth, positionLeftPercent);
			var positionTopPercent = jsonObjAftrSktc[i]["positionTopPercent"];
			var positionTopVal = getActualValue(deviceHeight, positionTopPercent);
			
			var deleteElementId = divId+"_x"; 
			if((divStyle == null) || (divStyle == "null") || (divStyle.search('block') != -1)){
				divStyle = "block;";
			}else{
				divStyle = "none;";
			}
			if(navigator.userAgent.match(/Android/i) || navigator.userAgent.match(/iPhone/i)) {
				var $element = $('<div id="' + divId +'" class="oval_small commonDrag commonDelete common_elem" style="position: absolute; top:'+positionTopVal+'px; left:'+positionLeftVal+'px; display:'+divStyle+'"><img src="../img/value.png" width="55" height="45" style="no-repeat 0 9px;"><div id="' + deleteElementId +'" class="removeService delete_element"></div><div class="oval_small_span" id="'+textId+'"><p>'+divValuea+'</p></div>');
			} else if(navigator.userAgent.match(/iPad/i) || navigator.userAgent.match(/IPad/i)) {
				var $element = $('<div id="' + divId +'" class="oval_small commonDrag commonDelete common_elem" style="position: absolute; top:'+positionTopVal+'px; left:'+positionLeftVal+'px; display:'+divStyle+'"><img src="../img/value.png" width="85" height="76" style="no-repeat 0 9px;"><div id="' + deleteElementId +'" class="removeService delete_element"></div><div class="oval_small_span" id="'+textId+'"><p>'+divValuea+'</p></div>');
			} else {
				var $element = $('<div id="' + divId +'" class="oval_small commonDrag commonDelete common_elem" style="position: absolute; top:'+positionTopVal+'px; left:'+positionLeftVal+'px; display:'+divStyle+'"><img src="../img/value.png" width="105" height="85" style="no-repeat 0 9px;"><div id="' + deleteElementId +'" class="removeService delete_element"></div><div class="oval_small_span" id="'+textId+'"><p>'+divValuea+'</p></div>');
			}								
			$("#outderDiv").append($element);
			$(".oval_small").css({ cursor: "default" });
		}
		
		if(divId.search("divPassionClone_") != -1) {
			var textId = jsonObjAftrSktc[i]["textId"];	
			var divValue = jsonObjAftrSktc[i]["divValue"];		
			var positionLeft = jsonObjAftrSktc[i]["positionLeft"];
			var positionTop = jsonObjAftrSktc[i]["positionTop"];	
			var divStyle = jsonObjAftrSktc[i]["style"];
			
			var aspectRatio = jsonObjAftrSktc[i]["aspectRatio"];
			var deviceWidth = $("#outderDiv").width();
			var deviceHeight = getActalHeight(aspectRatio, deviceWidth);
			var positionLeftPercent = jsonObjAftrSktc[i]["positionLeftPercent"];
			var positionLeftVal = getActualValue(deviceWidth, positionLeftPercent);
			var positionTopPercent = jsonObjAftrSktc[i]["positionTopPercent"];
			var positionTopVal = getActualValue(deviceHeight, positionTopPercent);
			
			var deleteElementId = divId+"_x"; 
			if((divStyle == null) || (divStyle == "null") || (divStyle.search('block') != -1)){
				divStyle = "block;";
			}else{
				divStyle = "none;";
			}					
			if(navigator.userAgent.match(/Android/i) || navigator.userAgent.match(/iPhone/i)) {
				var $element = $('<div id="' + divId +'" class="up_triangle_small commonDrag commonDelete common_elem" style="position: absolute; top:'+positionTopVal+'px; left:'+positionLeftVal+'px; display:'+divStyle+'"><img src="../img/passion.png" width="50" height="50" style="no-repeat 0 9px;"><div id="' + deleteElementId +'" class="removeService delete_element"></div><div class="up_triangle_small_span" id="'+textId+'"><p>'+divValue+'</p></div>');
			} else if(navigator.userAgent.match(/iPad/i) || navigator.userAgent.match(/IPad/i)) {
				var $element = $('<div id="' + divId +'" class="up_triangle_small commonDrag commonDelete common_elem" style="position: absolute; top:'+positionTopVal+'px; left:'+positionLeftVal+'px; display:'+divStyle+'"><img src="../img/passion.png" width="87" height="82" style="no-repeat 0 9px;"><div id="' + deleteElementId +'" class="removeService delete_element"></div><div class="up_triangle_small_span" id="'+textId+'"><p>'+divValue+'</p></div>');
			} else {
				var $element = $('<div id="' + divId +'" class="up_triangle_small commonDrag commonDelete common_elem" style="position: absolute; top:'+positionTopVal+'px; left:'+positionLeftVal+'px; display:'+divStyle+'"><img src="../img/passion.png" width="100" height="85" style="no-repeat 0 9px;"><div id="' + deleteElementId +'" class="removeService delete_element"></div><div class="up_triangle_small_span" id="'+textId+'"><p>'+divValue+'</p></div>');
			}
			$("#outderDiv").append($element);
			$(".up_triangle_small").css({ cursor: "default" });
		}
		
		if(divId.search("cloneOval_") != -1) {
			var roleValue = jsonObjAftrSktc[i]["roleValue"];	
			roleValue = roleValue.replace(/-_-/g, "#");
			roleValue = roleValue.replace(/;_;/g, "?");
			var positionLeft = jsonObjAftrSktc[i]["positionLeft"];
			var positionTop = jsonObjAftrSktc[i]["positionTop"];	
			var positionWidth = jsonObjAftrSktc[i]["positionWidth"];	
			var positionHeight = jsonObjAftrSktc[i]["positionHeight"];
			var divStyle = jsonObjAftrSktc[i]["style"];
			var deleteElementId = divId+"_x"; 
			//var type = jsonObjAftrSktc[i]["type"];			
			var divStyleFromJSON = divStyle;						
			if((divStyle == null) || (divStyle == "null") || (divStyle.search('block') != -1)){
				divStyleb = "block;";
				if (windowSize <= 1920) {
				//MOBILE - START
				var arr = divStyle.split(";");
				for (var index = 0; index < arr.length; index ++ ) {
					var innerString = arr[index];
					if (innerString.search("top") != -1) {
						positionTop = innerString.split(":")[1];
					}
				}
				for (var index = 0; index < arr.length; index ++ ) {
					var innerString = arr[index];
					if (innerString.search("left") != -1) {
						positionLeft = innerString.split(":")[1];
					}
				}
				//MOBILE - END
				}	
			}else{
				divStyleb = "none;";
			}
			
			var aspectRatio = jsonObjAftrSktc[i]["aspectRatio"];
			var deviceWidth = $("#outderDiv").width();
			var deviceHeight = getActalHeight(aspectRatio, deviceWidth);
			var positionLeftPercent = jsonObjAftrSktc[i]["positionLeftPercent"];
			var positionLeftVal = getActualValue(deviceWidth, positionLeftPercent);
			var positionTopPercent = jsonObjAftrSktc[i]["positionTopPercent"];
			var positionTopVal = getActualValue(deviceHeight, positionTopPercent);
			
			var positionWidthPercent = jsonObjAftrSktc[i]["positionWidthPercent"];
			var positionWidthVal = getActualValue(deviceWidth, positionWidthPercent);
			var positionHeightPercent = jsonObjAftrSktc[i]["positionHeightPercent"];
			var positionHeightVal = getActualValue(deviceWidth, positionHeightPercent);						
			
			var img = jsonObjAftrSktc[i]["roleImage"];			
			var balanceStyleClass = "ui-widget-content ui-draggable oval_cloned_div common_elem commonDelete horizontal_oval_active ui-resizable";
			
			var aspectRatio = jsonObjAftrSktc[i]["aspectRatio"];
			var deviceWidth = $("#outderDiv").width();
			var deviceHeight = getActalHeight(aspectRatio, deviceWidth);
			var positionLeftPercent = jsonObjAftrSktc[i]["positionLeftPercent"];
			var positionLeftVal = getActualValue(deviceWidth, positionLeftPercent);
			var positionTopPercent = jsonObjAftrSktc[i]["positionTopPercent"];
			var positionTopVal = getActualValue(deviceHeight, positionTopPercent);
			
			var positionWidthPercent = jsonObjAftrSktc[i]["positionWidthPercent"];
			var positionWidthVal = getActualValue(deviceWidth, positionWidthPercent);
			var positionHeightPercent = jsonObjAftrSktc[i]["positionHeightPercent"];
			var positionHeightVal = getActualValue(deviceWidth, positionHeightPercent);
			var $element = $('<div id="' + divId +'" class="'+balanceStyleClass+'" style="position: absolute; top:'+positionTopVal+'px; left:'+positionLeftVal+'px; display:'+divStyle+'; width:'+positionWidthVal+'px; height:'+positionHeightVal+'px;"><img src="'+img+'" width="100%" draggable="false"/></div>');
			$("#outderDiv").append($element);
			
			$(".oval_cloned_div").css({ cursor: "default" });
			/*$( "#" + divId ).resize(function() {
				checkResizeCordinates(divId);
			});*/
			
			var roleId = "role_" +divId;
			var textAreaId = "text_" +divId;
			var roleLeftPercent = jsonObjAftrSktc[i]["roleLeftPercent"];
			var roleLeftVal = getActualValue(deviceWidth, roleLeftPercent);
			var roleTopPercent = jsonObjAftrSktc[i]["roleTopPercent"];
			var roleTopVal = getActualValue(deviceHeight, roleTopPercent);
			
			var roleWidthPercent = jsonObjAftrSktc[i]["roleWidthPercent"];
			var roleWidthVal = getActualValue(deviceWidth, roleWidthPercent);
			var roleHeightPercent = jsonObjAftrSktc[i]["roleHeightPercent"];
			var roleHeightVal = getActualValue(deviceWidth, roleHeightPercent);
			var $elementRole = $('<div class="speech-bubble common_elem commonDelete" id="'+ roleId +'" style="position: absolute; top:'+roleTopVal+'px; left:'+roleLeftVal+'px; display:'+divStyle+'; width:'+roleWidthVal+'px; height:'+roleHeightVal+'px;"><div class="role_text" >Role</div><div id= "clickableElement1" class="clickable_element_1 hide_image"><img src=""></div><div id= "clickableElement2" class="clickable_element_2 hide_image" ><img src=""></div><textarea id = "'+textAreaId+'" onkeypress="return disableKey(event)" maxlength="120">'+roleValue+'</textarea><div id="' + deleteElementId +'" class="removeService delete_oval" onclick=deleteElement(this)><img src="../img/cross-black.png" alt="Delete" /></div></div>');

			$("#outderDiv").append($elementRole);

			setResizableCorner(img, roleId);
			/*Adjusting task area on load for devices
			 */
			var container = document.getElementById(textAreaId);
			if(navigator.userAgent.match(/Android/i) || navigator.userAgent.match(/iPhone/i) || navigator.userAgent.match(/iPad/i) || navigator.userAgent.match(/IPad/i)) {
				var insDisplay = sessionStorage.getItem("as2View");
				if(null == insDisplay || insDisplay == "" || insDisplay == "D"){
					reDragRoleFrame(divId);
					while($("#"+textAreaId).outerHeight() < container.scrollHeight + parseFloat(($("#"+textAreaId)).css("borderTopWidth")) + parseFloat(($("#"+textAreaId)).css("borderBottomWidth"))) {
						fixRoleFrameText(divId);
					}
				}
			}
		}
		}
		/*************************To fix the design ******************************/
		fixDesignAS();		
		setTimeout(function(){getOuterDivHeight();},200);
	} else {
		$(".loader_bg").fadeIn();
		$(".loader").fadeIn();
		var jsonObjAftrSktc = JSON.parse(ClientSession.get("total_json_after_sketch_final"));			
		var jsonLength = jsonObjAftrSktc.length;
		for (var i = 0; i<jsonLength;i++){	
			var divId = jsonObjAftrSktc[i]["divId"];
			if(divId.search("divImg_") != -1) {
				var position = "";
				var divStyleDiv = "";
				var divIdx = jsonObjAftrSktc[i]["divIdx"];
				var energyId = jsonObjAftrSktc[i]["energyId"];
				var energyValue = jsonObjAftrSktc[i]["energyValue"];
				var divStyle = jsonObjAftrSktc[i]["style"];
				if((divStyle == null) || (divStyle == "null") || (divStyle.search('block') != -1)){
					divStyleDiv = "block;";
				}else{
					divStyleDiv = "none;";
				}
				/******* ISSUE 162 START*******/
				if((divStyle.search('absolute') != -1)){
					position = "absolute;";
				}else{
					position = "relative";
				}	
				/******* ISSUE 162 END*******/
				var areaId = jsonObjAftrSktc[i]["areaId"];
				var areaValue = jsonObjAftrSktc[i]["areaValue"];
				areaValue = areaValue.replace(/`/g, ",");
				areaValue = areaValue.replace(/-_-/g, "#");
				areaValue = areaValue.replace(/;_;/g, "?");
				var additionalId = jsonObjAftrSktc[i]["additionalId"];
				var additionalValue = jsonObjAftrSktc[i]["additionalValue"];
				additionalValue = additionalValue.replace(/`/g, ",");
				additionalValue = additionalValue.replace(/-_-/g, "#");
				additionalValue = additionalValue.replace(/;_;/g, "?");
				var positionLeft = jsonObjAftrSktc[i]["positionLeft"];
				var positionTop = jsonObjAftrSktc[i]["positionTop"];
				var positionWidth = jsonObjAftrSktc[i]["positionWidth"];
				var positionHeight = jsonObjAftrSktc[i]["positionHeight"];
				
				var aspectRatio = jsonObjAftrSktc[i]["aspectRatio"];
				var deviceWidth = $("#outderDiv").width();
				var deviceHeight = getActalHeight(aspectRatio, deviceWidth);
				var positionLeftPercent = jsonObjAftrSktc[i]["positionLeftPercent"];
				var positionLeftVal = getActualValue(deviceWidth, positionLeftPercent);
				var positionTopPercent = jsonObjAftrSktc[i]["positionTopPercent"];
				var positionTopVal = getActualValue(deviceHeight, positionTopPercent);
				var positionWidthPercent = jsonObjAftrSktc[i]["positionWidthPercent"];
				var positionWidthVal = getActualValue(deviceWidth, positionWidthPercent);
				
				/****  Added for lock and unlock feature *****/
				var divLockId = jsonObjAftrSktc[i]["divLockId"];
				var lockImage = jsonObjAftrSktc[i]["lockImage"];
				/******* END *******/								
				if(navigator.userAgent.match(/Android/i) || navigator.userAgent.match(/iPhone/i)) {
					var $element = $('<div id="' + divId +'" class="col-md-4 after_sketch_task single_task_item_sketch draggableResizable common_elem" style="position:'+position+'; top:'+positionTopVal+'px; left:'+positionLeftVal+'px; width:'+positionWidthVal+'px ;display:'+divStyleDiv+'; height:'+positionWidthVal+'px;"><div class="col-md-12 task_item_top_sketch"><div class="col-md-3 col-xs-3 task" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 energy" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/ Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onchange="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm_sketch"><textarea class="col-md-12 col-xs-12 add_task_area_sketch" style="text-align: center;" id="'+areaId+'" type="text" onkeypress="return disableKey(event)" name="areaName" placeholder="Add Task Label Here" maxlength="200" onBlur="disableDraggableResizable()">'+areaValue+'</textarea></div><div class="col-md-12 additional_area"><img class="person_icon" src="../img/personIcon.png"><textarea class="form-control-additional input-sm" id="'+additionalId+'" onkeypress="return disableKey(event)" name="additional" maxlength="120" placeholder="Add Relational Crafting Here" onBlur="disableDraggableResizable()">'+additionalValue+'</textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divLockId+'")><img src="'+lockImage+'" alt="Lock" class="lock_unlock_img"/></div></div>');
				} else if(navigator.userAgent.match(/iPad/i) || navigator.userAgent.match(/IPad/i)) {
					var $element = $('<div id="' + divId +'" class="col-md-4 after_sketch_task single_task_item_sketch draggableResizable common_elem" style="position:'+position+'; top:'+positionTopVal+'px; left:'+positionLeftVal+'px; width:'+positionWidthVal+'px ;display:'+divStyleDiv+'; height:'+positionWidthVal+'px;"><div class="col-md-12 task_item_top_sketch"><div class="col-md-3 col-xs-3 task" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 energy" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onchange="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm_sketch"><textarea class="col-md-12 col-xs-12 add_task_area_sketch" onkeypress="return disableKey(event)" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" placeholder="Add Task Label Here" maxlength="200" onBlur="disableDraggableResizable()">'+areaValue+'</textarea></div><div class="col-md-12 additional_area"><img class="person_icon" src="../img/personIcon.png"><textarea class="form-control-additional input-sm" id="'+additionalId+'" onkeypress="return disableKey(event)" name="additional" maxlength="120" placeholder="Add Relational Crafting Here" onBlur="disableDraggableResizable()">'+additionalValue+'</textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divLockId+'")><img src="'+lockImage+'" alt="Lock" class="lock_unlock_img"/></div></div>');
				} else {
					var $element = $('<div id="' + divId +'" class="col-md-4 after_sketch_task single_task_item_sketch draggableResizable common_elem" style="position:'+position+'; top:'+positionTopVal+'px; left:'+positionLeftVal+'px; width:'+positionWidthVal+'px ;display:'+divStyleDiv+'; height:'+positionWidthVal+'px;"><div class="col-md-12 task_item_top_sketch"><div class="col-md-3 col-xs-3 task">TASK</div><div class="col-md-6 col-xs-5 energy">Time/Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onchange="return getEnergyValue(this)" /></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm_sketch"><textarea class="col-md-12 col-xs-12 add_task_area_sketch" onkeypress="return disableKey(event)" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" placeholder="Add Task Label Here" maxlength="200">'+areaValue+'</textarea></div><div class="col-md-12 additional_area"><img class="person_icon" src="../img/personIcon.png"><textarea class="form-control-additional input-sm" id="'+additionalId+'" onkeypress="return disableKey(event)" name="additional" maxlength="120" placeholder="Add Relational Crafting Here">'+additionalValue+'</textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divLockId+'")><img src="'+lockImage+'" alt="Lock" class="lock_unlock_img"/></div></div>');
				}	
				$("#outderDiv").append($element);
	
				/***************** To Disable the paste functionality ISSUE# 186***************/
			    $('#' +areaId).bind("paste",function(e) {
			       e.preventDefault();
			    });
			    $element.draggable();						
			    $element.resizable({		
					//ghost: 'true',
					handles: 'se',
					aspectRatio: 2 / 2,
				});	
				/**************** To reduce the Z-index value to fix the UI**********************/
				$(".single_task_item_sketch .ui-resizable-se").css({ 'z-index':'0'});
				setClassAgainstLockUnlockElememt(divId, divLockId, lockImage);										
			}
			
			if(divId.search("divStrengthClone_") != -1) {	
				var textId = jsonObjAftrSktc[i]["textId"];	
				var divValue = jsonObjAftrSktc[i]["divValue"];		
				var positionLeft = jsonObjAftrSktc[i]["positionLeft"];
				var positionTop = jsonObjAftrSktc[i]["positionTop"];	
				var divStyle = jsonObjAftrSktc[i]["style"];
				
				var aspectRatio = jsonObjAftrSktc[i]["aspectRatio"];
				var deviceWidth = $("#outderDiv").width();
				var deviceHeight = getActalHeight(aspectRatio, deviceWidth);
				var positionLeftPercent = jsonObjAftrSktc[i]["positionLeftPercent"];
				var positionLeftVal = getActualValue(deviceWidth, positionLeftPercent);
				var positionTopPercent = jsonObjAftrSktc[i]["positionTopPercent"];
				var positionTopVal = getActualValue(deviceHeight, positionTopPercent);
				
				var deleteElementId = divId+"_x";     	
				if((divStyle == null) || (divStyle == "null") || (divStyle.search('block') != -1)){
					divStyle = "block;";
				}else{
					divStyle = "none;";
				}
				if(navigator.userAgent.match(/Android/i) || navigator.userAgent.match(/iPhone/i)) {
					var $element = $('<div id="' + divId +'" class="plus_sign_small commonDrag commonDelete common_elem" style="position: absolute; top:'+positionTopVal+'px; left:'+positionLeftVal+'px; display:'+divStyle+'"><img src="../img/strength.png" width="50" height="50" style="no-repeat 0 9px;"><div id="' + deleteElementId +'" class="removeService delete_element" onclick=deleteElement(this)><img src="../img/cross-black.png" alt="Delete" /></div><div class="plus_sign_small_span" id="'+textId+'"><p>'+divValue+'</p></div>');
				} else if(navigator.userAgent.match(/iPad/i) || navigator.userAgent.match(/IPad/i)) {
					var $element = $('<div id="' + divId +'" class="plus_sign_small commonDrag commonDelete common_elem" style="position: absolute; top:'+positionTopVal+'px; left:'+positionLeftVal+'px; display:'+divStyle+'"><img src="../img/strength.png" width="80" height="80" style="no-repeat 0 9px;"><div id="' + deleteElementId +'" class="removeService delete_element" onclick=deleteElement(this)><img src="../img/cross-black.png" alt="Delete" /></div><div class="plus_sign_small_span" id="'+textId+'"><p>'+divValue+'</p></div>');
				} else {
					var $element = $('<div id="' + divId +'" class="plus_sign_small commonDrag commonDelete common_elem" style="position: absolute; top:'+positionTopVal+'px; left:'+positionLeftVal+'px; display:'+divStyle+'"><img src="../img/strength.png" width="100" height="80" style="no-repeat 0 9px;"><div id="' + deleteElementId +'" class="removeService delete_element" onclick=deleteElement(this)><img src="../img/cross-black.png" alt="Delete" /></div><div class="plus_sign_small_span" id="'+textId+'"><p>'+divValue+'</p></div>');
				}
			
				$("#outderDiv").append($element);
				$element.draggable();
			}
			
			if(divId.search("divValueClone_") != -1) {
				var textId = jsonObjAftrSktc[i]["textId"];	
				var divValue = jsonObjAftrSktc[i]["divValue"];		
				var positionLeft = jsonObjAftrSktc[i]["positionLeft"];
				var positionTop = jsonObjAftrSktc[i]["positionTop"];	
				var divStyle = jsonObjAftrSktc[i]["style"];
				
				var aspectRatio = jsonObjAftrSktc[i]["aspectRatio"];
				var deviceWidth = $("#outderDiv").width();
				var deviceHeight = getActalHeight(aspectRatio, deviceWidth);
				var positionLeftPercent = jsonObjAftrSktc[i]["positionLeftPercent"];
				var positionLeftVal = getActualValue(deviceWidth, positionLeftPercent);
				var positionTopPercent = jsonObjAftrSktc[i]["positionTopPercent"];
				var positionTopVal = getActualValue(deviceHeight, positionTopPercent);
				
				var deleteElementId = divId+"_x"; 
				if((divStyle == null) || (divStyle == "null") || (divStyle.search('block') != -1)){
					divStyle = "block;";
				}else{
					divStyle = "none;";
				}
				
				if(navigator.userAgent.match(/Android/i) || navigator.userAgent.match(/iPhone/i)) {
					var $element = $('<div id="' + divId +'" class="oval_small commonDrag commonDelete common_elem" style="position: absolute; top:'+positionTopVal+'px; left:'+positionLeftVal+'px; display:'+divStyle+'"><img src="../img/value.png" width="55" height="45" style="no-repeat 0 9px;"><div id="' + deleteElementId +'" class="removeService delete_element" onclick=deleteElement(this)><img src="../img/cross-black.png" alt="Delete" /></div><div class="oval_small_span" id="'+textId+'"><p>'+divValue+'</p></div>');
				} else if(navigator.userAgent.match(/iPad/i) || navigator.userAgent.match(/IPad/i)) {					
					var $element = $('<div id="' + divId +'" class="oval_small commonDrag commonDelete common_elem" style="position: absolute; top:'+positionTopVal+'px; left:'+positionLeftVal+'px; display:'+divStyle+'"><img src="../img/value.png" width="85" height="76" style="no-repeat 0 9px;"><div id="' + deleteElementId +'" class="removeService delete_element" onclick=deleteElement(this)><img src="../img/cross-black.png" alt="Delete" /></div><div class="oval_small_span" id="'+textId+'"><p>'+divValue+'</p></div>');
				} else {
					var $element = $('<div id="' + divId +'" class="oval_small commonDrag commonDelete common_elem" style="position: absolute; top:'+positionTopVal+'px; left:'+positionLeftVal+'px; display:'+divStyle+'"><img src="../img/value.png" width="105" height="85" style="no-repeat 0 9px;"><div id="' + deleteElementId +'" class="removeService delete_element" onclick=deleteElement(this)><img src="../img/cross-black.png" alt="Delete" /></div><div class="oval_small_span" id="'+textId+'"><p>'+divValue+'</p></div>');
				}
				$("#outderDiv").append($element);
				$element.draggable();
			}
			
			if(divId.search("divPassionClone_") != -1) {
				var textId = jsonObjAftrSktc[i]["textId"];	
				var divValue = jsonObjAftrSktc[i]["divValue"];		
				var positionLeft = jsonObjAftrSktc[i]["positionLeft"];
				var positionTop = jsonObjAftrSktc[i]["positionTop"];	
				var divStyle = jsonObjAftrSktc[i]["style"];
				
				var aspectRatio = jsonObjAftrSktc[i]["aspectRatio"];
				var deviceWidth = $("#outderDiv").width();
				var deviceHeight = getActalHeight(aspectRatio, deviceWidth);
				var positionLeftPercent = jsonObjAftrSktc[i]["positionLeftPercent"];
				var positionLeftVal = getActualValue(deviceWidth, positionLeftPercent);
				var positionTopPercent = jsonObjAftrSktc[i]["positionTopPercent"];
				var positionTopVal = getActualValue(deviceHeight, positionTopPercent);
				
				var deleteElementId = divId+"_x"; 
				if((divStyle == null) || (divStyle == "null") || (divStyle.search('block') != -1)){
					divStyle = "block;";
				}else{
					divStyle = "none;";
				}
				
				if(navigator.userAgent.match(/Android/i) || navigator.userAgent.match(/iPhone/i)) {
					var $element = $('<div id="' + divId +'" class="up_triangle_small commonDrag commonDelete common_elem" style="position: absolute; top:'+positionTopVal+'px; left:'+positionLeftVal+'px; display:'+divStyle+'"><img src="../img/passion.png" width="50" height="50" style="no-repeat 0 9px;"><div id="' + deleteElementId +'" class="removeService delete_element" onclick=deleteElement(this)><img src="../img/cross-black.png" alt="Delete" /></div><div class="up_triangle_small_span" id="'+textId+'"><p>'+divValue+'</p></div>');
				} else if(navigator.userAgent.match(/iPad/i) || navigator.userAgent.match(/IPad/i)) {
					var $element = $('<div id="' + divId +'" class="up_triangle_small commonDrag commonDelete common_elem" style="position: absolute; top:'+positionTopVal+'px; left:'+positionLeftVal+'px; display:'+divStyle+'"><img src="../img/passion.png" width="87" height="82" style="no-repeat 0 9px;"><div id="' + deleteElementId +'" class="removeService delete_element" onclick=deleteElement(this)><img src="../img/cross-black.png" alt="Delete" /></div><div class="up_triangle_small_span" id="'+textId+'"><p>'+divValue+'</p></div>');
				} else {
					var $element = $('<div id="' + divId +'" class="up_triangle_small commonDrag commonDelete common_elem" style="position: absolute; top:'+positionTopVal+'px; left:'+positionLeftVal+'px; display:'+divStyle+'"><img src="../img/passion.png" width="100" height="85" style="no-repeat 0 9px;"><div id="' + deleteElementId +'" class="removeService delete_element" onclick=deleteElement(this)><img src="../img/cross-black.png" alt="Delete" /></div><div class="up_triangle_small_span" id="'+textId+'"><p>'+divValue+'</p></div>');
				}
				$("#outderDiv").append($element);
				$element.draggable();
			}
			
			if(divId.search("cloneOval_") != -1) {
				var roleValue = jsonObjAftrSktc[i]["roleValue"];	
				roleValue = roleValue.replace(/-_-/g, "#");
				roleValue = roleValue.replace(/;_;/g, "?");	
				var divStyle = jsonObjAftrSktc[i]["style"];
				var deleteElementId = divId+"_x"; 
				//var type = jsonObjAftrSktc[i]["type"];
				if((divStyle == null) || (divStyle == "null") || (divStyle.search('block') != -1)){
					divStyle = "block;";
				}else{
					divStyle = "none;";
				}
				var img = jsonObjAftrSktc[i]["roleImage"];				
				var balanceStyleClass = "ui-widget-content ui-draggable oval_cloned_div common_elem commonDelete horizontal_oval_active ui-resizable";
				
				var aspectRatio = jsonObjAftrSktc[i]["aspectRatio"];
				var deviceWidth = $("#outderDiv").width();
				var deviceHeight = getActalHeight(aspectRatio, deviceWidth);
				var positionLeftPercent = jsonObjAftrSktc[i]["positionLeftPercent"];
				var positionLeftVal = getActualValue(deviceWidth, positionLeftPercent);
				var positionTopPercent = jsonObjAftrSktc[i]["positionTopPercent"];
				var positionTopVal = getActualValue(deviceHeight, positionTopPercent);
				
				var positionWidthPercent = jsonObjAftrSktc[i]["positionWidthPercent"];
				var positionWidthVal = getActualValue(deviceWidth, positionWidthPercent);
				var positionHeightPercent = jsonObjAftrSktc[i]["positionHeightPercent"];
				var positionHeightVal = getActualValue(deviceWidth, positionHeightPercent);
				var $element = $('<div id="' + divId +'" class="'+balanceStyleClass+'" style="position: absolute; top:'+positionTopVal+'px; left:'+positionLeftVal+'px; display:'+divStyle+'; width:'+positionWidthVal+'px; height:'+positionHeightVal+'px;"><img src="'+img+'" width="100%" draggable="false"/><div class="clearfix"></div></div>');
				$("#outderDiv").append($element);
				$element.draggable().resizable();
				$( "#" + divId ).resize(function() {
					checkResizeCordinates(divId);
				});
				
				var roleId = "role_" +divId;
				var textAreaId = "text_" +divId;
				var roleLeftPercent = jsonObjAftrSktc[i]["roleLeftPercent"];
				var roleLeftVal = getActualValue(deviceWidth, roleLeftPercent);
				var roleTopPercent = jsonObjAftrSktc[i]["roleTopPercent"];
				var roleTopVal = getActualValue(deviceHeight, roleTopPercent);
				
				var roleWidthPercent = jsonObjAftrSktc[i]["roleWidthPercent"];
				var roleWidthVal = getActualValue(deviceWidth, roleWidthPercent);
				var roleHeightPercent = jsonObjAftrSktc[i]["roleHeightPercent"];
				var roleHeightVal = getActualValue(deviceWidth, roleHeightPercent);
				var $elementRole = $('<div class="speech-bubble common_elem commonDelete" id="'+ roleId +'" style="position: absolute; top:'+roleTopVal+'px; left:'+roleLeftVal+'px; display:'+divStyle+'; width:'+roleWidthVal+'px; height:'+roleHeightVal+'px;"><div class="role_text" >Role</div><div id= "clickableElement1" class="clickable_element_1" onclick = moveSpeechBubble("'+roleId+'","'+"LEFT"+'")><img src=""></div><div id= "clickableElement2" class="clickable_element_2" onclick = moveSpeechBubble("'+roleId+'","'+"DOWN"+'")><img src=""></div><textarea id = "'+textAreaId+'" maxlength="120" onkeypress="return disableKey(event)" onkeyup=fixTextAreaSize("'+divId+'",event)>'+roleValue+'</textarea><div id="' + deleteElementId +'" class="removeService delete_oval" onclick=deleteElement(this)><img src="../img/cross-black.png" alt="Delete" /></div></div>');				
				//var $elementRole = $('<div class="speech-bubble common_elem" id="'+ roleId +'" style="position: absolute; top:'+roleTopVal+'px; left:'+roleLeftVal+'px; display:'+divStyle+'; width:'+roleWidthVal+'px; height:'+roleHeightVal+'px;"><div class="role_text" >Role</div><div id= "clickableElement1" class="clickable_element_1" onclick = moveSpeechBubble("'+roleId+'","'+"LEFT"+'")><img src=""></div><div id= "clickableElement2" class="clickable_element_2" onclick = moveSpeechBubble("'+roleId+'","'+"DOWN"+'")><img src=""></div><textarea id = "'+textAreaId+'" maxlength="120">'+roleValue+'</textarea><div class="ui-resizable-handle ui-resizable-ne" id="negrip"></div><div class="ui-resizable-handle ui-resizable-nw" id="nwgrip"></div><div class="ui-resizable-handle ui-resizable-se" id="segrip"></div><div class="ui-resizable-handle ui-resizable-sw" id="swgrip"></div></div>');
				$("#outderDiv").append($elementRole);	
				setResizableCorner(img, roleId);
				/*Adjusting task area on load for devices
				 */
				var container = document.getElementById(textAreaId);
				if(navigator.userAgent.match(/Android/i) || navigator.userAgent.match(/iPhone/i) || navigator.userAgent.match(/iPad/i) || navigator.userAgent.match(/IPad/i)) {
					var insDisplay = sessionStorage.getItem("as2View");
					if(null == insDisplay || insDisplay == "" || insDisplay == "D"){
						reDragRoleFrame(divId);
						while($("#"+textAreaId).outerHeight() < container.scrollHeight + parseFloat(($("#"+textAreaId)).css("borderTopWidth")) + parseFloat(($("#"+textAreaId)).css("borderBottomWidth"))) {
							fixRoleFrameText(divId);
						}
					}
				}
				}
			}
			/*************************To fix the design ******************************/
			fixDesignAS();		
			if(sessionStorage.getItem("isCompleted") == 1 && sessionStorage.getItem("next_page") == "BS"){
				changeTaskDesc();
			}	
			setMappingAttributes();
			getOuterDivHeight();
			$(".after_sketch_task").css("z-index", 3);
			$(".commonDrop").css("z-index", 3);
			
			$(".loader_bg").fadeOut();
    	    $(".loader").hide();
	}	
}


/**************** to get the resized width value**********************/
$(document).ready(function() {                
    if (!sessionStorage.getItem("bsEdit") || (null == sessionStorage.getItem("bsEdit"))) {
          $('.after_sketch_task').resizable({
            stop:function(event,ui){
            var divId = $(this).attr('id');
            var finalWidth = $(this).width();                
            calculateEnergy(finalWidth, divId);                          
            /*************************To fix the design ******************************/
            fixDesignAS();
            }
        });           
    }
});

/**
 * Function add to allow only number as a input
 * in time energy field
 * @param event
 */
function exceptNumberOnly(event, obj) {
	var key = window.event ? event.keyCode : event.which;	
	if(event.keyCode == 13) {
		getEnergyValue(obj);
	} else if (event.keyCode == 8 || event.keyCode == 46
	 || event.keyCode == 37 || event.keyCode == 39 || event.keyCode == 9) {
	    return true;
	}
	else if ( key < 48 || key > 57 || key == 96) {
	    return false;
	} 
	else {
		return true;
	}
}


/**
 * Function add to get the input value
 * for time/energy field
 * @param obj
 */
function getEnergyValue(obj) {
	var re = /[^0-9]/g;
	var divCount = $('#outderDiv').children('.single_task_item_sketch:visible').length;
	var maxValue = 100 - ((divCount-1)*5);
	if ( re.test(obj.value) ){  
		alertify.alert("<img src='../img/alert-icon.png'><br /><p>Only number is accepted for task value.</p>");
		calculateTotalEnergy();
	}
	if(obj.value<5){
		alertify.alert("<img src='../img/alert-icon.png'><br /><p>The percent of time/energy for each task must be at least 5%.</p>");
		calculateTotalEnergy();
	} else if(obj.value > 90){
		//alertify.alert("Please add the task value less than 90.");
		alertify.alert("<img src='../img/alert-icon.png'><br /><p>The percent of time/energy for each task must be at most 90%.</p>");
		calculateTotalEnergy();
	} else if (obj.value > maxValue){
		alertify.alert("<img src='../img/alert-icon.png'><br /><p>Task cannot be " +obj.value+ ". Please re-enter</p>");
		calculateTotalEnergy();
	} else {
		var status = null;
		var energyIdFinal = obj.id;
		var divCnt = energyIdFinal.substring(9);
		var divIdFinal = "divImg_"+divCnt;
		var valueDiff = 0;
		var initialDivJson = JSON.parse(ClientSession.get("div_intial_value"));	
		divLength = initialDivJson.length;		
		for (var i = 0; i<divLength;i++) {
			var energyId = initialDivJson[i]["energyId"];
			var energyValue = initialDivJson[i]["energyValue"];
			if(energyId == obj.id){
				valueDiff = parseInt(obj.value - energyValue);
			}		
		}
		if(valueDiff>0){
			status = "+";
		} else {
			valueDiff = -(valueDiff);
			status = "-";
		}
		var unlockDivCount = $('#outderDiv').children('.single_task_item_sketch:visible').length;	//count the unlock div
		
		if(unlockDivCount == 1){
			//alertify.alert("Please unlock one task to resize");
			calculateTotalEnergy();
		} else if(status=='+'){
			valueDiff = getMaxEnergyValueToAdd(divIdFinal, valueDiff, "INPUT"); 
			if(valueDiff == 222) {
				calculateTotalEnergy();
			} else {
				calculateEnergyValue(valueDiff,status,divIdFinal, "RESIZE");
			}
		} else {
			calculateEnergyValue(valueDiff,status,divIdFinal, "RESIZE");
		}	
		
		//calculateEnergyValue(valueDiff,status,divIdFinal, "RESIZE");	
	}
}

/**
 * Function add to delete the value, passion, strength and role frame
 * @param obj
 */
function deleteElement(obj) {
	var errMsg = "";
	var divId = obj.id;
	alertify.set({ buttonReverse: true });
	if(divId.search("divStrengthClone_") != -1){
		errMsg = "this strength?";
	} else if(divId.search("divValueClone_") != -1){
		errMsg = "this value?";
	} else if(divId.search("divPassionClone_") != -1){
		errMsg = "this passion?";
	} else if(divId.search("cloneOval_") != -1){
		errMsg = "this role frame?";
	}
	alertify.set({ buttonReverse: true });
	alertify.set({ labels: {
	    ok     : "Yes",
	    cancel : "No"
	} });
	alertify.confirm("<img src='../img/confirm-icon.png'><br /><p>Are you sure you want to delete "+errMsg+"</p>", function (e) {
	    if (e) {	
	    	var length = divId.length;
	    	var deleteDivId = divId.substr(0,length-2);
		    var divDelete = document.getElementById(deleteDivId);		    	    	
		    if (divDelete) {
		    	if(divId.search("cloneOval_") != -1) {
		    		var speechBubbleId  = "role_"+deleteDivId;
		    		var speechBubbleDiv = document.getElementById(speechBubbleId);
		    		divDelete.setAttribute("style", "display:none");   
		    		speechBubbleDiv.setAttribute("style", "display:none");   
			    } else {
			    	divDelete.setAttribute("style", "display:none");   
			    }		    		
		    }		    
	    } 
	});		
}

/**
 * Function add to change the task description from before sketch
 * @param null
 */
function changeTaskDesc () {
	if(ClientSession.get("total_json_add_task")!=null){		
		var jsonObj = JSON.parse(ClientSession.get("total_json_add_task"));
		var countLoop = jsonObj.length;
		for (var i = 0; i<countLoop;i++){	
			var divIdBS = jsonObj[i]["divId"];
			var areaIdBS = jsonObj[i]["areaId"];
			var areaValueBS = jsonObj[i]["areaValue"];
			if(divIdBS.search('divImg_') != -1){
				areaValueBS = areaValueBS.replace(/`/g, ",");
				areaValueBS = areaValueBS.replace(/-_-/g, "#");
			}			
			var jsonObjAftrSktc = JSON.parse(ClientSession.get("total_json_after_sketch_final"));
			var jsonLength = jsonObjAftrSktc.length;
			for (var j = 0; j<jsonLength;j++){	
				var divIdAS = jsonObjAftrSktc[j]["divId"];
				var areaIdAS = jsonObjAftrSktc[j]["areaId"];
				var areaValueAS = jsonObjAftrSktc[j]["areaValue"];
				if(divIdAS.search('divImg_') != -1){
					areaValueAS = areaValueAS.replace(/`/g, ",");
					areaValueAS = areaValueAS.replace(/-_-/g, "#");
				}		
				if(divIdAS == divIdBS) {						
					var areaDesc = areaValueAS.indexOf("(");
					if(areaDesc!= -1){
						var areaDescAS = areaValueAS.substr(0,areaDesc).trim();
						var areaDescFinal = areaValueAS.substr(areaDesc, areaValueAS.length).trim();
						if(areaDescAS != areaValueBS){
							document.getElementById(areaIdAS).value = areaValueBS.concat(areaDescFinal);								
						} 							
					} else if(areaDesc == -1){
						if(areaValueAS != areaValueBS){
							document.getElementById(areaIdAS).value = areaValueBS;								
						}						
					}
				}
			}
			
		}
	}
}


function checkResizeCordinates(cloneId) {
	var cloneImage = document.getElementById(cloneId).getElementsByTagName("img")[0].getAttribute("src");
	var outerDivHeight = document.getElementById("outderDiv").style.height.replace('px','');		
	var outerDivWidth = $("#outderDiv").width();		
	var height = $(this).height();
	$('.oval_cloned_div').resizable({
		stop: function(event, ui) {			
			var cloneId = $(this).attr('id');
			var speechBubbleId="role_"+cloneId;
			var cloneLeft = $("#" + cloneId).position().left;
			var cloneRight = cloneLeft + $("#" +cloneId).width();		
			var cloneTop = $("#" + cloneId).position().top;
			var cloneBottom = cloneTop + $("#" +cloneId).height() + 25;			//	25 added for ROLE
			// checking if speech Bubble on top or bottom
			//if(($("#" + speechBubbleId).children(".ui-resizable-sw").css("display") == "block") || ($("#" + speechBubbleId).children(".ui-resizable-se").css("display") == "block")){
			if((cloneImage.search('left_bottom.png') != -1) || (cloneImage.search('right_bottom.png') != -1)){		
			// SpeechBubble on the bottom
				var speechBubbleId="role_"+cloneId;
				//var cloneBottomWithRole = cloneBottom + $("#" + speechBubbleId).position().top;
				var speechBubbleBottom = cloneBottom + $("#" +speechBubbleId).height();
				
				if(( cloneRight >  outerDivWidth ) && ( speechBubbleBottom > outerDivHeight )){		// Dragging diagonally
					
					// ******* adjustment of Right edge **************					
					var extraWidth = cloneRight - outerDivWidth;			//	calculating how much width is overflowing
					cloneRight = cloneRight - extraWidth - cloneLeft;		//	calculating new Width or Right			
					$("#" + cloneId).css("width", cloneRight+"px");			// New Right have to added, but left will remain same			
					
					// ****** adjustment of bottom edge *************					
					var speechBubbleHeight = $("#" +speechBubbleId).height();
					var extraBottom = speechBubbleBottom - outerDivHeight;				//	calculating how much height is overflowing
					cloneBottom = cloneBottom - speechBubbleHeight - extraBottom - cloneTop;		//	calculating new Height (here, cloneBottom = height of the clone)
					$("#" + cloneId).css("height", cloneBottom+"px");		// New Bottom have to added, but Top will remain same
				
				} else if ( cloneRight >  outerDivWidth ) {		// Dragged towards right edge
					var extraWidth = cloneRight - outerDivWidth;			//	calculating how much width is overflowing
					cloneRight = cloneRight - extraWidth - cloneLeft;		//	calculating new Width or Right			
					$("#" + cloneId).css("width", cloneRight+"px");			// New Right have to added, but left will remain same
					
				} else if( speechBubbleBottom > outerDivHeight ){		// Dragged towards Bottom edge
					var speechBubbleHeight = $("#" +speechBubbleId).height();
					var extraBottom = speechBubbleBottom - outerDivHeight;								//	calculating how much height is overflowing					
					cloneBottom = cloneBottom - extraBottom - speechBubbleHeight  - cloneTop;		//	calculating new Height (here, cloneBottom = height of the clone)
					$("#" + cloneId).css("height", cloneBottom+"px");		// New Bottom will be added, but Top will remain same					
				}				
			} else {	//	speechBubble on TOP, so need to adjust only oval's position	
				if(( cloneRight >  outerDivWidth ) && ( cloneBottom > outerDivHeight )){	// Dragging diagonally					
					// ******* adjustment of Right edge **************
					var extraWidth = cloneRight - outerDivWidth;			//	calculating how much width is overflowing
					cloneRight = cloneRight - extraWidth - cloneLeft;		//	calculating new Width or Right			
					$("#" + cloneId).css("width", cloneRight+"px");			// New Right have to added, but left will remain same
					
					// ****** adjustment of bottom edge *************
					var cloneTop = $("#" + cloneId).position().top;
					var extraBottom = cloneBottom - outerDivHeight;			//	calculating how much width is overflowing
					cloneBottom = cloneBottom - extraBottom - cloneTop;		//	calculating new Width or Right	
					$("#" + cloneId).css("height", cloneBottom+"px");		// New Bottom have to added, but Top will remain same				
					
				} else if ( cloneRight >  outerDivWidth ){	// Dragged towards right edge
					var extraWidth = cloneRight - outerDivWidth;			//	calculating how much width is overflowing
					cloneRight = cloneRight - extraWidth - cloneLeft;		//	calculating new Width or Right			
					$("#" + cloneId).css("width", cloneRight+"px");			// New Right have to added, but left will remain same
				} else if( cloneBottom > outerDivHeight ){		// Dragged towards Bottom edge
					var cloneTop = $("#" + cloneId).position().top;
					var extraBottom = cloneBottom - outerDivHeight;			//	calculating how much width is overflowing
					cloneBottom = cloneBottom - extraBottom - cloneTop;		//	calculating new Width or Right	
					$("#" + cloneId).css("height", cloneBottom+"px");		// New Bottom have to added, but Top will remain same					
				}		
			}
			fixRoleFramePosition(cloneId);		//	after adjusing the Oval, speech bubble is adjusing 
		}
	});
}

/**************** To resized the Div**********************/
$("#pageWrapCopy .ui-resizable-se").mouseleave(function() {
	$("#outderDiv").css({'height':($("#pageWrapCopy").height()+'px')});
});

/**************** To reduce the Z-index value to fix the UI**********************/
$(".single_task_item_sketch .ui-resizable-se").css({ 'z-index':'0'});

/**
 * Function add to remove the focus from the old tasks
 * @param null
 */
function removeFocus() {
	var divCountLoop = $('#outderDiv').children('.after_sketch_task').length;	
	for(var i=1; i<=divCountLoop;i++){
		var divIdImg = document.getElementById("divImg_" +i).id;
		$("#" + divIdImg).children().find(".add_task_area_sketch").removeClass("add_task_area_sketch_focus");
		$("#" + divIdImg).children().find(".add_task_area_sketch").removeClass("add_task_area_sketch_delete_focus");
		$("#" + divIdImg).children().find(".form-control-additional").removeClass("add_task_area_sketch_focus");
		$("#" + divIdImg).children().find(".form-control-additional").removeClass("add_task_area_sketch_delete_focus");
	}
}

/**
 * This function is executed for device only
 * Function add to disable the drag and resize functinality 
 * for task block
 * @param null
 */
function disableDraggableResizable(){
	if((navigator.userAgent.match(/iPhone|iPad|iPod/i))) {
		$(".after_sketch_task").children(".ui-resizable-se").css("display","none");
		$(".after_sketch_task").draggable( "option", "disabled", true );
		$(".after_sketch_task").resizable( "option", "disabled", true ); 		
	}
}


/**
 * This function is executed for device only
 * Function add to enable the drag and resize functinality 
 * for task block
 * @param object
 */
function enableDraggableResizable(divId){
	if((navigator.userAgent.match(/iPhone|iPad|iPod/i))) {
		var lockUnlockImg = document.getElementById(divId).getElementsByTagName("img")[2].getAttribute("src");
		 if (lockUnlockImg.search('imgUnlock.png') != -1) {	
			 $("#" +divId).children(".ui-resizable-se").css("display","block");
				$("#" +divId).draggable( "option", "disabled", false );
				$("#" +divId).resizable( "option", "disabled", false );
		 } else {
			 $("#" +divId).draggable( "option", "disabled", false );
		 }
	}
}

/**
 * This function is executed for device only
 * Function add to disable the drag and resize functinality 
 * for role frames
 * @param null
 */
function disableDraggableResizableOval(){
	$(".oval_cloned_div").draggable( "option", "disabled", true );
	$(".oval_cloned_div").resizable( "option", "disabled", true ); 
}

/**
 * This function is executed for device only
 * Function add to enable the drag and resize functinality 
 * for role frames
 * @param object
 */
function enableDraggableResizableOval(){
	//$(object).draggable( "option", "disabled", false );
	//$(object).resizable( "option", "disabled", false );
	//$(object).addEventListener('touchstart', enableDraggableResizableOval(this), false);
	$(".oval_cloned_div").draggable( "option", "disabled", false );
	$(".oval_cloned_div").resizable( "option", "disabled", false ); 
}


/**
 * This function is executed for device only
 * Function add to make the time/ energy field
 * and task description area editable
 * @param null
 */
$(document).ready(function() {		
	if (!sessionStorage.getItem("bsEdit")) {
    var windowSize = $(window).width();  
    if (windowSize <= 766) {
  	    $(".single_task_item_sketch").draggable().click(function(){
  		 trackerObj = $(this);
  		$(this).draggable( "option", "disabled", true ); 
  	    $(this).children(".energy_field").attr('contenteditable','true'); 
  	    $(this).children(".add_task_area_sketch").attr('contenteditable','true'); 
  	    $(this).children(".form-control-additional").attr('contenteditable','true'); });  	   
  	  
  	  $(".add_task_area_sketch, .energy_field, .form-control-additional").blur(function(){
  			$(trackerObj).draggable( 'option', 'disabled', false); 
  		$(trackerObj).attr('contenteditable','false'); });
  	  
    }
    else if (windowSize <= 1024) {   	
    	$(".after_sketch_task").on('touchstart', function(e){
    		enableDraggableResizable($(this).attr('id'));
    	/*if($(this).hasClass( "ui-state-disabled" )){    		
    		$(this).draggable( "option", "disabled", false );
    		$(this).resizable( "option", "disabled", false );
    		//enableDraggableResizable($(this).attr(id));
    	}else { 
    		$(this).draggable( "option", "disabled", true );
    		$(this).resizable( "option", "disabled", true );
    		$(this).children("textarea").attr('contenteditable','true');
    		$(this).children("textarea").focus();
    		
    	}*/
		});
		$(".oval_cloned_div").on('touchstart', function(e){
    		if($(this).hasClass( "ui-state-disabled" )){
    		$(this).draggable( "option", "disabled", false );    		
    		$(this).resizable( "option", "disabled", false );
    		$(this).children("removeService").toggle();
    	}else { 
    		$(this).draggable( "option", "disabled", true );    		
    		$(this).resizable( "option", "disabled", true );
    		$(this).children("textarea").attr('contenteditable','true');
    		$(this).children("textarea").focus();    		
    	}		
     });
    	disableDraggableResizable(); 
    	disableDraggableResizableOval();
    }
	}
});


/********** ISSUE 100**********/
/**
 * Function add to fix the left value of Task div
 * @param null
 */
function fixDesignToAddDelete() {
	 var taskWidth =  "";	
	 var mainDivWidth = "";
	 var positionLeft = "";
	 mainDivWidth = $('#outderDiv').width();
	var divCountLoop = $('#outderDiv').children('.after_sketch_task').length;	
	for(var i=1; i<=divCountLoop;i++){
		taskWidth = document.getElementById("divImg_"+i).style.width;
		var eachTaskWidth = parseInt(taskWidth.replace("px", ''));
		positionLeft = $("#divImg_"+i).position().left;
		var style = document.getElementById("divImg_" +i).getAttribute("style");
		if((style.search('block') != -1)){
			if ((eachTaskWidth + positionLeft + 45) > mainDivWidth || (eachTaskWidth + positionLeft + 45) < 170) {  // added for issue #264
				$("#divImg_"+i).css({ position: "absolute"});
				$("#divImg_"+i).css({ left: "50px"});
				$("#divImg_"+i).css({ top: "30px"});
			} 			
		}	
	}
}

/**
 * Function call when the task exteed the
 * containment while resizing
 * @param null
 */
$('.main_warp').mouseup(function() {
	disableDraggableResizable();
	callInitialJsonOnResize();
});

/**
 * Function add to risize the task 
 * on clicking on the outside of the div
 * @param null
 */
function callInitialJsonOnResize(){	
	var taskWidthPx =  "";
	var taskHeightPx = "";
	var mainDivWidthPx = "";	
	var divCountLoop = $('#outderDiv').children('.after_sketch_task').length;
	if( null == document.getElementById("pageWrapCopy").style.width || document.getElementById("pageWrapCopy").style.width == ""){
		if (navigator.userAgent.match(/iPhone|iPad|iPod/i)) { //tab
			mainDivWidthPx = "858px";
		} else {
			mainDivWidthPx = "1043px";
		}	
	 } else {
		mainDivWidthPx = document.getElementById("pageWrapCopy").style.width;
	 }	
	 mainDivWidth = parseInt(mainDivWidthPx.replace("px", ''));	
	for(var i=1; i<=divCountLoop;i++){
		taskHeightPx = document.getElementById("divImg_"+i).style.height;
		taskWidthPx = document.getElementById("divImg_"+i).style.width;
		positionLeft = $("#divImg_"+i).position().left;
		var taskWidth = parseInt(taskWidthPx.replace("px", ''));
		var style = document.getElementById("divImg_" +i).getAttribute("style");
		if((style.search('block') != -1)){
			if(navigator.userAgent.match(/Android/i)) {
				if(taskWidth > 225 || taskWidth < 70) {
					calculateTotalEnergy();
				}
			} else if(navigator.userAgent.match(/iPad/i)) {
				if(taskWidth > 335 || taskWidth < 112) {
					calculateTotalEnergy();
				}
			} else if(!(navigator.userAgent.match(/iPhone|iPad|iPod/i)) && !(navigator.userAgent.match(/Android/i)) ){		
				if(taskWidth > 395 || taskWidth < 140) {
					calculateTotalEnergy();
				} 
			}
		}			
	}
}

/**
 * Function add to get the actual position of each element
 * while navigating
 * @param null
 */
function getActualPosition(){
	var positionJson = [];
	var idx = 0;
    $("#outderDiv").children('.ui-draggable:visible').each(function() {
        var element = $( this );       
        var elementId = $(this).attr('id');
        var style = document.getElementById(elementId).getAttribute("style");        
        if(style.search('block') != -1){
        	var position = element.position();
        	var unitJson = {};
        	unitJson["Id"] = elementId;
        	unitJson["LeftPosition"] = Math.round(position.left);
			unitJson["TopPosition"] = Math.round(position.top);
			positionJson[idx++] = unitJson;			
        }
        sessionStorage.setItem("position_json",JSON.stringify(positionJson));
    });
}

/**
 * Function add fix the position of the Horizontal oval
 * while drop in the canvas
 * @param obj
 */
function fixPositionWhileDrop(obj, countOval){
	var divWidth = $("#outderDiv").width();
	var divHeight = $("#outderDiv").height();
	var speechBubbleHeight = 0;
	if($(".speech-bubble").css('min-height') == undefined){
		speechBubbleHeight = 90;
	} else {
		speechBubbleHeight = parseInt($(".speech-bubble").css('min-height'), 10);
	}
	
	$(".copied").css({width: (divWidth-50)+"px"});  
	if(!(navigator.userAgent.match(/iPhone|Android/i))){				
		$(".copied").css({height: (divHeight-speechBubbleHeight-40)+"px"});
	} else {
		$(".copied").css({height: (divHeight-speechBubbleHeight-250)+"px"});
	}
	
	$(".copied").css({left:"20px"});	
	$(".copied").css({top:(speechBubbleHeight+20)+"px"});
}
/**
 * Function removes the session items which will not be used
 */
function removePresentSessionItems() {
	sessionStorage.removeItem("total_json_obj");
	sessionStorage.removeItem("strength_count");
	sessionStorage.removeItem("passion_count");
	sessionStorage.removeItem("value_count");
	sessionStorage.removeItem("checked_passion");
	sessionStorage.removeItem("checked_strength");
	sessionStorage.removeItem("checked_value");
	sessionStorage.removeItem("strength_count");
	sessionStorage.removeItem("passion_count");
	sessionStorage.removeItem("value_count");
	sessionStorage.removeItem("total_json_add_task");
	sessionStorage.removeItem("total_json_after_sketch_final");
	sessionStorage.removeItem("checked_element");
	sessionStorage.removeItem("define_role");
	sessionStorage.removeItem("div_intial_value");
	sessionStorage.removeItem("position_json");
	sessionStorage.removeItem("passion");
	sessionStorage.removeItem("strength");
	sessionStorage.removeItem("value");
	sessionStorage.removeItem("element_outside_role");
	sessionStorage.removeItem("element_in_role");
}

/**
 * Function add to calculate
 * the greatest common devicer 
 * to calculate the aspect ratio
 * @param n1,
 * @param n2,
 * @returns num1
 */
function calc(n1,n2)
{
  var num1,num2;
  if(n1 < n2){ 
      num1=n1;
      num2=n2;  
   }
   else{
      num1=n2;
      num2=n1;
    }
  var remain=num2%num1;
  while(remain>0){
      num2=num1;
      num1=remain;
      remain=num2%num1;
  }
  return num1;
} 

/**
 * Function add to calculate
 * the percentage value of top position,
 * left position and width of each task
 * @param mainDivWidth
 * @param individualItem
 * @returns value
 */
function getPercentageValue(mainDivWidth, individualItem) {
	var value = (individualItem / mainDivWidth) * 100;
	return value;
}

/**
 * Function add to calculate
 * the height of the outer div with respect to width of that div
 * @param aspectRatio
 * @param deviceWidth
 * @returns deviceHeight
 */
function getActalHeight(aspectRatio, deviceWidth) {
	var deviceHeight = "";
	var length = aspectRatio.length;
	var n = aspectRatio.indexOf(":");
	var aspectRatioWidth = parseInt(aspectRatio.substring(0, n)); 
	var aspectRatioHeight = parseInt(aspectRatio.substring(n+1, length)); 
	deviceHeight = Math.round((deviceWidth * aspectRatioHeight) / aspectRatioWidth);
	return deviceHeight;
}

/**
 * Function add to calculate
 * the pixel value of top position, left position
 * and width of each task when the JSON loaded
 * @param deviceWidth
 * @param positionLeftPercent
 * @returns value
 */
function getActualValue(deviceWidth, positionLeftPercent) {
	var value = Math.round((deviceWidth * positionLeftPercent)/100);
	return value;
}

/**
 * Function add to calculate
 * the required value for other devices
 * with respect to desktop
 * @param diffValue
 * @param mainDivWidth
 * @returns diffValuePixel
 */
function getWidthDifference(diffValue, mainDivWidth) {
	var diffValuePixel = "";
	//var diffValuePercent = (diffValue / mainDivWidth) * 100;
	diffValuePixel = Math.round((mainDivWidth * diffValue) /1041);
	return diffValuePixel;
}

/***************** AFTER SKETCH EDIT *******************/
/**
 * Function takes back to before sketch edit page.
 */
function backToPreviousPageForEdit() {
	var navigation = Spine.Model.sub();
	navigation.configure("/user/navigation/goPrevious", "navigationSeq", "jobReferenceNos" , "profileId", "linkClicked", "rowId");
	navigation.extend( Spine.Model.Ajax );
	
	//Populate the model
	var replyModel = new navigation({
		navigationSeq: "1",
		jobReferenceNos: sessionStorage.getItem("jobReferenceNo"),
		profileId: sessionStorage.getItem("profileId"), 
		linkClicked: "FE",
		rowId: sessionStorage.getItem("rowIdentity")
	});
	replyModel.save(); //POST
	
	navigation.bind("ajaxError", function(record, xhr, settings, error) {
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
		alertify.alert("Unable to connect to the server.");
		return false;
	});
	navigation.bind("ajaxSuccess", function(record, xhr, settings, error) {
		var jsonStr = JSON.stringify(xhr);
		var obj = jQuery.parseJSON(jsonStr);
		if(sessionStorage.getItem("isLogout") == "N"){
			sessionStorage.setItem("total_json_add_task", obj.bsJson); //Job reference
			var myJsonObj = JSON.parse(obj.bsJson);
			sessionStorage.setItem("total_count", myJsonObj.length+1);
			sessionStorage.setItem("div_intial_value", obj.initialJson);
			sessionStorage.setItem("isCompleted", 1);
			sessionStorage.setItem("pageSequence", 1);
			sessionStorage.setItem("myPage", "BS");
			window.location = "beforeSketchEdit.jsp";
		}
	});
}


/**
 * function handles when cancel is pressed
 */
function cancelForEdit() {
	alertify.set({ buttonReverse: true });
	alertify.set({ labels: {
	    ok     : "Yes",
	    cancel : "No"
	} });
	alertify.confirm("<img src='../img/alert-icon.png'><br /><p>Are you sure you want to cancel? <br /> Nothing will be saved in search database.</p>", function (e) {
		if (e) {
			//if (sessionStorage.getItem("bsEdit")) {
			//	window.location = "resultPageFinal.jsp";
			//} else {
			//	window.location = sessionStorage.getItem("bsEdit");
			//}
			/*if (sessionStorage.getItem("bsEdit")) {
				window.location = "resultPageFinal.jsp";
			} else {
				window.location = "resultPage.jsp";
			}*/
			window.location = sessionStorage.getItem("bsEdit");
		}
	});
}

/**
 * Function add to fix the task box position
 * while adding and deleting any task
 */
function fixPosition(action) {
	var divCountLoop = $('#outderDiv').children('.after_sketch_task').length;	
	for (var i = 1; i<=divCountLoop; i++){	
		var divId = "divImg_" + i;		
		var style = document.getElementById("divImg_" + i).getAttribute("style");
		var outerDivHeight = $('#outderDiv').height();
		if((style.search('block') != -1)){
			$("#" + divId).css({ position : "relative"}); 
			$("#" + divId).css({ left : "20px"}); 
			if(parseInt((document.getElementById(divId).style.top).replace("px", '')) < 0) {
				$("#" + divId).css({ top : "10px"});
			} if ($("#" + divId).position().top > 300) {
				$("#" + divId).css({ top : "20px"});
			}
			var maxHeight = parseInt($("#" + divId).position().top) + parseInt($("#" + divId).height());
			
			if(action == "DELETE") {
				if (maxHeight > outerDivHeight){				
					maxHeight = maxHeight + 30;
					$("#pageWrapCopy").css({'height':(maxHeight+'px')});						
				}	
				 $("#outderDiv").css({'height':($("#pageWrapCopy").height()+'px')});
			} else {
				 //$("#outderDiv").css({'height':($("#pageWrapCopy").height()+'px')});
			}
					
		}				
	}
}

/**
 * Function add to fix the height of the outer div
 */
function getOuterDivHeight() {
	var maxHeight = 0;
	$("#outderDiv").children('.common_elem:visible').each(function() {		
        var element = $( this );       
        elementId = $(this).attr('id');	  	   
        var position = element.position();	  
        if(maxHeight < position.top + $(this).height() + 20 ) {
        	maxHeight = position.top + $(this).height() + 20;		       	
        }
    });
	if(maxHeight > 600 ) {
		/******** JCTP-4 *********/
		maxHeight = maxHeight + 25;		
		$("#pageWrapCopy").css({'height':(maxHeight+'px')});
		$("#outderDiv").css({'height':($("#pageWrapCopy").height()+'px')});
	}	
}


/**
 * Function checks if at least one Strength, Value and Passion
 * should present in after diagram
 * @returns {Boolean}
 */
function validateAttribute() {
	if ($('#outderDiv').children('.plus_sign_small:visible').length == 0) {
		return false;
	} else if ($('#outderDiv').children('.oval_small:visible').length == 0) {
		return false;
	} else if ($('#outderDiv').children('.up_triangle_small:visible').length == 0) {
		return false;
	} else {
		return true;
	}
}

/**
 * Function add to fix the height of the outer div
 * when task is added
 */
function fixDivLength() {	
	var maxHeight = "";
	var divCountLoop = $('#outderDiv').children('.after_sketch_task').length;	
	var outerDivHeight = $('#outderDiv').height();
	for (var i = 1; i<=divCountLoop; i++){	
		var divId = "divImg_" + i;	
		maxHeight = parseInt($("#" + divId).position().top) + parseInt($("#" + divId).height());
		maxHeight = maxHeight + 30;	
		if (outerDivHeight < maxHeight) {
			$("#pageWrapCopy").css({'height':(maxHeight+'px')});
			$("#outderDiv").css({'height':($("#pageWrapCopy").height()+'px')});
		}
	}
}

/**
 * Function add to delete the existing 
 * mapping attributes from after diagram
 * which are being deleted from mapping yourself page
 */
function setMappingAttributes() {
	/******  check existing Strength with the selected list ******/
	var existingStrValue = "";
	var strengthOldCount = $('#outderDiv').children('.plus_sign_small').length;	
	var strengthArray = new Array();	
	var strengthNewCount = $('#strength_pageWrap').children('.plus_sign_small').length;	
	for(var StrengthCnt = 1; StrengthCnt <= strengthNewCount; StrengthCnt++){	
		var style = document.getElementById("divStrength_" +StrengthCnt).getAttribute("style");
		if((style.search('block') != -1)){
			var strength = $("#divStrength_" + StrengthCnt + "  p").text();
			strengthArray.push(strength);	
		}	
	}	
	for (var strCount = 1; strCount <= strengthOldCount; strCount++){	
		var divStyle = document.getElementById("divStrengthClone_" +strCount).getAttribute("style");
		if (divStyle.search('block') != -1){
			existingStrValue = $("#divStrengthClone_" + strCount + "  p").text();				
			if (strengthArray.indexOf(existingStrValue) == -1){
				$('#divStrengthClone_' +strCount).css('display','none');
			}
		}
	}
	
	/******  check existing Value with the selected list ******/
	var existingValValue = "";
	var valueOldCount = $('#outderDiv').children('.oval_small').length;	
	var valueArray = new Array();	
	var valueNewCount = $('#value_pageWrap').children('.oval_small').length;	
	for(var ValueCnt = 1; ValueCnt <= valueNewCount; ValueCnt++){	
		var style = document.getElementById("divValue_" +ValueCnt).getAttribute("style");
		if((style.search('block') != -1)){
			var value = $("#divValue_" + ValueCnt + "  p").text();
			valueArray.push(value);	
		}	
	}	
	for (var valCount = 1; valCount <= valueOldCount; valCount++){	
		var divStyle = document.getElementById("divValueClone_" +valCount).getAttribute("style");
		if (divStyle.search('block') != -1){
			existingValValue = $("#divValueClone_" + valCount + "  p").text();				
			if (valueArray.indexOf(existingValValue) == -1){
				$('#divValueClone_' +valCount).css('display','none');
			}
		}
	}
	
	/******  check existing Passion with the selected list ******/
	var existingPasValue = "";
	var passionOldCount = $('#outderDiv').children('.up_triangle_small').length;	
	var passionArray = new Array();	
	var passionNewCount = $('#passion_pageWrap').children('.up_triangle_small').length;	
	for(var passionCnt = 1; passionCnt <= passionNewCount; passionCnt++){	
		var style = document.getElementById("divPassion_" +passionCnt).getAttribute("style");
		if((style.search('block') != -1)){
			var passion = $("#divPassion_" + passionCnt + "  p").text();
			passionArray.push(passion);	
		}	
	}	
	for (var pasCount = 1; pasCount <= passionOldCount; pasCount++){	
		var divStyle = document.getElementById("divPassionClone_" +pasCount).getAttribute("style");
		if (divStyle.search('block') != -1){
			existingPasValue = $("#divPassionClone_" + pasCount + "  p").text();				
			if (passionArray.indexOf(existingPasValue) == -1){
				$('#divPassionClone_' +pasCount).css('display','none');
			}
		}
	}

}

/**
 * This function is executed for device only
 * Function add to make the time/ energy field
 * and task description area editable
 * @param null
 */
function contentEditable() {
	if (windowSize <= 1024) {   
    	$(".after_sketch_task").on('touchstart', function(e){
    		if($(this).hasClass( "ui-state-disabled" )){
    		$(this).draggable( "option", "disabled", false );
    		$(this).resizable( "option", "disabled", false );
    	}else { 
    		$(this).draggable( "option", "disabled", true );
    		$(this).resizable( "option", "disabled", true );
    		$(this).children("textarea").attr('contenteditable','true');
    		$(this).children("textarea").focus();
    		
    	}
		});
		$(".oval_cloned_div").on('touchstart', function(e){
    		if($(this).hasClass( "ui-state-disabled" )){
    		$(this).draggable( "option", "disabled", false );    		
    		$(this).resizable( "option", "disabled", false );
    		$(this).children("removeService").toggle();
    	}else { 
    		$(this).draggable( "option", "disabled", true );    		
    		$(this).resizable( "option", "disabled", true );
    		$(this).children("textarea").attr('contenteditable','true');
    		$(this).children("textarea").focus();    		
    	}		
     });
    	disableDraggableResizable(); 
    	disableDraggableResizableOval();
    }
}

/**
 * This function is executed for device only
 * Function add to make the oval textarea field description editable
 */
function ovalcontentEditable() {
	$(".oval_cloned_div").on('touchstart', function(e){
		if($(this).hasClass( "ui-state-disabled" )){
		$(this).draggable( "option", "disabled", false );			            		
		$(this).resizable( "option", "disabled", false );
		//$(".speech-bubble").resizable( "option", "disabled", false );	
	}else { 
		$(this).draggable( "option", "disabled", true );			            		
		$(this).resizable( "option", "disabled", true );
		$(this).children("textarea").attr('contenteditable','true');
		$(this).children("textarea").text = "test";
		$(this).children("textarea").focus();
		//$(".speech-bubble").resizable( "option", "disabled", true );
		
	}
		$(".oval_cloned_div").on('touchstart', function(e){
    		if($(this).hasClass( "ui-state-disabled" )){
    		$(this).draggable( "option", "disabled", false );				            		
    		$(this).resizable( "option", "disabled", false );
    		//$(".speech-bubble").resizable( "option", "disabled", false );	
    	}else { 
    		$(this).draggable( "option", "disabled", true );				            		
    		$(this).resizable( "option", "disabled", true );
    		$(this).children("textarea").attr('contenteditable','true');
    		$(this).children("textarea").focus();
    		//$(".speech-bubble").resizable( "option", "disabled", true );
    	}
     });
 });
}

/******************* Null check for Relational Crafting ***********************/
function validateRelationalCrafting(){
	var divCountLoop = $('#outderDiv').children('.after_sketch_task').length;	
	
	/******************* Set focus to the Relational Crafting which does not contain value***********************/
	if(sessionStorage.getItem("isLogout")=="N"){
		removeFocus();
		for(var i=1; i<=divCountLoop;i++){
			var style = document.getElementById("divImg_" +i).getAttribute("style");
			if((style.search('block') != -1)){
				taskId = document.getElementById("divImg_"+i).id;
				taskDesc = document.getElementById("additionalId_"+i).value.trim();	
				if(taskDesc == "") {
					$("#" + taskId).children().find(".form-control-additional").addClass("add_task_area_sketch_focus");
				}
			}		
		}
	}
		
	if(sessionStorage.getItem("isLogout")=="N"){
		for(var i=1; i<=divCountLoop;i++){
			var style = document.getElementById("divImg_" +i).getAttribute("style");
			if((style.search('block') != -1)){
				taskDesc = document.getElementById("additionalId_"+i).value.trim();
				if(taskDesc == "") {
					return false;
				}
			}		
		}	
	}
	return true;
}


/**
 * Function added to fix the task position 
 * if it exceed the outer div area
 * while deleting or resizing any task
 * 
 */
function adjustTaskWhileResize() {	
	var divCountLoop = $('#outderDiv').children('.after_sketch_task').length;	
	var mainDivWidth = "";
	for (var j = 1; j<=divCountLoop; j++){	
		var taskId = "divImg_" + j;	 
			var pos = $("#" + taskId).position();				
			var addition = pos.left + $('#'+ taskId).width();				
			if(navigator.userAgent.match(/Android|iPhone/i)) {
				mainDivWidth = $("#outderDiv").width() - 20;
			} else {
				mainDivWidth = $("#outderDiv").width() - 35;
			}
			if(addition > mainDivWidth) {
				var newLeft = addition - mainDivWidth;
				var newLeftPosition = pos.left - newLeft;
				$("#" + taskId).css({ left : newLeftPosition+'px'}); 
			}		
	}	
}

/**
 * Function add to fix the task position which is added
 * @param divWidth
 * @param divId
 */
function addTaskPosition(divWidth, divId) {
	 var topPosition = 0;
	 var maxTopDivId = "";
	 var maxLeftPosition = 0;
	 var maxFinalLeft = 0;
	 var maxFinalBottom = 0;
	 var maxTopRoleId = "";
	 var topRolePosition = 0;	 
	 /**
	  * calculate the highest final top among all the tasks
	  * store the highest final height
	  * store the div id of highest final top element
	  **/
	  $("#outderDiv").children('.after_sketch_task:visible').each(function() {		
	        var element = $( this );       
	        elementId = $(this).attr('id');	  	   
	        var position = element.position();	 	        	
	        if(topPosition <= position.top ) {
	        	topPosition = position.top;	
	        	maxTopDivId = elementId;	  
				maxLeftPosition = position.left + $(this).width();				
	        }	        		       
	    });
	 /**
	  * calculate the highest final top among all the role frames
	  * store the highest final height
	  * store the div id of highest final top role frame
	  **/
	  $("#outderDiv").children('.oval_cloned_div:visible').each(function() {
	        var element = $( this );       
	        elementId = $(this).attr('id');	  	       
	        var position = element.position();	 	        	
	        if(topRolePosition <= position.top ) {
	        	topRolePosition = position.top;	
	        	maxTopRoleId = elementId;	        		
	        }	        		       
	    });
	 
	 /**
	  * task added when role frame is not present in the canvas
	  */
	 if(($('#outderDiv').children('.oval_cloned_div:visible').length == 0)) {
		var maxFinalTop =  $("#" + maxTopDivId).position().top;
		maxFinalBottom =  $("#" + maxTopDivId).position().top + $("#" + maxTopDivId).height();
		var parentBottom = parseInt($("#outderDiv").position().top) + parseInt($("#outderDiv").height());		
		maxFinalLeft = maxLeftPosition;
		
		/*** check if any other task is leaning on the left and bottom of the selected div ***/
		$("#outderDiv").children('.after_sketch_task:visible').each(function() {
			 var element = $( this );       	
		     var position = element.position();	
		     var eachTaskBottom = position.top + $(this).height();
			 var eachTaskLeft = position.left + $(this).width();						
			 if((eachTaskBottom > maxFinalTop && eachTaskBottom < parentBottom) ) {		
				 if((maxLeftPosition < eachTaskLeft)){
					 if(maxFinalLeft < eachTaskLeft){
							maxFinalLeft = eachTaskLeft;
						} 						
				 } else {					
					if(eachTaskBottom > maxFinalBottom) {						
						maxFinalBottom = eachTaskBottom;
					}	
				 }
			 }			 
		});
		calculateFreeArea(maxFinalTop,maxFinalLeft,maxFinalBottom,divWidth, divId);
	 }
	 /**
	  * task added when role frame is present in the canvas
	  */
	 else {
		 var maximumTaskTop = $("#"+maxTopDivId).position().top + $("#"+maxTopDivId).height();
		 var maximumRoleTop = $("#"+maxTopRoleId).position().top + $("#"+maxTopRoleId).height();
		 /*** task added in the bottom of the role frame ***/
		 if (maximumRoleTop > maximumTaskTop) {
			 $("#" + divId).css({ position : "absolute"}); 
			 $("#" + divId).css({ top : (maximumRoleTop + 20)+'px'});		 
			 $("#" + divId).css({ left : "0px"}); 	
			 increaseOuterDivWhileAdd(divId);
		 } else {			 
			 var maxFinalTop =  $("#" + maxTopDivId).position().top;
			 maxFinalBottom =  $("#" + maxTopDivId).position().top + $("#" + maxTopDivId).height();
			 var parentBottom = parseInt($("#outderDiv").position().top) + parseInt($("#outderDiv").height());		
			 maxFinalLeft = maxLeftPosition;
				$("#outderDiv").children('.after_sketch_task:visible').each(function() {
					 var element = $( this );       
				     var position = element.position();	
				     var eachTaskBottom = position.top + $(this).height();
					 var eachTaskLeft = position.left + $(this).width();			
					 if((eachTaskBottom > maxFinalTop && eachTaskBottom < parentBottom) ) {		
						 if((maxLeftPosition < eachTaskLeft)){
							 if(maxFinalLeft < eachTaskLeft){
									maxFinalLeft = eachTaskLeft;
								} 						
						 } else {					
							if(eachTaskBottom > maxFinalBottom) {						
								maxFinalBottom = eachTaskBottom;
							}	
						 }
					 }
				});
			calculateFreeArea(maxFinalTop,maxFinalLeft,maxFinalBottom,divWidth, divId);
		 }
	 }
}

/**
 * Function add to calculate the free area 
 * for the added task
 * @param maxFinalTop
 * @param maxFinalLeft
 * @param maxFinalBottom
 * @param divWidth
 * @param divId
 */
function calculateFreeArea(maxFinalTop,maxFinalLeft,maxFinalBottom,divWidth, divId) {
	 var freeArea = $("#outderDiv").width() - maxFinalLeft - 20;
	 /*** Task added in the right side of the existing task ***/
	 if(divWidth  < freeArea) {	
		$("#" + divId).css({ position : "absolute"}); 
		$("#" + divId).css({ top : maxFinalTop+'px'});					
		$("#" + divId).css({ left : (maxFinalLeft+20) +'px'}); 
		adjustTaskWhileResize();
		increaseOuterDivWhileAdd(divId);
	 }
	 /*** Task added at the bottom of the existing task ***/
	 else { 
		$("#" + divId).css({ position : "absolute"}); 
		$("#" + divId).css({ top : (maxFinalBottom + 20)+'px'});					
		$("#" + divId).css({ left : '0px'}); 
		increaseOuterDivWhileAdd(divId);
	 }
}

/**
 * Function add to increase the outer div area
 * while any task is added at the bottom
 * @param divId
 */
function increaseOuterDivWhileAdd(divId) {
	var outerDivHeight = $('#outderDiv').height();
	var maxHeight = parseInt($("#" + divId).position().top) + parseInt($("#" + divId).height());
	 if (maxHeight > outerDivHeight){				
		maxHeight = maxHeight + 30;
		$("#pageWrapCopy").css({'height':(maxHeight+'px')});						
	 }	
	 $("#outderDiv").css({'height':($("#pageWrapCopy").height()+'px')});
}


/**
 * Function add to set the z-index property
 * task boxes get focus while click outside the role frame or on the role frame
 * role frame get focus while drag or resize the role frame
 */
function toogleFocus(){
	if (!sessionStorage.getItem("bsEdit")) {
		$(".after_sketch_task").draggable({start:updateTaskContent});
		}
		$(".after_sketch_task").click(updateTaskContent);
		$(".after_sketch_task").focus(updateTaskContent);
		
		function updateTaskContent() {	
			$(".speech-bubble").css("z-index", 3);
			$(".after_sketch_task").css("z-index", 3);
			$(this).css("z-index",4);
		}				
				
		$(document).on("click", ".pageWrap_after_sketch", function(event){
			//$(".oval_cloned_div").css("z-index", 4);
			$(".after_sketch_task").css("z-index", 3);
			$(".speech-bubble").css("z-index", 4);
			$(".commonDrop").css("z-index", 3);
			
			$('oval_cloned_div').draggable({	
            	start: function(){
            		$(".oval_cloned_div").css("z-index", 4);
                },
                stop: function(){
                	$(".oval_cloned_div").css("z-index", 2);
                },               
            });
			 
	    });

		if(($('#outderDiv').children('.oval_cloned_div:visible').length > 0)) {
	    $(document).on("click", ".oval_cloned_div", function(event){
	        event.stopPropagation();
	       // $(".oval_cloned_div").css("z-index", 3);
			$(".after_sketch_task").css("z-index", 4);
			$(".speech-bubble").css("z-index", 3);
			$(".commonDrop").css("z-index", 140);	
			$('oval_cloned_div').draggable({	
            	start: function(){
            		$(".oval_cloned_div").css("z-index", 3);
                },
                stop: function(){
                	$(".oval_cloned_div").css("z-index", 2);
                },               
            });
	    });
		}
	    
		if(($('#outderDiv').children('.oval_cloned_div:visible').length > 0)) {
	    $(document).on("click", ".after_sketch_task", function(event){
	        event.stopPropagation();
	        //$(".oval_cloned_div").css("z-index", 3);
			$(".after_sketch_task").css("z-index", 4);
			$(".speech-bubble").css("z-index", 3);
			$(".commonDrop").css("z-index", 140);	
			$('oval_cloned_div').draggable({	
            	start: function(){
            		$(".oval_cloned_div").css("z-index", 3);
                },
                stop: function(){
                	$(".oval_cloned_div").css("z-index", 2);
                },               
            });
	    });
		}
		
		if(($('#outderDiv').children('.oval_cloned_div:visible').length > 0)) {
		    $(document).on("click", ".commonDrop", function(event){
		        event.stopPropagation();
		       // $(".oval_cloned_div").css("z-index", 3);
				$(".after_sketch_task").css("z-index", 4);
				$(".speech-bubble").css("z-index", 3);
				$(".commonDrop").css("z-index", 140);
				$('oval_cloned_div').draggable({	
	            	start: function(){
	            		$(".oval_cloned_div").css("z-index", 3);
	                },
	                stop: function(){
	                	$(".oval_cloned_div").css("z-index", 2);
	                },               
	            });
		    });
		}
}
		
/**
 * Function call to set the height of the div
 * while the tasks are being deleted
 * 
 */
function setOuterDivWhileDelete() {
	var maxHeight = 0;
	var outerDivHeight = $("#outderDiv").height();
	$("#outderDiv").children('.common_elem:visible').each(function() {		
        var element = $( this );       
        elementId = $(this).attr('id');	  	   
        var position = element.position();	       
        if(maxHeight < position.top + $(this).height() ) {
        	 maxHeight = position.top + $(this).height();				
        }	        		       
    });
	if(outerDivHeight < maxHeight+25) {
		maxHeight = maxHeight + 25;
		$("#pageWrapCopy").css({'height':(maxHeight+'px')});
		$("#outderDiv").css({'height':($("#pageWrapCopy").height()+'px')});
	}	
}

function silentSave() {
	var t = window.location+"";
	if (t.search("afterSketchEdit.jsp") == -1) {
		sessionStorage.setItem("silentSave", "s");
		goToNextPageAftSktch();
	}
}

function storeAfterSketchDiagramSilent(json) {
//	getOuterDivHeight();
	var timeSpent = getTotalTimeSpent();
	var outerDivHeight = document.getElementById('pageWrapCopy').style.height;
	var outerDivWidth = document.getElementById('pageWrapCopy').style.width;	
	var element = document.getElementById("outderDiv");
	html2canvas(element, {
		onrendered: function(canvas) {
            var snapShotURL = canvas.toDataURL("image/png"); //get's image string
            var isEdit = "N";
            if (sessionStorage.getItem("bsEdit")) {
            	isEdit = "Y";
            }
            
            var afterSketchModel = Spine.Model.sub();
            afterSketchModel.configure("/user/saveAfterSketchDiagram", "createdBy", "jsonString", 
            		"jobReferenceString", "snapShot", "jobTitle", "roleJson", "divHeight", "divWidth", 
            		"totalCount", "divInitialVal", "timeSpents", "isCompleteds" , "profileId", 
            		"positionJson", "isEdits", "linkClicked", "rowId", "initialSave", "isPrevious", "jctUserId", "elementOutsideRoleJson");
            afterSketchModel.extend( Spine.Model.Ajax );
            var isCompleted = "N";
            if(sessionStorage.getItem("isLogout") == "Y"){
            	isCompleted = "Y";
            }
            //Populate the model with data to transfer
        	var modelPopulator = new afterSketchModel({  
        		createdBy: sessionStorage.getItem("jctEmail"), 
        		jsonString: json,
        		jobReferenceString: sessionStorage.getItem("jobReferenceNo"),
        		snapShot: snapShotURL,
        		jobTitle: sessionStorage.getItem("jobTitle"),
        		roleJson: jQuery.parseJSON(sessionStorage.getItem("define_role").replace(/\?/g, ";_;")),
        		divHeight: outerDivHeight,
        		divWidth: outerDivWidth,
        		totalCount: sessionStorage.getItem("total_count"),
        		divInitialVal: JSON.parse(sessionStorage.getItem("div_intial_value")),
        		timeSpents : "0",
        		isCompleteds: isCompleted,
        		profileId: sessionStorage.getItem("profileId"),
        		positionJson: sessionStorage.getItem("position_json"),
        		isEdits: isEdit,
        		linkClicked: sessionStorage.getItem("linkClicked"),
        		rowId: sessionStorage.getItem("rowIdentity"),
        		initialSave: "Y",
        		isPrevious: "S",
        		jctUserId: sessionStorage.getItem("jctUserId"),
        		elementOutsideRoleJson: jQuery.parseJSON(sessionStorage.getItem("element_outside_role").replace(/\?/g, ";_;"))
        	});
        	modelPopulator.save(); //POST
        	
        	afterSketchModel.bind("ajaxError", function(record, xhr, settings, error) {
        		$(".loader_bg").fadeOut();
        	    $(".loader").hide();
        	});
        	
        	afterSketchModel.bind("ajaxSuccess", function(record, xhr, settings, error) {
        		sessionStorage.removeItem("silentSave");
        	});
            }
        });
}

/**************************ADDED FOR LOCK UNLOCK FUNCIONALITY**********************************/

/**
 * Function call to lock and unlock the task
 * @param divId
 * @param divIdLock
 */
function lockUnlockDiv(divId, divIdLock) {
	var lockUnlockImg = document.getElementById(divIdLock).getElementsByTagName("img")[0].getAttribute("src");
	 if (lockUnlockImg.search('imgUnlock.png') != -1) {	
		 lockDiv(divId, divIdLock);
		 generateJsonObj();
	 } else {		
		 unlockDiv(divId, divIdLock);
		 generateJsonObj();
	 }
}

/**
 * Function call to set the class for the task boxes
 * based on lock and unlock functionality
 * @param divId
 * @param divLockId
 * @param lockImage
 */
function setClassAgainstLockUnlockElememt(divId, divLockId, lockImage) {
	 if (lockImage.search('imgUnlock.png') != -1) {	
		if((navigator.userAgent.match(/iPhone|iPad|iPod/i))) {	
			 unlockDivIpad(divId, divLockId);
		 } else {
			 unlockDiv(divId, divLockId);	 
		 }
	 } else {
		 lockDiv(divId, divLockId);
	 }	
}

/**
 * Function call while click on unlock icon
 * to lock the task
 * @param divId
 * @param divIdLock
 */
function unlockDiv(divId, divIdLock) {
	document.getElementById(divIdLock).getElementsByTagName("img")[0].setAttribute("src", "../img/imgUnlock.png");
	$("#" + divId).addClass("single_task_item_sketch");
	$("#" + divId).removeClass("lock_single_task_item_sketch");
	//if((navigator.userAgent.match(/iPhone|iPad|iPod/i)) || (navigator.userAgent.match(/Android/i)) ) {
	if((navigator.userAgent.match(/iPhone|iPad|iPod/i))) {
		$("#" + divId).children(".ui-resizable-se").css("display","block");
		enableDraggableResizable(divId);	
	} else {
		$("#" + divId).children(".ui-resizable-se").css("display","block");
	}
	var divCnt = divId.substring(7);
	$("#energyId_" + divCnt).removeClass("energy_lock");
	document.getElementById("energyId_" +divCnt).disabled = false;
	document.getElementById("divImg_"+divCnt+"_x").style.display = "block";
}

/**
 * This function call for Ipad and Iphone only
 * Function call while click on unlock icon
 * to lock the task
 * @param divId
 * @param divIdLock
 */
function unlockDivIpad(divId, divIdLock) {	
	document.getElementById(divIdLock).getElementsByTagName("img")[0].setAttribute("src", "../img/imgUnlock.png");
	$("#" + divId).addClass("single_task_item_sketch");
	$("#" + divId).removeClass("lock_single_task_item_sketch");
	var divCnt = divId.substring(7);
	$("#energyId_" + divCnt).removeClass("energy_lock");
	document.getElementById("energyId_" +divCnt).disabled = false;
	document.getElementById("divImg_"+divCnt+"_x").style.display = "block";
}

/**
 * Function call while click on lock icon
 * to unlock the task
 * @param divId
 * @param divIdLock
 */
function lockDiv(divId, divIdLock) {
	document.getElementById(divIdLock).getElementsByTagName("img")[0].setAttribute("src", "../img/imgLock.png");
	$("#" + divId).removeClass("single_task_item_sketch");
	$("#" + divId).addClass("lock_single_task_item_sketch");
	$("#" + divId).children(".ui-resizable-se").css("display","none");
	var divCnt = divId.substring(7);
	$("#energyId_" + divCnt).addClass("energy_lock");
	document.getElementById("energyId_" +divCnt).disabled = true;
	document.getElementById("divImg_"+divCnt+"_x").style.display = "none";
}


/**
 * This Function calculate the maximum energy value
 * that can be added
 * @param divIdFinal
 * @param differenceWidth
 * @param action
 * @returns differenceWidth
 */
function getMaxEnergyValueToAdd(divIdFinal, differenceWidth, action) {
	var lockDivEnergyVal = 0;
	var unlockDivEnergyVal = 0;
	var initialDiv = JSON.parse(ClientSession.get("div_intial_value"));	
	divLength = initialDiv.length;
	var divCnt = divIdFinal.substring(7);
	var energyValue = initialDiv[divCnt-1]["energyValue"];
	var unlockDivCount = $('#outderDiv').children('.single_task_item_sketch:visible').length;	//count the unlock div
	for (var i = 0; i<divLength;i++) {
		var lockImage = initialDiv[i]["lockImage"];	
		if (lockImage.search('imgUnlock.png') != -1) {	
			unlockDivEnergyVal = parseInt(unlockDivEnergyVal) + parseInt(initialDiv[i]["energyValue"]); //total energy value of unlock div // 66				
		} else {
			 lockDivEnergyVal = parseInt(lockDivEnergyVal) + parseInt(initialDiv[i]["energyValue"]); //total energy value of lock div //33
		}
	}	 
	var restTaskIntoFive = parseInt((unlockDivCount - 1) * 5);
	var maximumAddedValue = parseInt((100 - lockDivEnergyVal)) - restTaskIntoFive;
	if(parseInt(energyValue) + parseInt(differenceWidth) > maximumAddedValue) {
		if(action == "RESIZE"){
			var extraValue = (parseInt(energyValue) + parseInt(differenceWidth)) - parseInt(maximumAddedValue);
			differenceWidth = parseInt(differenceWidth) - parseInt(extraValue);
		} else {
			differenceWidth = 222;
		}	
	}
	return differenceWidth;
}

/**
 * This Function calculate the energy value
 * that will be added
 */
function checkEnergyValWhileAdd() {
	var unlockDivEnergyVal = 0;
	var initialDiv = JSON.parse(ClientSession.get("div_intial_value"));	
	divLength = initialDiv.length;
	for (var i = 0; i<divLength;i++) {
		var lockImage = initialDiv[i]["lockImage"];	
		if (lockImage.search('imgUnlock.png') != -1) {	
			var energyValue = initialDiv[i]["energyValue"];
			if(energyValue > 5){
				var valueLessFive = parseInt(initialDiv[i]["energyValue"]) - 5;
				unlockDivEnergyVal = parseInt(unlockDivEnergyVal) + parseInt(valueLessFive); //total energy value of unlock div // 66	
			}			
		} 
	}	
	return unlockDivEnergyVal;
}

/**
 * This Function save the data in session 
 * while the page refresh
 */
function saveAfterSketchWhileRefresh() {	
	var cntTask = 1;
	var totalJsonAftSktch = [];
	var idx = 0;
	var divValueTask = document.getElementById("divImg_" + cntTask);
	
	var mainDivWidth = $("#outderDiv").width();
	var mainDivHeight = $("#outderDiv").height();	
    
	var gcd=calc(mainDivWidth,mainDivHeight);
  	var r1=mainDivWidth/gcd;
    var r2=mainDivHeight/gcd;
    var ratio=r1+":"+r2;
    
	while(divValueTask != null) {
		var unitJson = {};					
		unitJson["divId"] = document.getElementById("divImg_" + cntTask).id;
		unitJson["divIdx"] = document.getElementById("divImg_" + cntTask +"_x").id;
		unitJson["energyId"] = document.getElementById("energyId_"+ cntTask).id;
		unitJson["energyValue"] = document.getElementById("energyId_"+ cntTask).value;
		unitJson["style"] = document.getElementById("divImg_" + cntTask).getAttribute("style");
		unitJson["areaId"] = document.getElementById("areaId_"+ cntTask).id;
		unitJson["areaValue"] = document.getElementById("areaId_"+ cntTask).value.replace(/,/g, "`").replace(/#/g, "-_-").replace(/\?/g, ";_;");
		unitJson["additionalId"] = document.getElementById("additionalId_"+ cntTask).id;
		unitJson["additionalValue"] = document.getElementById("additionalId_"+ cntTask).value.replace(/,/g, "`").replace(/#/g, "-_-").replace(/\?/g, ";_;");	
		unitJson["positionLeft"] = document.getElementById("divImg_" + cntTask).style.left;
		unitJson["positionTop"] = document.getElementById("divImg_" + cntTask).style.top;
		unitJson["positionWidth"] = document.getElementById("divImg_" + cntTask).style.width;
		unitJson["positionHeight"] = document.getElementById("divImg_" + cntTask).style.height;
		
		var positionLeftPixel = document.getElementById("divImg_" + cntTask).style.left;	
		var positionLeft = parseInt(positionLeftPixel.replace("px", ''));
		unitJson["positionLeftPixel"] = positionLeftPixel;
		unitJson["positionLeftPercent"] = getPercentageValue(mainDivWidth, positionLeft);
		var positionTopPixel = document.getElementById("divImg_" + cntTask).style.top;	
		var positionTop = parseInt(positionTopPixel.replace("px", ''));
		unitJson["positionTopPixel"] = positionTopPixel;
		unitJson["positionTopPercent"] = getPercentageValue(mainDivHeight, positionTop);
		var positionWidthPixel = document.getElementById("divImg_" + cntTask).style.width;	
		var positionWidth = parseInt(positionWidthPixel.replace("px", ''));
		unitJson["positionWidthPercent"] = getPercentageValue(mainDivWidth, positionWidth);
		unitJson["mainDivWidth"] = mainDivWidth;
		unitJson["mainDivHeight"] = mainDivHeight;
		unitJson["aspectRatio"] = ratio;
		
		/***** Added for lock and unlock feature *****/
		unitJson["divLockId"] = document.getElementById("divImg_" + cntTask +"_lock").id;
		unitJson["lockImage"] = document.getElementById("divImg_" + cntTask+ "_lock").getElementsByTagName("img")[0].getAttribute("src");
		
		totalJsonAftSktch[idx++] = unitJson;
		divValueTask = document.getElementById("divImg_" + ++cntTask);
	}
	
	
	var cntStrength = 1;
	var divStrength = document.getElementById("divStrengthClone_" + cntStrength);
	while(divStrength != null) {
			var unitJson = {};						
			unitJson["divId"] = document.getElementById("divStrengthClone_" + cntStrength).id;	
			var divStrengthId = unitJson["divId"];		
			var divStrengthVal = $("#" + divStrengthId + "  p").text();
			unitJson["divValue"] = divStrengthVal;		
			unitJson["style"] = document.getElementById("divStrengthClone_" + cntStrength).getAttribute("style");
			unitJson["positionLeft"] = document.getElementById("divStrengthClone_" + cntStrength).style.left;
			unitJson["positionTop"] = document.getElementById("divStrengthClone_" + cntStrength).style.top;	
			
			var positionLeftPixel = document.getElementById("divStrengthClone_" + cntStrength).style.left;	
			var positionLeft = parseInt(positionLeftPixel.replace("px", ''));
			unitJson["positionLeftPercent"] = getPercentageValue(mainDivWidth, positionLeft);
			var positionTopPixel = document.getElementById("divStrengthClone_" + cntStrength).style.top;	
			var positionTop = parseInt(positionTopPixel.replace("px", ''));
			unitJson["positionTopPercent"] = getPercentageValue(mainDivHeight, positionTop);
			unitJson["aspectRatio"] = ratio;
			totalJsonAftSktch[idx++] = unitJson;		
			divStrength = document.getElementById("divStrengthClone_" + ++cntStrength);		
	}	
	
	var cntValue = 1;
	var divValue = document.getElementById("divValueClone_" + cntValue);
	while(divValue != null) {
		var unitJson = {};		
		unitJson["divId"] = document.getElementById("divValueClone_" + cntValue).id;	
		var divValueId = unitJson["divId"];		
		var divValueVal = $("#" + divValueId + "  p").text();
		unitJson["divValue"] = divValueVal;	
		unitJson["style"] = document.getElementById("divValueClone_" + cntValue).getAttribute("style");		
		unitJson["positionLeft"] = document.getElementById("divValueClone_" + cntValue).style.left;
		unitJson["positionTop"] = document.getElementById("divValueClone_" + cntValue).style.top;	
		
		var positionLeftPixel = document.getElementById("divValueClone_" + cntValue).style.left;	
		var positionLeft = parseInt(positionLeftPixel.replace("px", ''));
		unitJson["positionLeftPercent"] = getPercentageValue(mainDivWidth, positionLeft);
		var positionTopPixel = document.getElementById("divValueClone_" + cntValue).style.top;	
		var positionTop = parseInt(positionTopPixel.replace("px", ''));
		unitJson["positionTopPercent"] = getPercentageValue(mainDivHeight, positionTop);
		unitJson["aspectRatio"] = ratio;
		totalJsonAftSktch[idx++] = unitJson;
		divValue = document.getElementById("divValueClone_" + ++cntValue);
	}

	var cntPassion = 1;
	var divPassion = document.getElementById("divPassionClone_" + cntPassion);
	while(divPassion != null) {
		var unitJson = {};		
		
		unitJson["divId"] = document.getElementById("divPassionClone_" + cntPassion).id;
		var divPassionId = unitJson["divId"];		
		var divPassionVal = $("#" + divPassionId + "  p").text();
		unitJson["divValue"] = divPassionVal;	
		unitJson["style"] = document.getElementById("divPassionClone_" + cntPassion).getAttribute("style");			
		unitJson["positionLeft"] = document.getElementById("divPassionClone_" + cntPassion).style.left;
		unitJson["positionTop"] = document.getElementById("divPassionClone_" + cntPassion).style.top;		
		
		var positionLeftPixel = document.getElementById("divPassionClone_" + cntPassion).style.left;	
		var positionLeft = parseInt(positionLeftPixel.replace("px", ''));
		unitJson["positionLeftPercent"] = getPercentageValue(mainDivWidth, positionLeft);
		var positionTopPixel = document.getElementById("divPassionClone_" + cntPassion).style.top;	
		var positionTop = parseInt(positionTopPixel.replace("px", ''));
		unitJson["positionTopPercent"] = getPercentageValue(mainDivHeight, positionTop);
		unitJson["aspectRatio"] = ratio;
		totalJsonAftSktch[idx++] = unitJson;
		divPassion = document.getElementById("divPassionClone_" + ++cntPassion);
	}
	
	var cntClone = 1;
	var divClone = document.getElementById("cloneOval_" + cntClone);
	while(divClone != null) {
		var unitJson = {};	
		
		unitJson["divId"] = document.getElementById("cloneOval_" + cntClone).id;
		var cloneId = unitJson["divId"];		
		//var cloneValue = $("#" + cloneId + " .speech-bubble textarea").val().replace(/#/g, "-_-").replace(/\?/g, ";_;");
		var cloneValue = $("#role_" + cloneId + " textarea").val().replace(/#/g, "-_-").replace(/\?/g, ";_;");
		unitJson["roleValue"] = cloneValue;			
		unitJson["positionLeft"] = document.getElementById("cloneOval_" + cntClone).style.left;
		unitJson["positionTop"] = document.getElementById("cloneOval_" + cntClone).style.top;	
		unitJson["positionWidth"] = document.getElementById("cloneOval_" + cntClone).style.width;
		unitJson["positionHeight"] = document.getElementById("cloneOval_" + cntClone).style.height;
		var position = $( "#cloneOval_" + cntClone ).position();
		unitJson["positionLeft"] = position.left;
		unitJson["positionTop"] = position.top;	
		unitJson["style"] = document.getElementById("cloneOval_" + cntClone).getAttribute("style");
		//unitJson["type"] = document.getElementById("cloneOval_" + cntClone).getAttribute("title");
		
		var positionLeftPixel = document.getElementById("cloneOval_" + cntClone).style.left;	
		var positionLeft = parseInt(positionLeftPixel.replace("px", ''));
		unitJson["positionLeftPercent"] = getPercentageValue(mainDivWidth, positionLeft);
		var positionTopPixel = document.getElementById("cloneOval_" + cntClone).style.top;	
		var positionTop = parseInt(positionTopPixel.replace("px", ''));
		unitJson["positionTopPercent"] = getPercentageValue(mainDivHeight, positionTop);
		var positionWidthPixel = document.getElementById("cloneOval_" + cntClone).style.width;	
		var positionWidth = parseInt(positionWidthPixel.replace("px", ''));
		unitJson["positionWidthPercent"] = getPercentageValue(mainDivWidth, positionWidth);
		var positionHeightPixel = document.getElementById("cloneOval_" + cntClone).style.height;	
		var positionHeight = parseInt(positionHeightPixel.replace("px", ''));
		unitJson["positionHeightPercent"] = getPercentageValue(mainDivWidth, positionHeight);
		unitJson["aspectRatio"] = ratio;
		
		unitJson["roleImage"] = document.getElementById(cloneId).getElementsByTagName("img")[0].getAttribute("src");
		var roleId = "role_" + cloneId;
		unitJson["roleId"] = roleId;
		
		var roleHeightPixel = document.getElementById(roleId).style.height;	
		var roleHeight = parseInt(roleHeightPixel.replace("px", ''));
		unitJson["roleHeightPercent"] = getPercentageValue(mainDivWidth, roleHeight);
		
		var roleWidthPixel = document.getElementById(roleId).style.width;	
		var roleWidth = parseInt(roleWidthPixel.replace("px", ''));
		unitJson["roleWidthPercent"] = getPercentageValue(mainDivWidth, roleWidth);
		
		var roleLeftPixel = document.getElementById(roleId).style.left;	
		var roleLeft = parseInt(roleLeftPixel.replace("px", ''));
		unitJson["roleLeftPercent"] = getPercentageValue(mainDivWidth, roleLeft);
		
		var roleTopPixel = document.getElementById(roleId).style.top;	
		var roleTop = parseInt(roleTopPixel.replace("px", ''));
		unitJson["roleTopPercent"] = getPercentageValue(mainDivHeight, roleTop);		
		
		totalJsonAftSktch[idx++] = unitJson;
		divClone = document.getElementById("cloneOval_" + ++cntClone);
	}
	ClientSession.set("total_json_after_sketch_final",JSON.stringify(totalJsonAftSktch)); //Total Json to be stored in db
}

/**
 * Function call while load the page
 * with the before sketch tasks
 * to set the task position with in the main div
 */
function setTaskWhileLoad() {
	 $("#outderDiv").children('.after_sketch_task:visible').each(function() {		
		 var maxLeftPosition = 0;
		 var leftPos = 0;
	        var element = $( this );       
	        elementId = $(this).attr('id');	  
	        var position = element.position();	
	        if(position.left + $(this).width() > ($("#outderDiv").width()-20)) {
	        	maxLeftPosition = (parseInt(position.left + $(this).width())) - (parseInt($("#outderDiv").width())-20);
	        	leftPos = (parseInt(position.left)) - maxLeftPosition - 10;
	        	 $("#" + elementId).css({ left : leftPos+'px'});
	        }	               		       
	    });	
}

/********************** ADDED FOR ROLE FRAME *****************************/

/**
 * Function add to fix the Role Frame with in the Div
 * @param cloneId
 */
function reDragRoleFrame(cloneId){	
	var roleId = "role_" +cloneId;
	var cloneImage = document.getElementById(cloneId).getElementsByTagName("img")[0].getAttribute("src");
	var roleTop = $("#" +roleId).outerHeight() + 20 ;
	if(($("#"+cloneId).position().top - $("#" +roleId).outerHeight())<20){
		 if ((cloneImage.search('left_top.png') != -1) || (cloneImage.search('right_top.png') != -1)) {
				$("#" +roleId).css({top:"20px"});	
				$("#"+cloneId).css({top: roleTop+"px"});	
			} else {
				$("#"+cloneId).css({top:"18px"});
			}
    } 
	 if ((cloneImage.search('left_bottom.png') != -1) || (cloneImage.search('right_bottom.png') != -1)) {
			var roleTop = $("#outderDiv").height() - $("#"+cloneId).height() - $("#" +roleId).height() - 20 ;
			var roleTextTop = $("#outderDiv").height() - $("#" +roleId).height() - 20;	
			if (($("#"+cloneId).position().top + $("#"+cloneId).height() + $("#" +roleId).outerHeight() + 15) > $("#outderDiv").height()){
				$("#" +roleId).css({top:(roleTextTop-20)+"px"});	
				$("#"+cloneId).css({top: (roleTop-20)+"px"});	
			}						
	}
	 var cloneLeft = $("#"+cloneId).position().left;
	 if(cloneLeft+$("#"+cloneId).width() > $("#outderDiv").width()) {
		 var newLeft = (cloneLeft+$("#"+cloneId).width()) - $("#outderDiv").width();
		 $("#"+cloneId).css({left:(cloneLeft-newLeft)+"px"});
	 }
    
    fixRoleFramePosition(cloneId);
   // setTimeout(function(){fixRoleFramePosition(cloneId);},500);	
}


$("body").mouseup(function(){
	if(($('#outderDiv').children('.oval_cloned_div:visible').length > 0)) {
		$("#outderDiv").children('.oval_cloned_div:visible').each(function() {			
		var cloneId = $(this).attr('id');
		//fixRoleFramePosition(cloneId);		 
		});
	}
});

function fixRoleFramePosition(cloneId) {
	var cloneImage = document.getElementById(cloneId).getElementsByTagName("img")[0].getAttribute("src");
		var position = $("#" +cloneId).position();
		var roleId = "role_" +cloneId;
		//if($("#" + roleId).children(".ui-resizable-ne").css("display") == "block") {
			
		if (cloneImage.search('left_top.png') != -1) {
		//var topPos = position.top - $("#"+roleId).height() -12;
		var topPos = position.top - $("#"+roleId).outerHeight();
		var leftPos = position.left;
		//alert(leftPos);
		if(leftPos < 0) {
			var leftPosVal = -(leftPos);
			$("#"+roleId).animate({
				 top: topPos+"px",
			    left: "0px"
			  });
			$("#"+cloneId).animate({
			    left: leftPosVal+"px"
			  });
			//alert(rolePosition.left);
		} else if (leftPos + $("#"+roleId).width() + 15 > $("#outderDiv").width()) {
			var newLeft = (leftPos + $("#"+roleId).width() + 15) - $("#outderDiv").width();
			var cloneLeft = leftPos - newLeft;
			$("#"+roleId).animate({
			    top: topPos+"px",
			    left: cloneLeft+"px"
			  }, 1);
			$("#"+cloneId).animate({
			    left: cloneLeft+"px"
			  });
		}
		else {
			$("#"+roleId).animate({
			    top: topPos+"px",
			    left: leftPos+"px"
			  }, 1);
		}		
	} 
	//else if($("#" + roleId).children(".ui-resizable-nw").css("display") == "block"){
	else if(cloneImage.search('right_top.png') != -1){				
		//var topPos = position.top - $("#"+roleId).height() -12;
		var topPos = position.top - $("#"+roleId).outerHeight();
		var leftPos = position.left + $("#"+cloneId).width() - $("#"+roleId).outerWidth();
		//alert(leftPos);
		if(leftPos < 0) {
			var leftPosVal = -(leftPos);
			$("#"+roleId).animate({
				 top: topPos+"px",
			    left: "0px"
			  });
			$("#"+cloneId).animate({
			    left: leftPosVal+"px"
			  });
			//alert(rolePosition.left);
		} else {
		$("#"+roleId).animate({
				    top: topPos+"px",
				    left: leftPos+"px"
				  }, 1);
		}
	} 
	//else if($("#" + roleId).children(".ui-resizable-se").css("display") == "block"){
	else if(cloneImage.search('left_bottom.png') != -1){
		var topPos = position.top + $("#"+cloneId).height();
		var leftPos = position.left;
		//alert(leftPos);
		if(leftPos < 0) {
			var leftPosVal = -(leftPos);
			$("#"+roleId).animate({
				 top: topPos+"px",
			    left: "0px"
			  });
			$("#"+cloneId).animate({
			    left: leftPosVal+"px"
			  });
		} else if (leftPos + $("#"+roleId).width() + 15 > $("#outderDiv").width()) {
			var newLeft = (leftPos + $("#"+roleId).width() + 15) - $("#outderDiv").width();
			var cloneLeft = leftPos - newLeft;
			$("#"+roleId).animate({
			    top: topPos+"px",
			    left: cloneLeft+"px"
			  }, 1);
			$("#"+cloneId).animate({
			    left: cloneLeft+"px"
			  });
		} else {
		$("#"+roleId).animate({
			    top: topPos+"px",
			    left: leftPos+"px"
			  }, 1);
	}
	} 
	//else if($("#" + roleId).children(".ui-resizable-sw").css("display") == "block"){
	else if(cloneImage.search('right_bottom.png') != -1){
		var topPos = position.top + $("#"+cloneId).height();
		//var leftPos = position.left + $("#"+cloneId).width() - $("#"+roleId).width() - 12;
		var leftPos = position.left + $("#"+cloneId).width() - $("#"+roleId).outerWidth();
	//	alert(leftPos);
		if(leftPos < 0) {
			var leftPosVal = -(leftPos);
			$("#"+roleId).animate({
				 top: topPos+"px",
			    left: "0px"
			  });
			$("#"+cloneId).animate({
			    left: leftPosVal+"px"
			  });
		}/* else if (leftPos > 725) {
			
		}*/
		else {
		$("#"+roleId).animate({
				    top: topPos+"px",
				    left: leftPos+"px"
				  }, 1);
	}
 }			
}


/**
 * FUnction to change the role frame image
 * and to move the speech bubble
 * @param roleId
 * @param actionPosition
 */
function moveSpeechBubble(roleId,  actionPosition) {
	var divCnt = roleId.substring(15);
	var cloneId = "cloneOval_" +divCnt;
	sessionStorage.setItem("RoleFrameId",cloneId);
	var cloneImage = document.getElementById(cloneId).getElementsByTagName("img")[0].getAttribute("src");
    var cloneIdImg = sessionStorage.getItem("RoleFrameId");
	var roleId = "role_" +cloneIdImg;
	var position = $("#" +cloneIdImg).position();
	/**
	 * When the role frame is in the top left side
	 */
    if (cloneImage.search('left_top.png') != -1) {
    	if(actionPosition == "LEFT") {    
    		/**
    		 * When user click on left arrow 
    		 * from top left role frame
    		 */
    		document.getElementById(roleId).getElementsByTagName("img")[0].setAttribute("src", "../img/right-arrow.png");
    		setTopRightImage(roleId, cloneIdImg);  		 
    		setTimeout(function(){roleFrameUp(roleId, cloneIdImg);},500);  		
    		setTimeout(function(){document.getElementById(cloneIdImg).getElementsByTagName("img")[0].setAttribute("src", "../img/right_top.png");},500);	
    	} else {
    		/**
    		 * When user click on down arrow 
    		 * from top left role frame
    		 */
    		document.getElementById(roleId).getElementsByTagName("img")[1].setAttribute("src", "../img/up-arrow.png");
    		setBottomLeftImage(roleId,cloneIdImg);
    		roleFrameDown(roleId, cloneIdImg);
    		setTimeout(function(){document.getElementById(cloneIdImg).getElementsByTagName("img")[0].setAttribute("src", "../img/left_bottom.png");},500);	
    		//document.getElementById(cloneIdImg).getElementsByTagName("img")[0].setAttribute("src", "../img/left_bottom.png");	
    	}    	    	
		roleFrameResize(roleId, cloneIdImg);
    }
    /**
	 * When the role frame is in the top right side
	 */
    else if (cloneImage.search('right_top.png') != -1) {
    	if(actionPosition == "LEFT") {   
    		/**
    		 * When user click on right arrow 
    		 * from top right role frame
    		 */
    		document.getElementById(roleId).getElementsByTagName("img")[0].setAttribute("src", "../img/left-arrow.png");
    		setTopLeftImage(roleId, cloneIdImg);    		    		
    		setTimeout(function(){roleFrameUp(roleId, cloneIdImg);},500);
    		//document.getElementById(cloneIdImg).getElementsByTagName("img")[0].setAttribute("src", "../img/left_top.png");
    		setTimeout(function(){document.getElementById(cloneIdImg).getElementsByTagName("img")[0].setAttribute("src", "../img/left_top.png");},500);	
    	} else {
    		/**
    		 * When user click on down arrow 
    		 * from top right role frame
    		 */
    		document.getElementById(roleId).getElementsByTagName("img")[1].setAttribute("src", "../img/up-arrow.png");
    		setBottomRightImage(roleId,cloneIdImg);
    		roleFrameDown(roleId, cloneIdImg);   
    		//document.getElementById(cloneIdImg).getElementsByTagName("img")[0].setAttribute("src", "../img/right_bottom.png");  
    		setTimeout(function(){document.getElementById(cloneIdImg).getElementsByTagName("img")[0].setAttribute("src", "../img/right_bottom.png");},500);	
    	}   	
		roleFrameResize(roleId, cloneIdImg);
    }
    /**
	 * When the role frame is in the bottom left side
	 */
    else if (cloneImage.search('left_bottom.png') != -1) {
    	if(actionPosition == "LEFT") {
    		/**
    		 * When user click on right arrow 
    		 * from bottom left role frame
    		 */
    		document.getElementById(roleId).getElementsByTagName("img")[0].setAttribute("src", "../img/right-arrow.png");
    		setBottomRightImage(roleId,cloneIdImg);
    		setTimeout(function(){roleFrameDown(roleId, cloneIdImg);},500);  	
    		//document.getElementById(cloneIdImg).getElementsByTagName("img")[0].setAttribute("src", "../img/right_bottom.png"); 
    		setTimeout(function(){document.getElementById(cloneIdImg).getElementsByTagName("img")[0].setAttribute("src", "../img/right_bottom.png");},500);	
    	} else {
    		/**
    		 * When user click on up arrow 
    		 * from bottom left role frame
    		 */
    		document.getElementById(roleId).getElementsByTagName("img")[0].setAttribute("src", "../img/left-arrow.png");
    		document.getElementById(roleId).getElementsByTagName("img")[1].setAttribute("src", "../img/down-arrw.png");
    		setTopLeftImage(roleId, cloneIdImg);
    		setTimeout(function(){roleFrameUp(roleId, cloneIdImg);},500);  		
    		//document.getElementById(cloneIdImg).getElementsByTagName("img")[0].setAttribute("src", "../img/left_top.png");
    		setTimeout(function(){document.getElementById(cloneIdImg).getElementsByTagName("img")[0].setAttribute("src", "../img/left_top.png");},500);	
    	}    	    	
		roleFrameResize(roleId, cloneIdImg);	   	
	}
    /**
	 * When the role frame is in the bottom right side
	 */
    else if (cloneImage.search('right_bottom.png') != -1) {
		if(actionPosition == "LEFT") {
			/**
    		 * When user click on left arrow 
    		 * from bottom right role frame
    		 */
			document.getElementById(roleId).getElementsByTagName("img")[0].setAttribute("src", "../img/left-arrow.png");
			setBottomLeftImage(roleId,cloneIdImg);
			roleFrameDown(roleId, cloneIdImg);	
			//document.getElementById(cloneIdImg).getElementsByTagName("img")[0].setAttribute("src", "../img/left_bottom.png");
			setTimeout(function(){document.getElementById(cloneIdImg).getElementsByTagName("img")[0].setAttribute("src", "../img/left_bottom.png");},500);	
		} else {
			/**
    		 * When user click on up arrow 
    		 * from bottom right role frame
    		 */
			document.getElementById(roleId).getElementsByTagName("img")[1].setAttribute("src", "../img/down-arrw.png");
			setTopRightImage(roleId,cloneIdImg);						
			setTimeout(function(){roleFrameUp(roleId, cloneIdImg);},500);	
			setTimeout(function(){document.getElementById(cloneIdImg).getElementsByTagName("img")[0].setAttribute("src", "../img/right_top.png");},500);
		}						
		roleFrameResize(roleId, cloneIdImg);
		
	}
	 fixRoleFramePosition(cloneId);
}
	
/**
 * Function call to set the role frame
 * on top left position
 * @param roleId
 * @param cloneIdImg
 */
function setTopLeftImage(roleId,cloneIdImg) {
	var position = $("#" +cloneIdImg).position();
    var topPos = position.top - $("#"+roleId).outerHeight();
	var leftPos = position.left;
	$("#"+roleId).animate({
		top: topPos+"px",
		left: leftPos+"px"
	}); 		
	setNEResizableTriangle(roleId);
	document.getElementById(cloneIdImg).getElementsByTagName("img")[0].setAttribute("src", "../img/common_role_image.png");	
}

/**
 * Function call to set the role frame
 * on top right position
 * @param roleId
 * @param cloneIdImg
 */
function setTopRightImage(roleId,cloneIdImg) {
	var position = $("#" +cloneIdImg).position();
	var topPos = position.top - $("#"+roleId).outerHeight();
	var leftPos = position.left + $("#"+cloneIdImg).width() - $("#"+roleId).outerWidth();
	$("#"+roleId).animate({
		top: topPos+"px",
		left: leftPos+"px"
	}); 
	setNWResizableTriangle(roleId);	
	document.getElementById(cloneIdImg).getElementsByTagName("img")[0].setAttribute("src", "../img/common_role_image.png");		
}
	
/**
 * Function call to set the role frame
 * on bottom left position
 * @param roleId
 * @param cloneIdImg
 */
function setBottomLeftImage(roleId,cloneIdImg) {
	var position = $("#" +cloneIdImg).position();
	var topPos = position.top + $("#"+cloneIdImg).height();
	var leftPos = position.left;
	$("#"+roleId).animate({
		top: topPos+"px",
		left: leftPos+"px"
	}); 		
	setSEResizableTriangle(roleId);
	document.getElementById(cloneIdImg).getElementsByTagName("img")[0].setAttribute("src", "../img/common_role_image.png");	
}

/**
 * Function call to set the role frame
 * on bottom right position
 * @param roleId
 * @param cloneIdImg
 */
function setBottomRightImage(roleId,cloneIdImg) {
	var position = $("#" +cloneIdImg).position();
	var topPos = position.top + $("#"+cloneIdImg).height();
	var leftPos = position.left + $("#"+cloneIdImg).width() - $("#"+roleId).outerWidth();
	$("#"+roleId).animate({
		top: topPos+"px",
		left: leftPos+"px"
	}); 
	setSWResizableTriangle(roleId);
	document.getElementById(cloneIdImg).getElementsByTagName("img")[0].setAttribute("src", "../img/common_role_image.png");			
}
	
/**
 * 
 * @param cloneId
 */
function enableSpeechBubble(cloneId) {
	var roleId = "role_" +cloneId;
	var textAreaId = "text_" +cloneId;
	var deleteElementId = cloneId+"_x"; 
	//var $element = $('<div class="speech-bubble common_elem" id="'+ roleId +'"><div class="role_text" >Role</div><div id= "clickableElement1" class="clickable_element_1" onclick = moveSpeechBubble("'+roleId+'","'+"LEFT"+'")><img src="../img/left-arrow.png"></div><div id= "clickableElement2" class="clickable_element_2" onclick = moveSpeechBubble("'+roleId+'","'+"DOWN"+'")><img src="../img/down-arrw.png"></div><textarea id = "'+textAreaId+'" onkeypress="return disableKey(event)" maxlength="120"></textarea></div>');
	var $element = $('<div class="speech-bubble common_elem commonDelete" id="'+ roleId +'"><div class="role_text" >Role</div><div id= "clickableElement1" class="clickable_element_1" onclick = moveSpeechBubble("'+roleId+'","'+"LEFT"+'")><img src="../img/left-arrow.png"></div><div id= "clickableElement2" class="clickable_element_2" onclick = moveSpeechBubble("'+roleId+'","'+"DOWN"+'")><img src="../img/down-arrw.png"></div><textarea id = "'+textAreaId+'" onkeypress="return disableKey(event)" maxlength="120"></textarea><div id="' + deleteElementId +'" class="removeService delete_oval" onclick=deleteElement(this)><img src="../img/cross-black.png" alt="Delete" /></div></div>');
	$("#outderDiv").append($element);
	if(!(navigator.userAgent.match(/iPhone|Android/i))){
		$("#" + roleId).children(".role_text").css({"top" : "-20px"});
	} else {
		$("#" + roleId).children(".role_text").css({"top" : "-12px"});
	}
	var position = $("#" +cloneId).position();
	var topPos = position.top - $("#"+roleId).outerHeight();
	var leftPos = position.left;
	$element.css({
			    top: topPos+"px",
			    left: leftPos+"px"
			  }, 200); 	
	
	$("#"+textAreaId).keyup(function(e) { 
		var textAreaId = $(this).attr("id"); //text_
		var cloneId = textAreaId.replace("text_", "");
		fixTextAreaSize(cloneId, e);	
	});       

}


function fixTextAreaSize(cloneId, e) {
	var roleId = "role_" +cloneId;
	var textAreaId = "text_" +cloneId;	
	var container = document.getElementById(textAreaId);
	
	var roleWidthInit = $(".speech-bubble").width();
	var roleWidth = $("#"+roleId).width(); //118
	var content = $.trim($("#"+textAreaId).val());
	 
    var additionalWidth = (container.scrollHeight) + parseFloat(($("#"+textAreaId)).css("borderTopWidth")) + parseFloat(($("#"+textAreaId)).css("borderBottomWidth")) - ($("#"+textAreaId).outerHeight());   
     if(additionalWidth>=0){
    	 $("#"+textAreaId).width(($("#"+textAreaId)).width()+parseInt(additionalWidth));
    	    $("#"+roleId).width($("#"+textAreaId).width());
     }
    

    if(content.length === 0 || content.length == 1) {
      	$("#"+roleId).width(roleWidthInit);
    	$("#"+textAreaId).width(roleWidthInit);
         } 
        else {
    	if (e.keyCode === 8 || e.keyCode === 46) {     
    		$("#"+textAreaId).css('width', roleWidth);
     	    $("#"+roleId).css('width', roleWidth);
     	   while($("#"+textAreaId).outerHeight() < container.scrollHeight + parseFloat(($("#"+textAreaId)).css("borderTopWidth")) + parseFloat(($("#"+textAreaId)).css("borderBottomWidth"))) {
     	         $("#"+textAreaId).width(($("#"+textAreaId)).width()+1);
     	    	 $("#"+roleId).width($("#"+textAreaId).width());
     	    }
        }
    }
   
    fixRoleFramePosition(cloneId);      
}
/**
 * 
 * @param roleId
 * @param cloneId
 * NOT IN USED
 */
function roleFrameResize(roleId, cloneId) {
	if($("#" + roleId).children(".ui-resizable-ne").css("display") == "block") {
	 $("#" + roleId).resizable({	
		 handles: {
	       'ne': '#negrip'
	    }
		}); 
	 $("#" + roleId).resizable({		 
			stop:function(event,ui){			
				var position = $("#" +cloneId).position();
				var topPos = position.top - $("#"+roleId).height() - 12;
				var leftPos = position.left;
				$("#"+roleId).css({
					top: topPos+"px",
					left: leftPos+"px"
				}, 200); 
				roleFrameUp(roleId, cloneId);			
		}
		});	
	} else if ($("#" + roleId).children(".ui-resizable-nw").css("display") == "block"){
		$("#" + roleId).resizable({	
		 handles: {
	       'nw': '#nwgrip'       
	    }
		}); 
		$("#" + roleId).resizable({		 
			stop:function(event,ui){			
				var position = $("#" +cloneId).position();
				var topPos = position.top - $("#"+roleId).height() - 12;
				var leftPos = position.left + $("#"+cloneId).width() - $("#"+roleId).width() - 12;
				$("#"+roleId).css({
					top: topPos+"px",
					left: leftPos+"px"
				}, 200); 
				roleFrameUp(roleId, cloneId);
		}
		});		
	} else if($("#" + roleId).children(".ui-resizable-se").css("display") == "block"){
		$("#" + roleId).resizable({	
		 handles: {
	       'se': '#segrip'       
	    }
		}); 
		 $("#" + roleId).resizable({		 
				stop:function(event,ui){			
					var position = $("#" +cloneId).position();
					var topPos = position.top + $("#"+cloneId).height();
					var leftPos = position.left;
					$("#"+roleId).css({
						top: topPos+"px",
						left: leftPos+"px"
					}, 200); 
					roleFrameDown(roleId, cloneId);	
			}
			});	
	} else if($("#" + roleId).children(".ui-resizable-sw").css("display") == "block"){
		$("#" + roleId).resizable({	
		 handles: {
	       'sw': '#swgrip'       
	    }
		}); 
		$("#" + roleId).resizable({		 
			stop:function(event,ui){			
				var position = $("#" +cloneId).position();
				var topPos = position.top + $("#"+cloneId).height();
				var leftPos = position.left + $("#"+cloneId).width() - $("#"+roleId).width() - 12;
				$("#"+roleId).css({
					top: topPos+"px",
					left: leftPos+"px"
				}, 200); 
				roleFrameDown(roleId, cloneId);
		}
		});	
	}			
}

/**
 * 
 * @param roleId
 * @param cloneId
 */
function roleFrameDown(roleId, cloneId) {
	var pos =  $("#" +cloneId).position();
	if( $("#outderDiv").height() < (pos.top + $("#" +cloneId).height() +  $("#" +roleId).height() + 60)) {
		var maxHeight = pos.top + $("#" +cloneId).height() +  $("#" +roleId).height() + 80;
		$("#pageWrapCopy").css({'height':(maxHeight+'px')});
		$("#outderDiv").css({'height':($("#pageWrapCopy").height()+'px')});
	}
}

/**
 * 
 * @param roleId
 * @param cloneId
 */
function roleFrameUp(roleId, cloneId) {
	var positionRole = $("#" +roleId).position();
	//var cloneTop = $("#" +roleId).height() + 20 + 12 ;
	var cloneTop = $("#" +roleId).outerHeight() + 20 ;
	if(positionRole.top <18) {
		$("#" +roleId).animate({top:"20px"}, 100);
		$("#" +cloneId).animate({top: cloneTop+"px"}, 100);	
	}		
	 if( $("#outderDiv").height() < ($("#" +cloneId).height() +  $("#" +roleId).height() + 30)) {
		var maxHeight = $("#" +cloneId).height() +  $("#" +roleId).height() + 40;
		$("#pageWrapCopy").css({'height':(maxHeight+'px')});
		$("#outderDiv").css({'height':($("#pageWrapCopy").height()+'px')});
	 }
}
	
/**
 * Function call when the page is loaded with the final JSON
 * This function set the arrow images, display the resizable triangle
 * and fix the "Role" text position depending on the role frame image
 * @param cloneImage
 * @param roleId
 */	
function setResizableCorner(cloneImage, roleId) {
	/*$("#" + roleId).resizable({	
		 handles: {
	        'nw': '#nwgrip',
	        'ne': '#negrip',
	        'sw': '#swgrip',
	        'se': '#segrip'
	    }
	}); */
	if (cloneImage.search('left_top.png') != -1) {	
		setNEResizableTriangle(roleId);
		document.getElementById(roleId).getElementsByTagName("img")[0].setAttribute("src", "../img/left-arrow.png");
		document.getElementById(roleId).getElementsByTagName("img")[1].setAttribute("src", "../img/down-arrw.png");
	 } else if (cloneImage.search('right_top.png') != -1) {
		 setNWResizableTriangle(roleId);	 
		document.getElementById(roleId).getElementsByTagName("img")[0].setAttribute("src", "../img/right-arrow.png");
		document.getElementById(roleId).getElementsByTagName("img")[1].setAttribute("src", "../img/down-arrw.png");			
	 } else if (cloneImage.search('left_bottom.png') != -1) {		 
		 setSEResizableTriangle(roleId);		
		document.getElementById(roleId).getElementsByTagName("img")[0].setAttribute("src", "../img/left-arrow.png");
		document.getElementById(roleId).getElementsByTagName("img")[1].setAttribute("src", "../img/up-arrow.png");			
	 } else if (cloneImage.search('right_bottom.png') != -1) {		 
		setSWResizableTriangle(roleId);		 
		document.getElementById(roleId).getElementsByTagName("img")[0].setAttribute("src", "../img/right-arrow.png");
		document.getElementById(roleId).getElementsByTagName("img")[1].setAttribute("src", "../img/up-arrow.png");
	 }
	var cloneId = roleId.substr(5);
	roleFrameResize(roleId, cloneId);
	fixRoleFramePosition(cloneId);
}

/**
 * Function call to display the top left resizable triangle
 * and fix the "Role" text position
 * Function call when the page is loaded
 * @param roleId
 */
function setNEResizableTriangle(roleId) {
	$("#" + roleId).children(".ui-resizable-ne").css("display","block");
	$("#" + roleId).children(".ui-resizable-nw").css("display","none");							
	$("#" + roleId).children(".ui-resizable-sw").css("display","none");
	$("#" + roleId).children(".ui-resizable-se").css("display","none");	
	if(!(navigator.userAgent.match(/iPhone|Android/i))){
		$("#" + roleId).children(".role_text").css({"top" : "-20px"});
	} else {
		$("#" + roleId).children(".role_text").css({"top" : "-12px"});
	}
	$("#" + roleId).children(".role_text").css({"bottom" : " "});
}

/**
 * Function call to display the top right resizable triangle
 * and fix the "Role" text position
 * Function call when the page is loaded
 * @param roleId
 */
function setNWResizableTriangle(roleId) {
	$("#" + roleId).children(".ui-resizable-nw").css("display","block");
	$("#" + roleId).children(".ui-resizable-ne").css("display","none");
	$("#" + roleId).children(".ui-resizable-sw").css("display","none");
	$("#" + roleId).children(".ui-resizable-se").css("display","none");	
	if(!(navigator.userAgent.match(/iPhone|Android/i))){
		$("#" + roleId).children(".role_text").css({"top" : "-20px"});
	} else {
		$("#" + roleId).children(".role_text").css({"top" : "-12px"});
	}
	$("#" + roleId).children(".role_text").css({"bottom" : " "});
}

/**
 * Function call to display the left bottom resizable triangle
 * and fix the "Role" text position
 * Function call when the page is loaded
 * @param roleId
 */
function setSEResizableTriangle(roleId) {
	$("#" + roleId).children(".ui-resizable-se").css("display","block");	
	$("#" + roleId).children(".ui-resizable-nw").css("display","none");							
	$("#" + roleId).children(".ui-resizable-sw").css("display","none");
	$("#" + roleId).children(".ui-resizable-ne").css("display","none");	 
	$("#" + roleId).children(".role_text").css({"top" : ""});
	if(!(navigator.userAgent.match(/iPhone|Android/i))){
		$("#" + roleId).children(".role_text").css({"bottom" : "-20px"});
	} else {
		$("#" + roleId).children(".role_text").css({"bottom" : "-15px"});
	}
}

/**
 * Function call to display the right bottom resizable triangle
 * and fix the "Role" text position
 * Function call when the page is loaded
 * @param roleId
 */
function setSWResizableTriangle(roleId) {
	$("#" + roleId).children(".ui-resizable-sw").css("display","block");
	$("#" + roleId).children(".ui-resizable-nw").css("display","none");							
	$("#" + roleId).children(".ui-resizable-se").css("display","none");
	$("#" + roleId).children(".ui-resizable-ne").css("display","none");	 
	
	$("#" + roleId).children(".role_text").css({"top" : ""});
	if(!(navigator.userAgent.match(/iPhone|Android/i))){
		$("#" + roleId).children(".role_text").css({"bottom" : "-20px"});
	} else {
		$("#" + roleId).children(".role_text").css({"bottom" : "-15px"});
	}
}

/**
 * Function call to hide the resizable traiangle of speech bubble
 * This function is called for after sketch edit only 
 * @param roleId
 */
function hideAllResizableTriangle(roleId) {
	$("#" + roleId).children(".ui-resizable-sw").css("display","none");
	$("#" + roleId).children(".ui-resizable-nw").css("display","none");							
	$("#" + roleId).children(".ui-resizable-se").css("display","none");
	$("#" + roleId).children(".ui-resizable-ne").css("display","none");	 
}

function fixTextAreaSize1(cloneId, e) {
	
	var roleId = "role_" +cloneId;
	var textAreaId = "text_" +cloneId;
	var container = document.getElementById(textAreaId);
	
	//var roleWidth = $("#"+roleId).width();
	var roleWidth = 80;
	var content = $.trim($("#"+textAreaId).val());
	
    //  the following will help the text expand as typing takes place
    while($("#"+textAreaId).outerHeight() < container.scrollHeight + parseFloat(($("#"+textAreaId)).css("borderTopWidth")) + parseFloat(($("#"+textAreaId)).css("borderBottomWidth"))) {
         $("#"+textAreaId).width(($("#"+textAreaId)).width()+1);
    	 $("#"+roleId).width($("#"+textAreaId).width());
    }
    
    if (e.keyCode === 8 || e.keyCode === 46 || content.length === 0) {         
        $("#"+textAreaId).css('width', roleWidth);
 	    $("#"+roleId).css('width', roleWidth);
 	   while($("#"+textAreaId).outerHeight() < container.scrollHeight + parseFloat(($("#"+textAreaId)).css("borderTopWidth")) + parseFloat(($("#"+textAreaId)).css("borderBottomWidth"))) {
 	         $("#"+textAreaId).width(($("#"+textAreaId)).width()+1);
 	    	 $("#"+roleId).width($("#"+textAreaId).width());
 	    }
    
    }
    fixRoleFramePosition(cloneId);      
}


/* Function to 
 * adjust textarea for devices for first time login after logging out from desktop
 * */
function fixRoleFrameText(cloneId){
	var roleId = "role_" +cloneId;
	var textAreaId = "text_" +cloneId;
	var container = document.getElementById(textAreaId);		
	var additionalWidth = (container.scrollHeight) + parseFloat(($("#"+textAreaId)).css("borderTopWidth")) + parseFloat(($("#"+textAreaId)).css("borderBottomWidth")) - ($("#"+textAreaId).outerHeight());   
    $("#"+textAreaId).width(($("#"+textAreaId)).width()+parseInt(additionalWidth) + 5);
    $("#"+roleId).width($("#"+textAreaId).width());
}

/**
 * Function add to restrict ~ and `
 * @param evt
 */
function disableKey(evt) {
    var charCode = (evt.which) ? evt.which : event.keyCode;
    if (charCode == 96 || charCode == 126)
        return false;

    return true;
}
