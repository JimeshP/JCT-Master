var pageTime = 0;

/**
 * Function call while refresh the page
 */  
window.onunload = function(event) {
	return saveBeforeSketchWhileRefresh();
}; 

window.setInterval(function(){
	updateTime();
}, 60000);

window.setInterval(function(){
	updateCurrentPageTime();
}, 1000);

function updateCurrentPageTime() {
	pageTime = pageTime + 1;
	sessionStorage.setItem("pageTime", pageTime);
	//console.log(pageTime);
}
/**
 * Function add to populate the instruction description
 * while the page is loaded
 * @param null
 */
$(document).ready(function() {
	sessionStorage.setItem("pageTime", 0);
	
	sessionStorage.setItem("pageSequence", 1);
	populateInstruction();
	if (sessionStorage.getItem("giveUserOption")) {
		document.getElementById("restartLinkId").style.display = "block";
	}
});

/************** Element will be drag with in "#pageWrap" element ****************/
$(document).ready(function() {
	if(sessionStorage.getItem("isCompleted") == 0){
		if (sessionStorage.getItem("bsEdit") || (null == sessionStorage.getItem("bsEdit"))) {
		$('.before_sketch_task').draggable({
			cursor: 'move',       
			containment: '#pageWrap'    
		}); 
		var windowSize = $(window).width();  
			if (windowSize <= 1024) { 
				$(".before_sketch_task").on('touchstart', function(e){
		    		enableDraggableResizable($(this).attr('id'));		    	
				});
			}
		}
	}
});

/**
 * Function add to disable browser back button
 * while the page is loaded
 * @param null
 */
window.location.hash="";
window.location.hash="Again-No-back-button";//again because google chrome don't insert first hash into history
window.onhashchange=function(){window.location.hash="";};


var windowSize = $(window).width(); 

/********** Added for instruction toggle ***********/
    $(".btn-slide").click(function(){
    $("#panel").slideToggle("slow");
    $(this).toggleClass("active"); 
    return false;
  });

var snapShotURL = "";
var allCanvasArray = new Array();
var count = 1;
/**
 * function handles the add task
 */
$('#addtask').click(function() {	
	var unlockDivCount = $('#pageWrap').children('.single_task_item:visible').length;
	if(sessionStorage.getItem("isCompleted") == 1){
		alertify.alert("<img src='../img/server-na-icon.png'><br /><p>Cannot add tasks as you have already completed your after diagram.</p>");
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
	var divCount = $('#pageWrap').children('.before_sketch_task:visible').length;
	if(divCount>=20){
		alertify.alert("<img src='../img/server-na-icon.png'><br /><p>You can have a maximum of 20 tasks.</p>");
		return false;
	}

	/****************** Check the Energy value of each Task ********************/	
	var array = new Array();
	var divCount = $('#pageWrap').children('.before_sketch_task:visible').length;
	var divCountLoop = $('#pageWrap').children('.before_sketch_task').length;
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
	var lockDivCount = $('#pageWrap').children('.lock_single_task_item:visible').length;
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
	var mainDivWidth = $('#pageWrap').width();
	if (!(navigator.userAgent.match(/iPhone|iPad|iPod/i)) && !(navigator.userAgent.match(/Android/i))) {
		if(mainDivWidth < 1000) {
			mainDivWidth = 1021;
		}
	 } 
	
	/****************** When the energy values are same ********************/
	if(check==true){			
		var existingDivId = $('#pageWrap').children('.before_sketch_task:visible').attr('id');		
		var divCnt = existingDivId.substring(7);	    			
		var previousWidth = document.getElementById(existingDivId).style.width;
		var previousEnergyVal = document.getElementById("energyId_" +divCnt).value;	
		var previousWidthValue = parseInt(previousWidth.replace("px",''));	
		totalDiv = divCount+1;
		energyValue = Math.round(100/totalDiv);
		if(mainDivWidth >1000) {
			energyStyle = previousWidthValue - (parseInt(previousEnergyVal - energyValue)*3);	
		} else {
			var diffValue = getWidthDifference(parseInt(previousEnergyVal - energyValue)*3, mainDivWidth);	
			energyStyle = previousWidthValue - diffValue;
		}	
	    if(next_page_nav == 0){
	    	count = ClientSession.get("total_count");
	    	var divId = "divImg_" + count;
			var divIdx = "divImg_" + count+"_x";		
			var energyId = "energyId_" + count;
			var areaId = "areaId_" + count;	
			/**
			 * Added for lock and unlock feature
			 */			
			var divLockId = "divImg_" + count+"_lock";
			/******* END *******/
			if(navigator.userAgent.match(/Android/i) || navigator.userAgent.match(/iPhone/i)){
				var $element = $('<div id="' + divId +'" class="col-md-4 single_task_item before_sketch_task draggableResizable" style="display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top"><div class="col-md-3 col-xs-3 task" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 energy" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/ Energy:&nbsp;</div><div class="col-md-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="col-md-12 add_task_area" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divLockId+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');
			}  else if(navigator.userAgent.match(/iPad/i)){
				var $element = $('<div id="' + divId +'" class="col-md-4 single_task_item before_sketch_task draggableResizable" style="display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top"><div class="col-md-3 col-xs-3 task" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 energy" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/Energy:&nbsp;</div><div class="col-md-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="col-md-12 add_task_area" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divLockId+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');
			} 	else {
				var $element = $('<div id="' + divId +'" class="col-md-4 single_task_item before_sketch_task draggableResizable" style="display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top"><div class="col-md-3 col-xs-3 task">TASK</div><div class="col-md-6 col-xs-5 energy">Time/Energy:&nbsp;</div><div class="col-md-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)" /></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="col-md-12 add_task_area" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)" ></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divLockId+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');
			}
			
		 } else {
			 count = ClientSession.get("total_count");
			 var divId = "divImg_" + count;
			 var divIdx = "divImg_" + count+"_x";		
			 var energyId = "energyId_" + count;
			 var areaId = "areaId_" + count;
			 /**
			  * Added for lock and unlock feature
			 */			
			 var divLockId = "divImg_" + count+"_lock";
			 if(navigator.userAgent.match(/Android/i) || navigator.userAgent.match(/iPhone/i)){
				 var $element = $('<div id="' + divId +'" class="col-md-4 single_task_item before_sketch_task draggableResizable" style="display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top"><div class="col-md-3 col-xs-3 task" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 energy" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/ Energy: </div><div class="col-md-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="col-md-12 add_task_area" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divLockId+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');		
			 }  else if(navigator.userAgent.match(/iPad/i)){
				 var $element = $('<div id="' + divId +'" class="col-md-4 single_task_item before_sketch_task draggableResizable" style="display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top"><div class="col-md-3 col-xs-3 task" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 energy" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/Energy:&nbsp; </div><div class="col-md-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="col-md-12 add_task_area" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divLockId+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');			
			 }	else {
				 var $element = $('<div id="' + divId +'" class="col-md-4 single_task_item before_sketch_task draggableResizable" style="display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top"><div class="col-md-3 col-xs-3 task">TASK</div><div class="col-md-6 col-xs-5 energy">Time/Energy:&nbsp; </div><div class="col-md-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)" /></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="col-md-12 add_task_area" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divLockId+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');
			 }
			}				
	  	 count++;	  	
	  	 ClientSession.set("total_count", count);
	  	 calculateEnergySameValue();
	  	 /************** To remove the focus to the new old task *************/
	  	 removeFocus();
	  	 
	     $("#pageWrap").append($element);		     	     	 		 		
	     $element.draggable();
				$('.before_sketch_task').draggable({
					cursor: 'move',       
					containment: '#pageWrap'    
				});  								
	     $element.resizable({		
				//ghost: 'true',
				handles: 'se',
				aspectRatio: 2 / 2,
			});			
	  
	     allCanvasArray.push($element);
	     $(".single_task_item").css({
		    width: energyStyle+"px",
		    height: energyStyle+"px"
		  }, 200); 	
	     
	     addTaskPosition(energyStyle,divId);			//		Newly Added		
	     /***************** To set the focus to the new Task ***************/
	     $("#" + divId).children().find(".add_task_area").addClass("add_task_area_focus");
	     $("#" + areaId).focus();
	     
	     /***************** To Disable the paste functionality ISSUE# 186***************/
	     $('#' +areaId).bind("paste",function(e) {
	         e.preventDefault();
	     });
	     
	     $('input[type=text], textarea').placeholder();
	     /*************************To fix the design ******************************/	     	     
	     fixDesignBS();
	    
	   }
	
	/****************** When the energy values are different ********************/
	if(check==false){
		var status = "+";
		var energyStyle = null;
		totalDiv = divCount+1;
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
		var divCountLoop = $('#pageWrap').children('.before_sketch_task').length;
		var totalEnergy = 0;
		for(var i=1; i<=divCountLoop;i++){
			var style = document.getElementById("divImg_" +i).getAttribute("style");
			if((style.search('block') != -1)){
				totalEnergy = totalEnergy + parseInt(document.getElementById("energyId_" +i).value);	
			}
		}
		if(energyValToAdd == 333 ){
			alertify.alert("<img src='../img/alert-icon.png'><br /><p>Task will not be added.</p>");
			return false;
		} else if(totalEnergy<105){
			if(next_page_nav == 0){
				count = ClientSession.get("total_count");
			    var divId = "divImg_" + count;
				var divIdx = "divImg_" + count+"_x";		
				var energyId = "energyId_" + count;
				var areaId = "areaId_" + count;		
				/**
				 * Added for lock and unlock feature
				 */			
				var divLockId = "divImg_" + count+"_lock";
				/******* END *******/
				if(totalDiv==20){
					if(navigator.userAgent.match(/Android/i)){
						var $element = $('<div id="' + divId +'" class="col-md-4 single_task_item before_sketch_task draggableResizable css_5" style="display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top "><div class="col-md-3 col-xs-3 task task_smaller" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 energy energy_smaller" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/ Energy: </div><div class="col-md-2 energy_field"><input type="text" class="form-control-sketch form-control-sketch_smaller input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="col-md-12 add_task_area" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divLockId+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');				
					}else if(navigator.userAgent.match(/iPhone/i)){
						var $element = $('<div id="' + divId +'" class="col-md-4 single_task_item before_sketch_task draggableResizable bs_upto_9" style="display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top "><div class="col-md-3 col-xs-3 task task_smaller" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 energy energy_smaller" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/ Energy: </div><div class="col-md-2 energy_field"><input type="text" class="form-control-sketch form-control-sketch_smaller input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="col-md-12 add_task_area" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divLockId+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');			
					}else if(navigator.userAgent.match(/iPad/i)){
						var $element = $('<div id="' + divId +'" class="col-md-4 single_task_item before_sketch_task draggableResizable bs_upto_9" style="display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top "><div class="col-md-3 col-xs-3 task task_smaller" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 energy energy_smaller" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/Energy:&nbsp; </div><div class="col-md-2 energy_field"><input type="text" class="form-control-sketch form-control-sketch_smaller input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="col-md-12 add_task_area" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divLockId+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');			
					}
					else {
						 var $element = $('<div id="' + divId +'" class="col-md-4 single_task_item before_sketch_task draggableResizable uptoDescBS_9" style="display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top "><div class="col-md-3 col-xs-3 task task_smaller">TASK</div><div class="col-md-6 col-xs-5 energy energy_smaller">Time/Energy:&nbsp; </div><div class="col-md-2 energy_field"><input type="text" class="form-control-sketch form-control-sketch_smaller input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)" /></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="col-md-12 add_task_area" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divLockId+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');
					}										
				} else {
					if(navigator.userAgent.match(/Android/i)){
						if (windowSize>550) {
						 var $element = $('<div id="' + divId +'" class="col-md-4 single_task_item before_sketch_task draggableResizable uptoBs_10" style="display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top "><div class="col-md-3 col-xs-3 task task_smaller" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 energy energy_small" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/ Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch form-control-sketch_small input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="col-md-12 add_task_area" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divLockId+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');					
						}else{
						 var $element = $('<div id="' + divId +'" class="col-md-4 single_task_item before_sketch_task draggableResizable bs_upto_10" style="display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top "><div class="col-md-3 col-xs-3 task" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 energy energy_small" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/ Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch form-control-sketch_small input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="col-md-12 add_task_area" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divLockId+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');						
						}
					}else if(navigator.userAgent.match(/iPhone/i)){
						 var $element = $('<div id="' + divId +'" class="col-md-4 single_task_item before_sketch_task draggableResizable bs_upto_10" style="display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top "><div class="col-md-3 col-xs-3 task" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 energy energy_small" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/ Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch form-control-sketch_small input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="col-md-12 add_task_area" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divLockId+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');
					}else if(navigator.userAgent.match(/iPad/i)){
						 var $element = $('<div id="' + divId +'" class="col-md-4 single_task_item before_sketch_task draggableResizable bs_upto_10" style="display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top "><div class="col-md-3 col-xs-3 task task_smaller" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 energy energy_small" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/Energy:&nbsp; </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch form-control-sketch_small input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="col-md-12 add_task_area" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divLockId+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');					
					} else {
						 var $element = $('<div id="' + divId +'" class="col-md-4 single_task_item before_sketch_task draggableResizable uptoDescBS_10" style="display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top "><div class="col-md-3 col-xs-3 task task_smaller">TASK</div><div class="col-md-6 col-xs-5 energy energy_small">Time/Energy:&nbsp; </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch form-control-sketch_small input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)" /></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="col-md-12 add_task_area" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divLockId+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');
					}										 
				}
			   						   
			} else {
				count = ClientSession.get("total_count");
				var divId = "divImg_" + count;
				var divIdx = "divImg_" + count+"_x";		
				var energyId = "energyId_" + count; 
				var areaId = "areaId_" + count;
				/**
				 * Added for lock and unlock feature
				 */			
				var divLockId = "divImg_" + count+"_lock";
				/******* END *******/
				if(totalDiv==20){
					if(navigator.userAgent.match(/Android/i)){
						var $element = $('<div id="' + divId +'" class="col-md-4 single_task_item before_sketch_task draggableResizable css_5" style="display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top"><div class="col-md-3 col-xs-3 task task_smaller" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 energy energy_smaller" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/ Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch form-control-sketch_smaller input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="col-md-12 add_task_area" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divLockId+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');				
					}else if(navigator.userAgent.match(/iPhone/i)){
						var $element = $('<div id="' + divId +'" class="col-md-4 single_task_item before_sketch_task draggableResizable bs_upto_9" style="display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top "><div class="col-md-3 col-xs-3 task task_smaller" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 energy energy_smaller" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/ Energy: </div><div class="col-md-2 energy_field"><input type="text" class="form-control-sketch form-control-sketch_smaller input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="col-md-12 add_task_area" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divLockId+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');
					}else if(navigator.userAgent.match(/iPad/i)){
						var $element = $('<div id="' + divId +'" class="col-md-4 single_task_item before_sketch_task draggableResizable bs_upto_9" style="display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top "><div class="col-md-3 col-xs-3 task task_smaller" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 energy energy_smaller" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/Energy:&nbsp; </div><div class="col-md-2 energy_field"><input type="text" class="form-control-sketch form-control-sketch_smaller input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="col-md-12 add_task_area" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divLockId+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');
					}else {
						var $element = $('<div id="' + divId +'" class="col-md-4 single_task_item before_sketch_task draggableResizable uptoDescBS_9" style="display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top"><div class="col-md-3 col-xs-3 task task_smaller">TASK</div><div class="col-md-6 col-xs-5 energy energy_smaller">Time/Energy:&nbsp; </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch form-control-sketch_smaller input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)" /></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="col-md-12 add_task_area" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divLockId+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');
					}										 					
				} else {
					if(navigator.userAgent.match(/Android/i)){
					 if (windowSize>550){
						var $element = $('<div id="' + divId +'" class="col-md-4 single_task_item before_sketch_task draggableResizable uptoBs_10" style="display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top"><div class="col-md-3 col-xs-3 task task_smaller" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 energy energy_small" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/ Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch form-control-sketch_small input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="col-md-12 add_task_area" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divLockId+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');
					 } else{
						var $element = $('<div id="' + divId +'" class="col-md-4 single_task_item before_sketch_task draggableResizable bs_upto_10" style="display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top"><div class="col-md-3 col-xs-3 task" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 energy energy_small" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/ Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch form-control-sketch_small input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="col-md-12 add_task_area" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divLockId+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');
					 }				       
					}else if(navigator.userAgent.match(/iPhone/i)){
						 var $element = $('<div id="' + divId +'" class="col-md-4 single_task_item before_sketch_task draggableResizable bs_upto_10" style="display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top "><div class="col-md-3 col-xs-3 task" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 energy energy_small" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/ Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch form-control-sketch_small input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="col-md-12 add_task_area" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divLockId+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');
					}else if(navigator.userAgent.match(/iPad/i)){
						 var $element = $('<div id="' + divId +'" class="col-md-4 single_task_item before_sketch_task draggableResizable bs_upto_10" style="display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top "><div class="col-md-3 col-xs-3 task task_smaller" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 energy energy_small" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/Energy:&nbsp; </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch form-control-sketch_small input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="col-md-12 add_task_area" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divLockId+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');
					} else {
						var $element = $('<div id="' + divId +'" class="col-md-4 single_task_item before_sketch_task draggableResizable uptoDescBS_10" style="display:block; width:'+energyStyle+'px; height:'+energyStyle+'px;"><div class="col-md-12 task_item_top"><div class="col-md-3 col-xs-3 task task_smaller">TASK</div><div class="col-md-6 col-xs-5 energy energy_small">Time/Energy:&nbsp; </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch form-control-sketch_small input-sm_custom" id="'+energyId+'" value="'+energyValue+'" name="MMM" maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)" /></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="col-md-12 add_task_area" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divLockId+'")><img src="../img/imgUnlock.png" alt="Lock" class="lock_unlock_img"/></div></div>');
					}										
				}
				}			
		  	count++;	  	
		  	ClientSession.set("total_count", count);	  	
		  	calculateEnergyValue(10, status, null, "ADD");		
		  	/************** To remove the focus to the new old task *************/
		  	removeFocus();
		    $("#pageWrap").append($element);			    
		    $element.draggable();	

			/************** Element will be drag with in "#pageWrap" element ****************/
			$('.before_sketch_task').draggable({
				cursor: 'move',       
				containment: '#pageWrap'    
			}); 
		    $element.resizable({		
				//ghost: 'true',
				handles: 'se',
				aspectRatio: 2 / 2,
			});			
	  /*  $element.resizable({
			//stop: resized
			stop:function(event,ui){
				var divId = $(this).attr('id');				
		    	var finalWidth = $(this).width();     	
		    	calculateEnergy(finalWidth, divId);		
		    	*//*************************To fix the design ******************************//*
		    	fixDesignBS();
			}
	    	});	*/  
		    allCanvasArray.push($element);	
		    addTaskPosition(energyStyle,divId);			//		Newly Added	
		    /***************** ISSUE# 207***************/
		    $("#" + divId).children(".ui-resizable-handle").css({'bottom' : '-8px'});
		    
		    /***************** To set the focus to the new Task ***************/
		    $("#" + divId).children().find(".add_task_area").addClass("add_task_area_focus");
		    $("#" + areaId).focus();
		    
		    /***************** To Disable the paste functionality ISSUE# 186***************/
		    $('#' +areaId).bind("paste",function(e) {
		         e.preventDefault();
		     });
		    
		    $('input[type=text], textarea').placeholder();
		} else {
			alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please resize your tasks first. Total task should be within 100</p>");
		}			  
	   	    	    
	}		   
    /**************** Store the Json Object**********************/
    generateJsonObj();        
    /**************** to get the resized width value**********************/    
    /*$('.ui-resizable-se').mouseup(function() {
    	var divId = $(this).parent().attr('id');
    	var finalWidth = $(this).parent().width();    	
    	calculateEnergy(finalWidth, divId);	
    	*//*************************To fix the design ******************************//*
    	fixDesignBS();
    });  */ 
    $(".single_task_item").resizable({
		stop:function(event,ui){			
			var divId = $(this).attr('id');				
	    	var finalWidth = $(this).width();     	
	    	calculateEnergy(finalWidth, divId);		
		}
    	});	
	}
return false;
});


/**
 * Function add to delete the Task div
 * @param id
 */
function deleteDiv(id) {
	removeFocus();
	var divCount = $('#pageWrap').children('div:visible').length;
	if(divCount>3){
		$("#" + id).children().find(".add_task_area").addClass("add_task_area_delete_focus");
		var unlockDivCount = $('#pageWrap').children('.single_task_item:visible').length;	
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
				    setOuterDivWhileDelete("decrease");
				    generateJsonObj();
			    } else {
			    	$("#" + id).children().find(".add_task_area").removeClass("add_task_area_delete_focus");
			    }
			});
		} else {
			alertify.alert("<img src='../img/alert-icon.png'><br /><p>Cannot delete the task since rest all the tasks are locked.</p>");
			$('#alertify-ok').click(function() {
				$("#" + id).children().find(".add_task_area").removeClass("add_task_area_delete_focus");
			});		
			
			}	
	} else {
		alertify.alert("<img src='../img/alert-icon.png'><br /><p>Minimum three tasks are mandatory.</p>");
	}		
}


/**
 * Function to store the json string  
 * while navigate to next page
 * @param null
 */
function goToNextPage() {
	var outerDivHeight = $('#pageWrap').height();
	sessionStorage.setItem("bsDivHeight", outerDivHeight);
	//callInitialJsonOnResize();
	var descArray = new Array();	
    var divCountLoop = $('#pageWrap').children('.before_sketch_task').length;
	var totalEnergy = 0;
	var taskDesc = null;
	var check = true;	
	/******************* check if the total energy is with in 100***********************/
	for(var i=1; i<=divCountLoop;i++){
		var style = document.getElementById("divImg_" +i).getAttribute("style");
		if((style.search('block') != -1)){
			totalEnergy = totalEnergy + parseInt(document.getElementById("energyId_" +i).value);				
		}		
	}	
	
	/******************* Set focus to the Task Area which does not contain value***********************/
	if(sessionStorage.getItem("src") != "Admin") {
		if(sessionStorage.getItem("isLogout")=="N"){
			removeFocus();
			for(var i=1; i<=divCountLoop;i++){
				var style = document.getElementById("divImg_" +i).getAttribute("style");
				if((style.search('block') != -1)){
					taskId = document.getElementById("divImg_"+i).id;
					taskDesc = document.getElementById("areaId_"+i).value.trim();	
					if(taskDesc == "") {
						$("#" + taskId).children().find(".add_task_area").addClass("add_task_area_focus");
					}
				}		
			}
		}
	}
	
	/******************* Null check for Task Description ***********************/
	if(sessionStorage.getItem("src") != "Admin") {
		if(sessionStorage.getItem("isLogout")=="N" && (!sessionStorage.getItem("bktp"))){
			for(var i=1; i<=divCountLoop;i++){
				var style = document.getElementById("divImg_" +i).getAttribute("style");
				if((style.search('block') != -1)){
					taskDesc = document.getElementById("areaId_"+i).value.trim();	
					if(taskDesc == "") {										
						check = false;
						alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please add task description.</p>");
						//break;
						return false;
					} else {										
						descArray.push(taskDesc);
					}
				}		
			}
			/******************* Same value not allow for Task Description ***********************/
			if (!sessionStorage.getItem("bktp")) {
				for (var i = 0; i < descArray.length; i++) {
			        for (var j = 0; j < descArray.length; j++) {
			            if (i != j) {
			                if (descArray[i].trim().toUpperCase() == descArray[j].trim().toUpperCase()) {
			                	check = false;
			                	alertify.alert("<img src='../img/alert-icon.png'><br /><p>Task description should be different.</p>");
			                	for(var k=1; k<=divCountLoop;k++){
			            			var style = document.getElementById("divImg_" +k).getAttribute("style");
			            			if((style.search('block') != -1)){
			            				taskId = document.getElementById("divImg_"+k).id;
			            				taskDesc = document.getElementById("areaId_"+k).value;	
			            				if(taskDesc.trim().toUpperCase() == descArray[i].trim().toUpperCase()) {
			            					$("#" + taskId).children().find(".add_task_area").addClass("add_task_area_focus");
			            				}
			            			}		
			            		}	
			                	return false;
			                }
			            }
			        }
			    }
			}			
		}
	}
	if(check == true) {
		    if(totalEnergy<105 && totalEnergy>97) {
		    	$(".loader_bg").fadeIn();
		        $(".loader").fadeIn();
		    	sessionStorage.setItem("pageSequence", 2);
		    	nextPageNavigation = 1;
		    	var cnt = 1;
		    	var totalJson = [];
		    	var idx = 0;
		    	var divValue = document.getElementById("divImg_" + cnt);
		    	
		    	var mainDivWidth = $("#pageWrap").width();
		    	if (!(navigator.userAgent.match(/iPhone|iPad|iPod/i)) && !(navigator.userAgent.match(/Android/i))) {
		    		if(mainDivWidth < 1000) {
		    			mainDivWidth = 1021;
		    		}
		    	 }
		    	var mainDivHeight = $("#pageWrap").height();	
	    	    
		    	var gcd=calc(mainDivWidth,mainDivHeight);
	    	  	var r1=mainDivWidth/gcd;
	    	    var r2=mainDivHeight/gcd;
	    	    var ratio=r1+":"+r2;
		    	     	   
		    	if(divValue == null){
		    		alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please add atleast one task to continue.</p>");	
		    		return false;
		    	}else{
		    		while(divValue != null) {
		    			var unitJson = {};					
		    			unitJson["divId"] = document.getElementById("divImg_" + cnt).id;
		    			unitJson["divIdx"] = document.getElementById("divImg_" + cnt +"_x").id;
		    			unitJson["energyId"] = document.getElementById("energyId_"+ cnt).id;
		    			unitJson["energyValue"] = document.getElementById("energyId_"+ cnt).value;
		    			unitJson["style"] = document.getElementById("divImg_" + cnt).getAttribute("style");
		    			unitJson["areaId"] = document.getElementById("areaId_"+ cnt).id;
		    			unitJson["areaValue"] = document.getElementById("areaId_"+ cnt).value.replace(/,/g, "`").replace(/#/g, "-_-").replace(/\?/g, ";_;");
		    			unitJson["positionLeft"] = document.getElementById("divImg_" + cnt).style.left;		    					    			
		    			unitJson["positionTop"] = document.getElementById("divImg_" + cnt).style.top;
		    			unitJson["positionWidth"] = document.getElementById("divImg_" + cnt).style.width;
		    			unitJson["positionHeight"] = document.getElementById("divImg_" + cnt).style.height;
		    			var positionLeftPixel = document.getElementById("divImg_" + cnt).style.left;				 
		    			unitJson["positionLeftPixel"] = positionLeftPixel;
		    			unitJson["positionLeftPercent"] = getPercentageValue(mainDivWidth, positionLeftPixel);
		    			var positionTopPixel = document.getElementById("divImg_" + cnt).style.top;	
		    			unitJson["positionTopPixel"] = positionTopPixel;
		    			unitJson["positionTopPercent"] = getPercentageValue(mainDivHeight, positionTopPixel);
		    			var positionWidthPixel = document.getElementById("divImg_" + cnt).style.width;	
		    			var positionWidth = parseInt(positionWidthPixel.replace("px", ''));
		    			unitJson["positionWidthPercent"] = getPercentageValue(mainDivWidth, positionWidth);
		    			unitJson["mainDivWidth"] = mainDivWidth;
		    			unitJson["mainDivHeight"] = mainDivHeight;
		    			unitJson["aspectRatio"] = ratio;		    		
		    			var position = $("#divImg_" + cnt).position();
		    			var positionLeftNewPercent = getPercentageValue(mainDivWidth, position.left);
		    			var positionTopNewPercent = getPercentageValue(mainDivWidth, position.top);
		    			unitJson["positionLeftNew"] = positionLeftNewPercent;		    					    			
		    			unitJson["positionTopNew"] = positionTopNewPercent;
		    			
		    			/***** Added for lock and unlock feature *****/
		    			unitJson["divLockId"] = document.getElementById("divImg_" + cnt +"_lock").id;
		    			unitJson["lockImage"] = document.getElementById("divImg_" + cnt+ "_lock").getElementsByTagName("img")[0].getAttribute("src");
		    			
		    			totalJson[idx++] = unitJson;
		    			divValue = document.getElementById("divImg_" + ++cnt);
		    		}
		    		console.log(totalJson)
		    		var hasBlock = JSON.stringify(totalJson);
		    		if(hasBlock.search("block") == -1){
		    			alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please add atleast one task to continue.</p>");
		    			return false;
		    		}		
		    	}
		    	ClientSession.set("total_json_add_task",JSON.stringify(totalJson));
		    	ClientSession.set("Next_page_navigation", nextPageNavigation);
		    	if(ClientSession.get("total_json_obj_sketch1")!=null){
		    		var after_sketch= 1;
		    	} else {
		    		var after_sketch = 0; 
		    	}
		    	ClientSession.set("after_sketch_navigation", after_sketch);
		    	sessionStorage.setItem("next_page", "BS");
		    	storeToDatabase(totalJson);
		    } else {
		    	if(sessionStorage.getItem("src") != "Admin") {
		    	alertify.alert("<img src='../img/alert-icon.png'><br /><p>Please resize your tasks. Total task should be within 100</p>");
		    	}
		    }
		}
}

/**
 * Function add to calculate the total time
 * @param null
 */
function getTotalTimeSpent(){
	if (document.getElementById("stwa")) {
		var timeCounterSplit = document.getElementById("stwa").value.split(":");
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
	var totalDivCount = $('#pageWrap').children('.before_sketch_task').length;		// total task block count
	var taskDesc = "";
	for(var index = 1; index <= totalDivCount; index++) {
		var elementStyle = document.getElementById("divImg_"+index).style.display;
		if ((elementStyle.search('block') != -1 )) { 				//BLOCK -- DISPLAYING
			document.getElementById("divImg_"+index+"_x").style.display = "none";
			document.getElementById("divImg_"+index+"_lock").style.display = "none";  // remove the lock unlock image in screenshot
			
			$("#divImg_" +index).children().find(".add_task_area").addClass("task_area_bright");
			if(document.getElementById("energyId_"+index).value >=5 && document.getElementById("energyId_"+index).value <=9){					
				$("#divImg_" +index).children().find(".energy_field").children().addClass("bs_small_remove_text");	
			} else {
				$("#divImg_" +index).children().find(".energy_field").children().addClass("bs_remove_input_text");
			}						
			if(document.getElementById("energyId_"+index).value >=5 && document.getElementById("energyId_"+index).value <=9){										
				taskDesc = document.getElementById("areaId_"+index).value;
				if(taskDesc.length <= 100) {
					$("#divImg_" +index).children().find(".add_task_area").addClass("smaller_task_font_bs_5_1");	
				} else if (taskDesc.length > 100 && taskDesc.length <= 150) {
					$("#divImg_" +index).children().find(".add_task_area").addClass("smaller_task_font_bs_5_2");	
				} else {
					$("#divImg_" +index).children().find(".add_task_area").addClass("smaller_task_font_bs_5_3");	
				}
			} else if(document.getElementById("energyId_"+index).value >9 && document.getElementById("energyId_"+index).value <=14){				
				taskDesc = document.getElementById("areaId_"+index).value;
				if(taskDesc.length <= 100) {
					$("#divImg_" +index).children().find(".add_task_area").addClass("smaller_task_font_bs_10_1");	
				} else if (taskDesc.length > 100 && taskDesc.length <= 150) {
					$("#divImg_" +index).children().find(".add_task_area").addClass("smaller_task_font_bs_10_2");	
				} else {
					$("#divImg_" +index).children().find(".add_task_area").addClass("smaller_task_font_bs_10_3");	
				}
			} else if(document.getElementById("energyId_"+index).value >14 && document.getElementById("energyId_"+index).value <20){
				taskDesc = document.getElementById("areaId_"+index).value;				
				if(taskDesc.length <= 150) {
					$("#divImg_" +index).children().find(".add_task_area").addClass("smaller_task_font_bs_15_1");	
				} else {
					$("#divImg_" +index).children().find(".add_task_area").addClass("smaller_task_font_bs_15_2");	
				}
			} else if (document.getElementById("energyId_"+index).value >=20 && document.getElementById("energyId_"+index).value <30) {				
				taskDesc = document.getElementById("areaId_"+index).value;
				if(taskDesc.length <= 150) {
					$("#divImg_" +index).children().find(".add_task_area").addClass("smaller_task_font_bs_20_1");	
				} else {
					$("#divImg_" +index).children().find(".add_task_area").addClass("smaller_task_font_bs_20_2");	
				}
			} else if (document.getElementById("energyId_"+index).value >=30 && document.getElementById("energyId_"+index).value <=40) {				
				taskDesc = document.getElementById("areaId_"+index).value;
				if(taskDesc.length <= 150) {
					$("#divImg_" +index).children().find(".add_task_area").addClass("smaller_task_font_bs_30_1");	
				} else {
					$("#divImg_" +index).children().find(".add_task_area").addClass("smaller_task_font_bs_30_2");	
				}
			}
			else if (document.getElementById("energyId_"+index).value >40 && document.getElementById("energyId_"+index).value <=60) {
				$("#divImg_" +index).children().find(".add_task_area").addClass("large_task_font_bs");	
			} else {
				$("#divImg_" +index).children().find(".add_task_area").addClass("larger_task_font_bs");	
			}
			$("#divImg_" + index).removeClass("lock_single_task_item");
			$("#divImg_" + index).addClass("single_task_item");	
			$("#energyId_" + index).removeClass("energy_lock");
		}
	}
}

/**
 * Function add to enable the cross image
 * while capturing the image
 * @param null
 */
function openAllCrossForBlockDiv() {
	var totalDivCount = $('#pageWrap').children('.before_sketch_task').length;		// total task block count
	for(var index = 1; index <= totalDivCount; index++) {
		var elementStyle = document.getElementById("divImg_"+index).style.display;
		if ((elementStyle.search('block') != -1 )) { 				//BLOCK -- DISPLAYING
			document.getElementById("divImg_"+index+"_x").style.display = "block";
			document.getElementById("divImg_"+index+"_lock").style.display = "block";
			$(".before_sketch_task").children().find(".energy_field").children().removeClass("bs_remove_input_text");
			$(".before_sketch_task").children().find(".energy_field").children().removeClass("bs_small_remove_text");
			$("#divImg_" +index).children().find(".add_task_area").removeClass("smaller_task_font_bs_5_1");
			$("#divImg_" +index).children().find(".add_task_area").removeClass("smaller_task_font_bs_5_2");
			$("#divImg_" +index).children().find(".add_task_area").removeClass("smaller_task_font_bs_5_3");
			$("#divImg_" +index).children().find(".add_task_area").removeClass("smaller_task_font_bs_10_1");
			$("#divImg_" +index).children().find(".add_task_area").removeClass("smaller_task_font_bs_10_2");
			$("#divImg_" +index).children().find(".add_task_area").removeClass("smaller_task_font_bs_10_3");
			$("#divImg_" +index).children().find(".add_task_area").removeClass("smaller_task_font_bs_15_1");
			$("#divImg_" +index).children().find(".add_task_area").removeClass("smaller_task_font_bs_15_2");
			$("#divImg_" +index).children().find(".add_task_area").removeClass("smaller_task_font_bs_15_3");
			
			$("#divImg_" +index).children().find(".add_task_area").removeClass("small_task_font_bs_20");
			$("#divImg_" +index).children().find(".add_task_area").removeClass("small_task_font_bs_30");
			
			$("#divImg_" +index).children().find(".add_task_area").removeClass("large_task_font_bs");
			$("#divImg_" +index).children().find(".add_task_area").removeClass("larger_task_font_bs");
			$("#divImg_" +index).children().find(".add_task_area").removeClass("task_area_bright");
		}
	}
}
/**
 * Function sends the json string to the server for storing
 * @param jsonString
 */
function storeToDatabase(jsonString){
	setOuterDivWhileDelete("decrease");
	var totalTime = getTotalTimeSpent();
	sessionStorage.setItem("bsView","H");
	var element = document.getElementById("pageWrap");
	closeAllCrossForBlockDiv();
	html2canvas(element, {
        onrendered: function(canvas) {
            snapShotURL = canvas.toDataURL("image/png"); //get's image string
            sessionStorage.setItem("snapShotURLS", snapShotURL);
            openAllCrossForBlockDiv();
            generateJsonObj(); // test
            
            var isEdit = "N";
            if (sessionStorage.getItem("bsEdit")) {
            	isEdit = "Y";
            }
            var beforeSketchModel = Spine.Model.sub();
        	beforeSketchModel.configure("/user/storeBeforeSketch", "createdBy", "jsonString", "jobReferenceString", "snapShot", "jobTitle", "initialJsonVal", "totalTimes" , "profileId", "isEdits", "linkClicked", "rowId", "jctUserId");
        	beforeSketchModel.extend( Spine.Model.Ajax );
        	var modelPopulator = new beforeSketchModel({  
        		createdBy: sessionStorage.getItem("jctEmail"), 
        		jsonString: jsonString,
        		jobReferenceString: sessionStorage.getItem("jobReferenceNo"),
        		snapShot: sessionStorage.getItem("snapShotURLS"),
        		jobTitle: sessionStorage.getItem("jobTitle"),
        		initialJsonVal: JSON.parse(sessionStorage.getItem("div_intial_value")),
        		//totalTimes: totalTime,
        		totalTimes: sessionStorage.getItem("pageTime"),
        		profileId: sessionStorage.getItem("profileId"),
        		isEdits : isEdit,
        		linkClicked: sessionStorage.getItem("linkClicked"),
        		rowId: sessionStorage.getItem("rowIdentity"),
        		jctUserId: sessionStorage.getItem("jctUserId")
        	});
        	modelPopulator.save(); //POST
        	beforeSketchModel.bind("ajaxError", function(record, xhr, settings, error) {
        		$(".loader_bg").fadeOut();
        	    $(".loader").hide();
        	});
        	
        	beforeSketchModel.bind("ajaxSuccess", function(record, xhr, settings, error) {
        		var jsonStr = JSON.stringify(xhr);
    			var obj = jQuery.parseJSON(jsonStr);
    			if(sessionStorage.getItem("isLogout") == "N") { //Normal next
    				sessionStorage.setItem("questions", obj.list);
    				if (sessionStorage.getItem("bktp")) {
    					if (sessionStorage.getItem("bktp")) {
    						sessionStorage.removeItem("bktp");
    					}
    					sessionStorage.setItem("pageSequence", 1);
    					window.location = "landing-page.jsp";
    				} else if (isEdit == "N") {
    					removePresentSessionItems();
    					window.location = "Questionaire.jsp";
    				} else {
    					sessionStorage.setItem("total_json_obj", obj.pageOneJson);
    					sessionStorage.setItem("div_intial_value", obj.divInitialValue);
    					sessionStorage.setItem("total_json_after_sketch_final", obj.afterSketch2TotalJsonObj);
    					sessionStorage.setItem("total_count", obj.asTotalCount);
    					sessionStorage.setItem("divHeight", obj.divHeight);
    					sessionStorage.setItem("divWidth", obj.divWidth);
    					sessionStorage.setItem("myPage", "AS");
    					window.location = "afterSketchEdit.jsp";
    				}
    			} else { //It is logout
    				//Create a model
    				var myLastActivePage = "/user/view/beforeSketch.jsp";
    				if (sessionStorage.getItem("bsEdit")) {
    					myLastActivePage = "/user/view/beforeSketchEdit.jsp";
    				}
    				var LogoutVO = Spine.Model.sub();
    				LogoutVO.configure("/user/auth/logout", "jobReferenceString", "rowId", "lastActivePage");
    				LogoutVO.extend( Spine.Model.Ajax );
    				//Populate the model with data to transder
    				var logoutModel = new LogoutVO({  
    					jobReferenceString: sessionStorage.getItem("jobReferenceNo"),
    					rowId: sessionStorage.getItem("rowIdentity"),
    					//lastActivePage: "/user/view/beforeSketch.jsp"
    					lastActivePage: myLastActivePage
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
 * Check if the json object and 
 * draw the task according to the json object
 */
var nextPageNavigation = 0; 
if(ClientSession.get("Next_page_navigation")!= null){
	var next_page_nav = ClientSession.get("Next_page_navigation");
} else {
	var next_page_nav = 0;
} 
if(ClientSession.get("total_json_add_task")!=null){
	var insDisplay = sessionStorage.getItem("bsView");
	if(null == insDisplay || insDisplay == "" || insDisplay == "D"){
		document.getElementById('panel').style.display = "block";
		document.getElementById('drp').setAttribute("class", "btn-slide active");
	} else {
		document.getElementById('panel').style.display = "none";
		document.getElementById('drp').setAttribute("class", "btn-slide");
	}
	if(sessionStorage.getItem("isCompleted") == 1){
		if (sessionStorage.getItem("bsEdit")) {
			$("#addtask").addClass("disable_button");
			var jsonObj = JSON.parse(ClientSession.get("total_json_add_task"));
			var count1 = jsonObj.length;
			for (var i = 0; i<count1;i++){	
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
				
				var aspectRatio = jsonObj[i]["aspectRatio"];
				var deviceWidth = $("#pageWrap").width();
				if (!(navigator.userAgent.match(/iPhone|iPad|iPod/i)) && !(navigator.userAgent.match(/Android/i))) {
                    if(deviceWidth < 1000) {
                    	deviceWidth = 1021;
                    }
				}
				var deviceHeight = getActalHeight(aspectRatio, deviceWidth);
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
				
				/****  Added for lock and unlock feature *****/
				var divLockId = jsonObj[i]["divLockId"];
				var lockImage = jsonObj[i]["lockImage"];
				/******* END *******/
				
				/*
				 * Make the text area non-editable
				 * Dated : 17.07.2014
				 */
				if(navigator.userAgent.match(/Android/i) || navigator.userAgent.match(/iPhone/i)){
					var $element2 = $('<div id="' + divIda +'" class="col-md-4 single_task_item before_sketch_task draggableResizable" style="position: absolute; top:'+positionTopNewVal+'px; left:'+positionLeftNewVal+'px; width:'+positionWidthVal+'px; display:'+divStyle+'; height:'+positionWidthVal+'px;"><div class="col-md-12 task_item_top"><div class="col-md-3 col-xs-3 task" onclick=enableDraggableResizable("'+divIda+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 energy" onclick=enableDraggableResizable("'+divIda+'") onblur="disableDraggableResizable()">Time/ Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom " id="'+energyIda+'" value="'+energyValue+'" name="MMM"  maxlength="2" readonly="true" onblur="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="well col-md-12 add_task_area" style="text-align: center;" id="'+areaIda+'" type="text" name="areaName" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)" onBlur="disableDraggableResizable()">'+areaValue+'</textarea></div><div id="' + divIdx +'" class="cross"></div><div id="' + divLockId +'" class="lock_unlock none_display"><img src="'+lockImage+'" alt="Lock" /></div></div>');
				}  else if(navigator.userAgent.match(/iPad/i) ){
					var $element2 = $('<div id="' + divIda +'" class="col-md-4 single_task_item before_sketch_task draggableResizable" style="position: absolute; top:'+positionTopNewVal+'px; left:'+positionLeftNewVal+'px; width:'+positionWidthVal+'px; display:'+divStyle+'; height:'+positionWidthVal+'px;"><div class="col-md-12 task_item_top"><div class="col-md-3 col-xs-3 task" onclick=enableDraggableResizable("'+divIda+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 energy" onclick=enableDraggableResizable("'+divIda+'") onblur="disableDraggableResizable()">Time/Energy:&nbsp; </div><div class="col-md-2 col-xs-2 energy_field "><input type="text" class="form-control-sketch input-sm_custom" id="'+energyIda+'" value="'+energyValue+'" name="MMM"  maxlength="2" readonly="true" onblur="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="well col-md-12 add_task_area" style="text-align: center;" id="'+areaIda+'" type="text" name="areaName" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)" onBlur="disableDraggableResizable()">'+areaValue+'</textarea></div><div id="' + divIdx +'" class="cross"></div><div id="' + divLockId +'" class="lock_unlock none_display"><img src="'+lockImage+'" alt="Lock" /></div></div>');						
				} else {
					var $element2 = $('<div id="' + divIda +'" class="col-md-4 single_task_item before_sketch_task draggableResizable" style="position: absolute; top:'+positionTopNewVal+'px; left:'+positionLeftNewVal+'px; width:'+positionWidthVal+'px; display:'+divStyle+'; height:'+positionWidthVal+'px;"><div class="col-md-12 task_item_top"><div class="col-md-3 col-xs-3 task">TASK</div><div class="col-md-6 col-xs-5 energy">Time/Energy:&nbsp; </div><div class="col-md-2 col-xs-2 energy_field "><input type="text" class="form-control-sketch input-sm_custom" id="'+energyIda+'" value="'+energyValue+'" name="MMM"  maxlength="2" readonly="true" /></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="well col-md-12 add_task_area" style="text-align: center;" id="'+areaIda+'" type="text" name="areaName" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)">'+areaValue+'</textarea></div><div id="' + divIdx +'" class="cross"></div><div id="' + divLockId +'" class="lock_unlock none_display"><img src="'+lockImage+'" alt="Lock" /></div></div>');
				}								
				$("#pageWrap").append($element2);
				if(energyValue >=5 && energyValue <=9){		
					$("#" +divIda).children().find(".energy_field").children().addClass("bs_small_remove_text");					
				} else {
					$("#" +divIda).children().find(".energy_field").children().addClass("bs_remove_input_text");
				}
				/***************** To Disable the paste functionality ISSUE# 186***************/
				$('#' +areaIda).bind("paste",function(e) {
			         e.preventDefault();
			     });     
			}
		} else {
			$("#addtask").addClass("disable_button");
			var jsonObj = JSON.parse(ClientSession.get("total_json_add_task"));
			var count1 = jsonObj.length;
			for (var i = 0; i<count1;i++){	
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
				
				var aspectRatio = jsonObj[i]["aspectRatio"];
				var deviceWidth = $("#pageWrap").width();
				if (!(navigator.userAgent.match(/iPhone|iPad|iPod/i)) && !(navigator.userAgent.match(/Android/i))) {
                    if(deviceWidth < 1000) {
                    	deviceWidth = 1021;
                    }
				}
				var deviceHeight = getActalHeight(aspectRatio, deviceWidth);
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
				
				/****  Added for lock and unlock feature *****/
				var divLockId = jsonObj[i]["divLockId"];
				var lockImage = jsonObj[i]["lockImage"];
				/******* END *******/
				/**
				 * Make the text area non-editable
				 * Dated : 17.07.2014
				 */
				if(navigator.userAgent.match(/Android/i) || navigator.userAgent.match(/iPhone/i)){
					var $element2 = $('<div id="' + divIda +'" class="col-md-4 single_task_item before_sketch_task" style="position: absolute; top:'+positionTopNewVal+'px; left:'+positionLeftNewVal+'px; width:'+positionWidthVal+'px; display:'+divStyle+'; height:'+positionWidthVal+'px;"><div class="col-md-12 task_item_top"><div class="col-md-3 col-xs-3 task" onclick=enableDraggableResizable("'+divIda+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 energy" onclick=enableDraggableResizable("'+divIda+'") onblur="disableDraggableResizable()">Time/ Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyIda+'" value="'+energyValue+'" name="MMM"  maxlength="2" readonly="true" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="well col-md-12 add_task_area" style="text-align: center;" id="'+areaIda+'" type="text" name="areaName" placeholder="Add Task Label Here" maxlength="200" readonly="true" onkeypress="return disableKey(event)" onBlur="disableDraggableResizable()">'+areaValue+'</textarea></div><div id="' + divIdx +'" class="cross"></div><div id="' + divLockId +'" class="lock_unlock none_display"><img src="'+lockImage+'" alt="Lock" /></div></div>');			
				}  else if(navigator.userAgent.match(/iPad/i) ){
					var $element2 = $('<div id="' + divIda +'" class="col-md-4 single_task_item before_sketch_task" style="position: absolute; top:'+positionTopNewVal+'px; left:'+positionLeftNewVal+'px; width:'+positionWidthVal+'px; display:'+divStyle+'; height:'+positionWidthVal+'px;"><div class="col-md-12 task_item_top"><div class="col-md-3 col-xs-3 task" onclick=enableDraggableResizable("'+divIda+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 energy" onclick=enableDraggableResizable("'+divIda+'") onblur="disableDraggableResizable()">Time/Energy:&nbsp; </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyIda+'" value="'+energyValue+'" name="MMM"  maxlength="2" readonly="true" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="well col-md-12 add_task_area" style="text-align: center;" id="'+areaIda+'" type="text" name="areaName" placeholder="Add Task Label Here" maxlength="200" readonly="true" onkeypress="return disableKey(event)" onBlur="disableDraggableResizable()">'+areaValue+'</textarea></div><div id="' + divIdx +'" class="cross"></div><div id="' + divLockId +'" class="lock_unlock none_display"><img src="'+lockImage+'" alt="Lock" /></div></div>');					
				}	else {
					var $element2 = $('<div id="' + divIda +'" class="col-md-4 single_task_item before_sketch_task" style="position: absolute; top:'+positionTopNewVal+'px; left:'+positionLeftNewVal+'px; width:'+positionWidthVal+'px; display:'+divStyle+'; height:'+positionWidthVal+'px;"><div class="col-md-12 task_item_top"><div class="col-md-3 col-xs-3 task">TASK</div><div class="col-md-6 col-xs-5 energy">Time/Energy:&nbsp; </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyIda+'" value="'+energyValue+'" name="MMM"  maxlength="2" readonly="true" /></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="well col-md-12 add_task_area" style="text-align: center;" id="'+areaIda+'" type="text" name="areaName" placeholder="Add Task Label Here" maxlength="200" readonly="true" onkeypress="return disableKey(event)">'+areaValue+'</textarea></div><div id="' + divIdx +'" class="cross"></div><div id="' + divLockId +'" class="lock_unlock none_display"><img src="'+lockImage+'" alt="Lock" /></div></div>');
				}		
				$("#pageWrap").append($element2);
				if(energyValue >=5 && energyValue <=9){		
					$("#" +divIda).children().find(".energy_field").children().addClass("bs_small_remove_text");					
				} else {
					$("#" +divIda).children().find(".energy_field").children().addClass("bs_remove_input_text");
				}
				/***************** To Disable the paste functionality ISSUE# 186***************/
				$('#' +areaIda).bind("paste",function(e) {
			         e.preventDefault();
			     });     
			}
		}
		//$(".ui-resizable-handle ui-resizable-se ui-icon ui-icon-gripsmall-diagonal-se").css({display: "none"});		
							
	} else {
		var jsonObj = JSON.parse(ClientSession.get("total_json_add_task"));
		var count1 = jsonObj.length;
		for (var i = 0; i<count1;i++){	
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
			
			var aspectRatio = jsonObj[i]["aspectRatio"];
			var deviceWidth = $("#pageWrap").width();
			if (!(navigator.userAgent.match(/iPhone|iPad|iPod/i)) && !(navigator.userAgent.match(/Android/i))) {
                if(deviceWidth < 1000) {
                	deviceWidth = 1021;
                }
			}
			var deviceHeight = getActalHeight(aspectRatio, deviceWidth);
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
			
			/****  Added for lock and unlock feature *****/
			var divLockId = jsonObj[i]["divLockId"];
			var lockImage = jsonObj[i]["lockImage"];
			/******* END *******/
			
			//var $element2 = $('<div id="' + divIda +'" class="col-md-4 single_task_item draggableResizable" style="position: relative; top:'+positionTopa+'; left:'+positionLefta+'; width:'+positionWidth+' ;display:'+divStyle+'; height:'+positionHeight+';"><div class="col-md-12 task_item_top"><div class="col-md-3 col-xs-3 task">TASK</div><div class="col-md-6 col-xs-5 energy">Time/Energy:&nbsp; </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyIda+'" value="'+energyValue+'" name="MMM"  maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)" /></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="well col-md-12 add_task_area" style="text-align: center;" id="'+areaIda+'" type="text" name="areaName" placeholder="Add Task Label Here" maxlength="100" onkeypress="return disableKey(event)">'+areaValue+'</textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divIda+'")><img src="../img/cross-black.png" alt="Delete" /></div></div>');
			if(navigator.userAgent.match(/Android/i) || navigator.userAgent.match(/iPhone/i)){
				var $element2 = $('<div id="' + divIda +'" class="col-md-4 single_task_item before_sketch_task draggableResizable" style="position: absolute; top:'+positionTopNewVal+'px; left:'+positionLeftNewVal+'px; width:'+positionWidthVal+'px ;display:'+divStyle+'; height:'+positionWidthVal+'px;"><div class="col-md-12 task_item_top"><div class="col-md-3 col-xs-3 task" onclick=enableDraggableResizable("'+divIda+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 energy" onclick=enableDraggableResizable("'+divIda+'") onblur="disableDraggableResizable()">Time/ Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyIda+'" value="'+energyValue+'" name="MMM"  maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="well col-md-12 add_task_area" style="text-align: center;" id="'+areaIda+'" type="text" name="areaName" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)" onBlur="disableDraggableResizable()">'+areaValue+'</textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divIda+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divIda+'","'+divLockId+'")><img src="'+lockImage+'" alt="Lock" class="lock_unlock_img"/></div></div>');
			}  else if(navigator.userAgent.match(/iPad/i) ){
				var $element2 = $('<div id="' + divIda +'" class="col-md-4 single_task_item before_sketch_task draggableResizable" style="position: absolute; top:'+positionTopNewVal+'px; left:'+positionLeftNewVal+'px; width:'+positionWidthVal+'px ;display:'+divStyle+'; height:'+positionWidthVal+'px;"><div class="col-md-12 task_item_top"><div class="col-md-3 col-xs-3 task" onclick=enableDraggableResizable("'+divIda+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 energy" onclick=enableDraggableResizable("'+divIda+'") onblur="disableDraggableResizable()">Time/Energy:&nbsp; </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyIda+'" value="'+energyValue+'" name="MMM"  maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="well col-md-12 add_task_area" style="text-align: center;" id="'+areaIda+'" type="text" name="areaName" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)" onBlur="disableDraggableResizable()">'+areaValue+'</textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divIda+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divIda+'","'+divLockId+'")><img src="'+lockImage+'" alt="Lock" class="lock_unlock_img"/></div></div>');
			} else {
				var $element2 = $('<div id="' + divIda +'" class="col-md-4 single_task_item before_sketch_task draggableResizable" style="position: absolute; top:'+positionTopNewVal+'px; left:'+positionLeftNewVal+'px; width:'+positionWidthVal+'px ;display:'+divStyle+'; height:'+positionWidthVal+'px;"><div class="col-md-12 task_item_top"><div class="col-md-3 col-xs-3 task">TASK</div><div class="col-md-6 col-xs-5 energy">Time/Energy:&nbsp; </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyIda+'" value="'+energyValue+'" name="MMM"  maxlength="2" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)" /></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="well col-md-12 add_task_area" style="text-align: center;" id="'+areaIda+'" type="text" name="areaName" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)">'+areaValue+'</textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divIda+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divIda+'","'+divLockId+'")><img src="'+lockImage+'" alt="Lock" class="lock_unlock_img"/></div></div>');
			}
			$("#pageWrap").append($element2);
			
			/***************** To Disable the paste functionality ISSUE# 186***************/
			$('#' +areaIda).bind("paste",function(e) {
		         e.preventDefault();
		     });	
			
			 $element2.draggable();
			 //	*******************     element will be allowed to dragable within #pageWrap
			$('.before_sketch_task').draggable({
					cursor: 'move',       
					containment: '#pageWrap'
				});  
												 
			 $element2.resizable({	
					//ghost: 'true',
					handles: 'se',
					aspectRatio: 2 / 2,
			});
			setClassAgainstLockUnlockElememt(divIda, divLockId, lockImage);	
		}
	}	
	/*************************To fix the design ******************************/
	fixDesignBS();
	setOuterDivWhileDelete("increasePageWrap");
}


/*************** Page loaded with three tasks **************/
if(ClientSession.get("total_json_add_task")==null){
	var insDisplay = sessionStorage.getItem("bsView");
	if(null == insDisplay || insDisplay == "" || insDisplay == "D"){
		document.getElementById('panel').style.display = "block";
		document.getElementById('drp').setAttribute("class", "btn-slide active");
	} else {
		document.getElementById('panel').style.display = "none";
		document.getElementById('drp').setAttribute("class", "btn-slide");
	}
	var count_1 = 1;	
	var mainDivWidth = $("#pageWrap").width();
	if (!(navigator.userAgent.match(/iPhone|iPad|iPod/i)) && !(navigator.userAgent.match(/Android/i))) {
		if(mainDivWidth < 1000) {
			mainDivWidth = 1021;
		}
	 }
	while(count_1<=3){	
		var allCanvasArray = new Array();							
		var divId = "divImg_" + count_1;
		var divIdx = "divImg_" + count_1+"_x";		
		var energyId = "energyId_" + count_1;
		var areaId = "areaId_" + count_1;		
		/**
		 * Added for lock and unlock feature
		 */		
		var divLockId = "divImg_" + count_1+"_lock";
		var lockImage = "../img/imgUnlock.png";
		/******* END *******/
		
		var taskWidth = getWidthDifference(226, mainDivWidth);
		if(navigator.userAgent.match(/Android/i) || navigator.userAgent.match(/iPhone/i)){
			var $element = $('<div id="' + divId +'" class="col-md-4 single_task_item before_sketch_task draggableResizable" style="display:block;  width:'+taskWidth+'px; height:'+taskWidth+'px;"><div class="col-md-12 task_item_top"><div class="col-md-3 col-xs-3 task" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 energy" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/ Energy: </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyId+'" name="MMM" maxlength="2" value="33" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="col-md-12 add_task_area" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divLockId+'")><img src="'+lockImage+'" alt="Lock" class="lock_unlock_img"/></div></div>');
		} else if(navigator.userAgent.match(/iPad/i)){
			var $element = $('<div id="' + divId +'" class="col-md-4 single_task_item before_sketch_task draggableResizable" style="display:block;  width:'+taskWidth+'px; height:'+taskWidth+'px;"><div class="col-md-12 task_item_top"><div class="col-md-3 col-xs-3 task" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">TASK</div><div class="col-md-6 col-xs-5 energy" onclick=enableDraggableResizable("'+divId+'") onblur="disableDraggableResizable()">Time/Energy:&nbsp; </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyId+'" name="MMM" maxlength="2" value="33" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)" onBlur="disableDraggableResizable()"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="col-md-12 add_task_area" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)" onBlur="disableDraggableResizable()"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divLockId+'")><img src="'+lockImage+'" alt="Lock" class="lock_unlock_img"/></div></div>');
		} else {
			var $element = $('<div id="' + divId +'" class="col-md-4 single_task_item before_sketch_task draggableResizable" style="display:block;  width:'+taskWidth+'px; height:'+taskWidth+'px;"><div class="col-md-12 task_item_top"><div class="col-md-3 col-xs-3 task">TASK</div><div class="col-md-6 col-xs-5 energy">Time/Energy:&nbsp; </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyId+'" name="MMM" maxlength="2" value="33" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="col-md-12 add_task_area" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="200" onkeypress="return disableKey(event)"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div><div id="' + divLockId +'" class="lock_unlock" onclick=lockUnlockDiv("'+divId+'","'+divLockId+'")><img src="'+lockImage+'" alt="Lock" class="lock_unlock_img"/></div></div>');
		}	    	
	    //var $element = $('<div id="' + divId +'" class="col-md-4 single_task_item draggableResizable" style="display:block;  width:226px; height:226px;"><div class="col-md-12 task_item_top"><div class="col-md-3 col-xs-3 task">TASK</div><div class="col-md-6 col-xs-5 energy">Time/Energy:&nbsp; </div><div class="col-md-2 col-xs-2 energy_field"><input type="text" class="form-control-sketch input-sm_custom" id="'+energyId+'" name="MMM" maxlength="2" value="33" onkeypress="return exceptNumberOnly(event, this)" onblur="return getEnergyValue(this)"/></div><div class="col-md-1 col-xs-1 energy_percent">%</div><div class="clearfix"></div></div><div class="col-md-12 task_item_botm"><textarea class="col-md-12 add_task_area" style="text-align: center;" id="'+areaId+'" type="text" name="areaName" value="value" placeholder="Add Task Label Here" maxlength="100" onkeypress="return disableKey(event)"></textarea></div><div id="' + divIdx +'" class="cross" onclick=deleteDiv("'+divId+'")><img src="../img/cross-black.png" alt="Delete" /></div></div>');	
	    count_1++;	  	
	 	ClientSession.set("total_count", count_1);
	    $("#pageWrap").append($element); 
	    allCanvasArray.push($element);  
	    /***************** To Disable the paste functionality ISSUE# 186***************/
	    $('#' +areaId).bind("paste",function(e) {
	         e.preventDefault();
	     });
	    
	    $element.draggable();
		 $('.before_sketch_task').draggable({
			cursor: 'move',       
			containment: '#pageWrap'    
		});  
					 		
	    $element.resizable({		
	    	//ghost: 'true',
			handles: 'se',
			aspectRatio: 2 / 2,
		});			
	    $element.resizable({
			//stop: resized
			stop:function(event,ui){			
				var divId = $(this).attr('id');
				var finalWidth = $(this).width();     	
				calculateEnergy(finalWidth, divId);		
				/*************************To fix the design ******************************/
				//fixDesignBS();
			}
	    });	  
	}	
	disableDraggableResizable();
	relativeToAbsolute();
	fixDesignBS();
	generateJsonObj();
	sessionStorage.setItem("restartExcersize",0);
}


/*************** Json Object of loaded three tasks **************/
function generateJsonObj() {
	var cnt = 1;
	var divJson = [];
	var idx = 0;
	var divValue = document.getElementById("divImg_" + cnt);
	var mainDivWidth = $('#pageWrap').width();
	if (!(navigator.userAgent.match(/iPhone|iPad|iPod/i)) && !(navigator.userAgent.match(/Android/i))) {
		if(mainDivWidth < 1000) {
			mainDivWidth = 1021;
		}
	 }
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
	//var totalEnergy = 0;
	var divCount = $('#pageWrap').children('.single_task_item:visible').length;
	var divCountLoop = $('#pageWrap').children('.single_task_item').length;
	totalDiv = divCount+1;
	var energyValue = Math.round(100/totalDiv);
	var totalEnergy =  energyValue*totalDiv;
	//var energyStyle = energyValue+5;
	for(var i=1; i<=divCountLoop;i++) {
	var divIdLock = "divImg_"+i+"_lock";
	var lockImage = document.getElementById(divIdLock).getElementsByTagName("img")[0].getAttribute("src");
	var style = document.getElementById("divImg_" +i).getAttribute("style");
		if((style.search('block') != -1)) {		
			/****** Added to fix first round defect date-14.02.2014*******/
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
	/*************************To fix the design ******************************/
	fixDesignBS();
}

/******** Added to fix second round defect **********/
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
	 var mainDivWidth = $('#pageWrap').width();
	 if (!(navigator.userAgent.match(/iPhone|iPad|iPod/i)) && !(navigator.userAgent.match(/Android/i))) {
			if(mainDivWidth < 1000) {
				mainDivWidth = 1021;
		}
	}
	/********** Called agsinst add task and when resize the boxes to increase the energy value *****************/
	if(status=='+'){
		var initialDiv = JSON.parse(ClientSession.get("div_intial_value"));	
		divLength = initialDiv.length;	
		var divCount = $('#pageWrap').children('.single_task_item:visible').length;
		var maxValue = 100 - ((divCount-1)*5);
		var divCountLoop = $('#pageWrap').children('.single_task_item').length;// total task block
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
						var widthValuePixel = "";
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
											//var widthValuePixel = Math.round(parseInt(divWidthPixel-((initialEnergyVal - energyValue)*3)));										
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
									//var widthValuePixel = Math.round(parseInt(divWidthPixel+(tempValue*3)));									
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
											//var widthValuePixel = Math.round(parseInt(divWidthPixel-((initialEnergyVal - energyValue)*3)));										
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
						//widthValuePixel = getTaskWidthValue(mainDivWidth, 140);
						widthValuePixel = getWidthDifference(140, mainDivWidth);									
					}
					//var widthValuePixel = 140;					
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
									//var widthValuePixel = Math.round(parseInt(divWidthPixel+(tempValue*3)));									
									elemPixel.style.width = widthValuePixel+"px";
									elemPixel.style.height = widthValuePixel+"px";
								}								
							}					
						else if (check == true){	
							if((style == null) || (style == "null") || (style.search('block') != -1)) {	
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
											//var widthValuePixel = Math.round(parseInt(divWidthPixel-((initialEnergyVal - energyValue)*3)));										
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
		var divCount = $('#pageWrap').children('.single_task_item:visible').length;
		var divCountLoop = $('#pageWrap').children('.single_task_item').length;// total task block
		totalDiv = divCount+1;	
		for (var l=0;l<tempValue;){
			for (var i = 0; i<divLength;i++) {
				var divId = initialDiv[i]["divId"];
				var energyValue = initialDiv[i]["energyValue"];				
				var positionWidthPercent = initialDiv[i]["positionWidthPercent"];
				var divWidthPixel = getActualValue(mainDivWidth, positionWidthPercent);				
				//var divWidthPixel = initialDiv[i]["divWidthPixel"];	
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
								//var widthValuePixel = Math.round(parseInt(divWidthPixel-(tempValue*3)));							
								elemPixel.style.width = widthValuePixel+"px";
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
									//var widthValuePixel = Math.round(parseInt(divWidthPixel-(tempValue*3)));							
									elemPixel.style.width = widthValuePixel+"px";
									elemPixel.style.height = widthValuePixel+"px";
								} else {
									calculateTotalEnergy();
								}							
							}														
						}						
					} else{							
						if((style == null) || (style == "null") || (style.search('block') != -1)) {
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
									//var widthValuePixel = Math.round(parseInt(divWidthPixel+((energyValue - initialEnergyVal)*3)));
									
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
	setOuterDivWhileDelete("increasePageWrap");
	fixDesignBS();
	adjustTaskWhileResize();
	generateJsonObj();
}

$(document).ready(function() {		
	if(sessionStorage.getItem("isCompleted") == 0){
		if (sessionStorage.getItem("bsEdit") || (null == sessionStorage.getItem("bsEdit"))) {			
			$('.before_sketch_task').resizable({
				stop:function(event,ui){
					var divId = $(this).attr('id');
					var finalWidth = $(this).width();    
					calculateEnergy(finalWidth, divId);		
					/*************************To fix the design ******************************/
					fixDesignBS();
				}
			});
		}
	}
});



/**
 * Function add to calculate width difference
 * @param widthFinal, divIdFinal
 */
function calculateEnergy(widthFinal, divIdFinal){	
	var status = null;
	var taskHeight = "";
	var taskWidth =  "";	
	var mainDivWidth = $('#pageWrap').width();
	if (!(navigator.userAgent.match(/iPhone|iPad|iPod/i)) && !(navigator.userAgent.match(/Android/i))) {
		if(mainDivWidth < 1000) {
			mainDivWidth = 1021;
		}
	 }
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
	//var differenceWidth = Math.round(parseInt(differenceValue/3));
	if(differenceWidth>0){
		status = "+";
	} else {
		differenceWidth = -(differenceWidth);
		status = "-";
	}		
	
	/*************** LOCK UNLOCK ******************/
	var unlockDivCount = $('#pageWrap').children('.single_task_item:visible').length;	//count the unlock div
	if(unlockDivCount == 1){
		calculateTotalEnergy();
	} else if(status=='+'){
		differenceWidth = getMaxEnergyValueToAdd(divIdFinal, differenceWidth, "RESIZE");
	}
	/**** END ****/
	
	taskHeight = $("#"+divIdFinal).height();
	taskWidth = $("#"+divIdFinal).width();
	
	if(unlockDivCount == 1){
		calculateTotalEnergy();
	} else if(differenceWidth == 0){
		calculateTotalEnergy();
	} else if(navigator.userAgent.match(/Android|iPhone/i)) {
		if(widthFinal > 225 || widthFinal < 70) {
			calculateTotalEnergy();
		} /*else if((taskHeight - taskWidth)>8 || (taskHeight - taskWidth)<-8){
			calculateTotalEnergy();
		}*/ else {
			calculateEnergyValue(differenceWidth,status,divIdFinal, "RESIZE");
		}
	} else if(navigator.userAgent.match(/iPad/i)) {
		if(widthFinal > 335 || widthFinal < 112) {
			calculateTotalEnergy();
		} /*else if((taskHeight - taskWidth)>15 || (taskHeight - taskWidth)<-15){
			calculateTotalEnergy();
		}*/ else {
			calculateEnergyValue(differenceWidth,status,divIdFinal, "RESIZE");
		}
	} else if(!(navigator.userAgent.match(/iPhone|iPad|iPod/i)) && !(navigator.userAgent.match(/Android/i)) ){
		if(widthFinal > 395 || widthFinal < 140) {
			calculateTotalEnergy();
		} /*else if((taskHeight - taskWidth)>15 || (taskHeight - taskWidth)<-15){
			calculateTotalEnergy();
		}*/ else {
			calculateEnergyValue(differenceWidth,status,divIdFinal, "RESIZE");
		}
	 }  /*else {
			calculateEnergyValue(differenceWidth,status,divIdFinal, "RESIZE");
		}
	} */else {
		calculateEnergyValue(differenceWidth,status,divIdFinal, "RESIZE");
	}
}


/**
 * Function add to set the initial position if the total task is more thn 100 
 * @param widthFinal, divIdFinal
 */
function calculateTotalEnergy() {
	var mainDivWidth = $('#pageWrap').width();
	if (!(navigator.userAgent.match(/iPhone|iPad|iPod/i)) && !(navigator.userAgent.match(/Android/i))) {
		if(mainDivWidth < 1000) {
			mainDivWidth = 1021;
		}
	 }
	var initialDiv = JSON.parse(ClientSession.get("div_intial_value"));		
	divLength = initialDiv.length;		
	for (var i = 0; i<divLength;i++) {
		var divId = initialDiv[i]["divId"];
		var energyId = initialDiv[i]["energyId"];
		var energyValue = initialDiv[i]["energyValue"];	
		var positionWidthPercent = initialDiv[i]["positionWidthPercent"];
		var divWidthPixel = getActualValue(mainDivWidth, positionWidthPercent);		
		var elemEnergyId = document.getElementById(energyId);	
		var elemDivId = document.getElementById(divId);	
		elemEnergyId.value = energyValue;
		elemDivId.style.width = divWidthPixel+"px";
		elemDivId.style.height = divWidthPixel+"px";
	}
	
	/************** To fix the design ************/
	fixDesignBS();
}


/**
 * Function add to restrict first bracket and enter as a input
 * @param evt
 */
function disableKey(evt) {
    var charCode = (evt.which) ? evt.which : event.keyCode;
    if (charCode == 40 || charCode == 41 || charCode == 13 || charCode == 96 || charCode == 126)
        return false;

    return true;
}


/**
 * Function add to allow only number as a input
 * @param evt
 */
function exceptNumberOnly(event, obj) {
	var key = window.event ? event.keyCode : event.which;
	if(event.keyCode == 13) {
		getEnergyValue(obj);
	} else if (event.keyCode == 8 || event.keyCode == 46
	 || event.keyCode == 37 || event.keyCode == 39 || event.keyCode == 9) {
	    return true;
	} else if ( key < 48 || key > 57 ) {
	    return false;
	} else if (key > 126) {
		return false;
	} else {
		return true;
	}
}


/**
 * Function add to get the input value
 * @param energyValue
 */
function getEnergyValue(obj) {
	var re = /[^0-9]/g;
	var divCount = $('#pageWrap').children('.single_task_item:visible').length;
	var maxValue = 100 - ((divCount-1)*5);
	if ( re.test(obj.value) ){  
		alertify.alert("<img src='../img/alert-icon.png'><br /><p>Only number is accepted for task value.</p>");
		calculateTotalEnergy();
	}
	if(obj.value<5){
		//alertify.alert("Please add the task value more than 5.");
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
		
		var unlockDivCount = $('#pageWrap').children('.single_task_item:visible').length;	//count the unlock div
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
 * Function add to remove the focus from the old tasks
 * @param null
 */
function removeFocus() {
	var divCountLoop = $('#pageWrap').children('div').length;
	for(var i=1; i<=divCountLoop;i++){
		var divIdImg = document.getElementById("divImg_" +i).id;
		$("#" + divIdImg).children().find(".add_task_area").removeClass("add_task_area_focus");
		$("#" + divIdImg).children().find(".add_task_area").removeClass("add_task_area_delete_focus");
		
	}
}

/**
 * Function call when the task exteed the
 * containment while resizing
 * @param null
 */
$('.main_warp').mouseup(function() {
	var windowSize = $(window).width(); 
	/*if(!(navigator.userAgent.match(/iPhone|iPad|iPod/i)) && !(navigator.userAgent.match(/Android/i)) ){
		if($('#pageWrap').width() > 1000){
			callInitialJsonOnResize();
		}
	} else {
		callInitialJsonOnResize();
	}*/
	disableDraggableResizable();
	callInitialJsonOnResize();
});

/**
 * Function add to to risize the task 
 * on clicking on the outside of the div
 * @param null
 */
function callInitialJsonOnResize(){
	var taskHeightPx = "";
	var taskWidthPx =  "";
	var divCountLoop = $('#pageWrap').children('.before_sketch_task').length;
	for(var i=1; i<=divCountLoop;i++){
		taskHeightPx = $("#divImg_"+i).height();
		taskWidthPx = $("#divImg_"+i).width();
		var style = document.getElementById("divImg_" +i).getAttribute("style");
		if((style.search('block') != -1)){
			if(navigator.userAgent.match(/Android/i)) {
				/*if(taskHeightPx != taskWidthPx){
					calculateTotalEnergy();
				} else*/ if(taskWidthPx > 215 || taskWidthPx < 74) {
					calculateTotalEnergy();
				} 
			} else if(navigator.userAgent.match(/iPad/i)) {
				if(taskWidthPx > 335 || taskWidthPx < 112) {
					calculateTotalEnergy();
				} /*else if(taskHeightPx != taskWidthPx) {
					calculateTotalEnergy();
				} */
			} else if(!(navigator.userAgent.match(/iPhone|iPad|iPod/i)) && !(navigator.userAgent.match(/Android/i)) ){		
				/*if(taskHeightPx != taskWidthPx){
					calculateTotalEnergy();
				} else*/ if(taskWidthPx > 395 || taskWidthPx < 140) {
					calculateTotalEnergy();
				} 
			}
		}				
	}
}

/**
 * Function to store the json string  
 * while navigate to previous page
 * @param null
 */
function backToPreviousPage() {	
	sessionStorage.setItem("bktp", "Y");
	sessionStorage.setItem("pageSequence", 1);
	goToNextPage();
}

/**
 * Function removes the session items which will not be used
 */
function removePresentSessionItems() {
	sessionStorage.removeItem("div_intial_value");
	sessionStorage.removeItem("total_json_add_task");
	if (sessionStorage.getItem("bktp")) {
		sessionStorage.removeItem("bktp");
	}
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
	//deviceWidth = 505;
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

/*function getTaskWidthValue(mainDivWidth, eachTaskWidth) {
	var taskWidth = Math.round((eachTaskWidth/1021)* mainDivWidth);
	return taskWidth;
	
}*/


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
	diffValuePixel = Math.round((mainDivWidth * diffValue) /1021);
	return diffValuePixel;
}

/***************** BEFORE SKETCH EDIT *******************/
/**
 * function handles when cancel is pressed
 */
function cancelForEdit() {
	alertify.set({ buttonReverse: true });
	alertify.set({ labels: {
	    ok     : "Yes",
	    cancel : "No"
	} });
	alertify.confirm("<img src='../img/confirm-icon.png'><br /><p>Are you sure you want to cancel? <br /> Nothing will be saved in search database.</p>", function (e) {
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
 * Function helps in restarting the link
 */
function restartExcersize() {
	alertify.set({ buttonReverse: true });
	alertify.set({ labels: {
	    ok     : "Yes",
	    cancel : "No"
	} });
	alertify.confirm("<img src='../img/confirm-icon.png'><br /><p>Are you sure to restart the exercise? <br /><b>Note: </b>This will delete your current Before Sketch and After Diagram.</p>", function (e) {
	    if (e) {
	    	sessionStorage.setItem("restartExcersize",1);
	    	$(".loader_bg").fadeIn();
	    	$(".loader").fadeIn();
	    	var navigation = Spine.Model.sub();
	    	navigation.configure("/user/navigation/restartExcersize", "jobReferenceNo", "emailId", "distinction");
	    	navigation.extend( Spine.Model.Ajax );
	    	//Populate the model
	    	var replyModel = new navigation({
	    		jobReferenceNo: sessionStorage.getItem("jobReferenceNo"),
	    		emailId: sessionStorage.getItem("jctEmail"),
	    		distinction: "Y"
	    	});
	    	replyModel.save(); //POST
	    	
	    	navigation.bind("ajaxError", function(record, xhr, settings, error) {
	    		alertify.alert("Unable to connect to the server.");
	    		$(".loader_bg").fadeOut();
	    	    $(".loader").hide();
	    	});
	    	
	    	navigation.bind("ajaxSuccess", function(record, xhr, settings, error) {
	    		var jsonStr = JSON.stringify(xhr);
	    		var obj = jQuery.parseJSON(jsonStr);
	    		if (obj.statusCode == "200") {
	    			if (sessionStorage.getItem("total_json_add_task")) {
	    				sessionStorage.removeItem("total_json_add_task");
	    			}
	    			if (sessionStorage.getItem("total_count")) {
	    				sessionStorage.removeItem("total_count");
	    			}
	    			if (sessionStorage.getItem("div_intial_value")) {
	    				sessionStorage.removeItem("div_intial_value");
	    			}
	    			sessionStorage.setItem("next_page", obj.lastPage);
	    			sessionStorage.setItem("bsView", "D");
	    			sessionStorage.setItem("as1View", "D");
	    			sessionStorage.setItem("as2View", "D");
	    			sessionStorage.setItem("profileId", obj.profileId);
	    			sessionStorage.setItem("groupId", obj.groupId);
	    			sessionStorage.setItem("totalTime",obj.totalTime);
	    			sessionStorage.setItem("rowIdentity", obj.identifier);
	    			sessionStorage.setItem("jobTitle", obj.jobTitle);
	    			sessionStorage.setItem("isCompleted", 0);
	    			sessionStorage.setItem("pageSequence", 1);
	    			sessionStorage.setItem("myPage", "BS");
	    			if (sessionStorage.getItem("giveUserOption")) {
	    				sessionStorage.removeItem("giveUserOption");
	    			}	    			
	    			window.location = "beforeSketch.jsp";
	    		} else {
	    			alertify.alert("<img src='../img/alert-icon.png'><br /><p>"+ob.statusDesc+"</p>");
	    			$(".loader_bg").fadeOut();
	    		    $(".loader").hide();
	    		}
	    	});
		    } 
		});
}

/**
 * Function add to fix the NEW task position which is added
 * @param divWidth
 * @param divId
 */
function addTaskPosition(divWidth, divId) {
	
	 var topPosition = 0;
	 var maxTopDivId = "";
	 var maxLeftPosition = 0;
	 var maxFinalLeft = 0;
	 var maxFinalBottom = 0; 
	 /**
	  * calculate the highest final top among all the tasks
	  * store the highest final height
	  * store the div id of highest final top element
	  **/
	  $("#pageWrap").children('.before_sketch_task:visible').each(function() {		
	        var element = $( this );       
	        elementId = $(this).attr('id');	  
			var position = element.position();	 	        	
	        if(topPosition <= position.top && elementId != divId) {
	        	topPosition = position.top;	
	        	maxTopDivId = elementId;  
				maxLeftPosition = position.left + $(this).width();				
	        }	        		       
	    });
		
	
	  /**
		  * calculate the highest final top among all the Task Divs
		  * store the highest final height
		  * store the div id of highest final top role frame
	  **/
		  
		var maxFinalTop =  $("#" + maxTopDivId).position().top;
		maxFinalBottom =  $("#" + maxTopDivId).position().top + $("#" + maxTopDivId).height();
		var parentBottom = parseInt($("#pageWrap").position().top) + parseInt($("#pageWrap").height());		
		maxFinalLeft = maxLeftPosition;
		
		/*** check if any other task is leaning on the left and bottom of the selected div ***/
		$("#pageWrap").children('.before_sketch_task:visible').each(function() {
			
			 var element = $( this );   
			 var elementId = $(this).attr('id');
		     var position = element.position();	
		     var eachTaskBottom = position.top + $(this).height();
			 var eachTaskLeft = position.left + $(this).width();						
			 if((eachTaskBottom > maxFinalTop && eachTaskBottom < parentBottom) && (elementId != divId)) {		
				 if((maxLeftPosition < eachTaskLeft)){
					 if(maxFinalLeft < eachTaskLeft){
							maxFinalLeft = eachTaskLeft;
						} 						
				 } else {					
					if(eachTaskBottom > (maxFinalBottom+10)) {						
						maxFinalBottom = eachTaskBottom;
					}	
				 }
			 }			 
		});
		calculateFreeArea(maxFinalTop,maxFinalLeft,maxFinalBottom,divWidth, divId);

}

/**
 * Function added to calculate the free area inside #pageWrap 
 * for the added task
 * @param maxFinalTop
 * @param maxFinalLeft
 * @param maxFinalBottom
 * @param divWidth
 * @param divId
 */
function calculateFreeArea(maxFinalTop,maxFinalLeft,maxFinalBottom,divWidth, divId) {
	 var freeArea = "";
	 if(!(navigator.userAgent.match(/iPhone|iPad|iPod/i)) && !(navigator.userAgent.match(/Android/i)) ){
		 freeArea = $("#pageWrap").width() - maxFinalLeft - 45;
	 } else {
		 freeArea = $("#pageWrap").width() - maxFinalLeft - 20;
	 }	
	 /*** Task added in the right side of the existing task ***/
	 if(divWidth  < freeArea) {	
		$("#" + divId).css({ position : "absolute"}); 
		$("#" + divId).css({ top : maxFinalTop+'px'});					
		$("#" + divId).css({ left : (maxFinalLeft+30) +'px'}); 
		adjustTaskWhileResize();
		increaseOuterDivWhileAdd(divId);
	 }
	 /*** Task added at the bottom of the existing task ***/
	 else {
		$("#" + divId).css({ position : "absolute"}); 
		$("#" + divId).css({ top : (maxFinalBottom + 20)+'px'});					
		$("#" + divId).css({ left : '20px'}); 
		adjustTaskWhileResize();
		increaseOuterDivWhileAdd(divId);
	 }
}

/**
* Function add to increase the outer div area
* while any task is added at the bottom
* @param divId
*/
function increaseOuterDivWhileAdd(divId) {
	var outerDivHeight = $('#pageWrap').height();
	
	var maxHeight = parseInt($("#" + divId).position().top) + parseInt($("#" + divId).height());
	 if (maxHeight > outerDivHeight){				
		maxHeight = maxHeight + 30;
		$("#pageWrap").css({'height':(maxHeight+'px')});						
	 }	
	 // $("#pageWrap").css({'height':($("#pageWrapCopy").height()+'px')});
}

/**
* Function added to fix the task position 
* if it exceed the outer div area
* while deleting or resizing any task
* 
*/
function adjustTaskWhileResize() {	
		var divCountLoop = $('#pageWrap').children('.before_sketch_task').length;	
		var mainDivWidth = "";
		for (var j = 1; j<=divCountLoop; j++){	
			var taskId = "divImg_" + j;	 
				var pos = $("#" + taskId).position();				
				var addition = pos.left + $('#'+ taskId).width();
				if(navigator.userAgent.match(/Android|iPhone/i)) {
					mainDivWidth = $("#pageWrap").width() - 20;
				} 
				else if(navigator.userAgent.match(/IPad|iPad/i)){
					mainDivWidth = $("#pageWrap").width() - 10;
				}
				else {
					mainDivWidth = $("#pageWrap").width() - 35;
				}
				if(addition > mainDivWidth) {
					var newLeft = addition - mainDivWidth;
					var newLeftPosition = pos.left - newLeft;
					$("#" + taskId).css({ left : newLeftPosition+'px'}); 
				}
							
		}	
}
/**
 *	Function added to change the 
 *	position from Relative -to- Absolute
 *	when default 3 task is loaded  
 **/
function relativeToAbsolute(){
	var divTop = new Array();
	var divLeft = new Array();
	var count = 0;
	$("#pageWrap").children('.before_sketch_task:visible').each(function() {
        var element = $( this );       
        var elementId = $(this).attr('id');
		var position = element.position();
		divTop[count] = position.top;
		divLeft[count] = position.left;
		count++;
	});
	count = 0;
	$("#pageWrap").children('.before_sketch_task:visible').each(function() {
		var element = $( this );       
        var elementId = $(this).attr('id');
		$("#" + elementId).css({'top': (divTop[count]+'px'), 'left': (divLeft[count]+'px'), position:'absolute'});
		count++;
	});
	setOuterDivWhileDelete("increasePageWrap");
}

/**
 * 	function added to adjust the size of the outer div
 * 	when any task is resized , added or deleted
 *  @param choice
 * */
function setOuterDivWhileDelete(choice) {
	var maxHeight = 0;
	var outerDivHeight = $("#pageWrap").height();
	$("#pageWrap").children('.before_sketch_task:visible').each(function() {		
        var element = $( this );       
        elementId = $(this).attr('id');	  	   
        var position = element.position();	
        if(maxHeight < position.top + $(this).height() ) {
        	 maxHeight = position.top + $(this).height();				
        }	        		       
    });
	maxHeight = maxHeight + 25;
	if( choice == 'increasePageWrap'){						//	When outerDiv need to Increase		
		if(maxHeight > outerDivHeight){
			$("#pageWrap").css({'height':(maxHeight+'px')});
		}
	} else {												//	When outerDiv need to Decrease
		if(maxHeight < outerDivHeight){
			$("#pageWrap").css({'height':(maxHeight+'px')});
		}
	}	
}

/**
 * This function is executed for device only
 * Function add to disable the drag and resize functinality 
 * for task block
 * @param null
 */
function disableDraggableResizable(){
if((navigator.userAgent.match(/iPhone|iPad|iPod/i)) || (navigator.userAgent.match(/Android/i)) ) {
	$(".before_sketch_task").children(".ui-resizable-se").css("display","none");
	$(".before_sketch_task").draggable( "option", "disabled", true );
	$(".before_sketch_task").resizable( "option", "disabled", true ); 
}
}

/**
 * This function is executed for device only
 * Function add to enable the drag and resize functinality 
 * for task block
 * @param object
 */
function enableDraggableResizable(divId){
	if((navigator.userAgent.match(/iPhone|iPad|iPod/i)) || (navigator.userAgent.match(/Android/i)) ) {
		var lockUnlockImg = document.getElementById(divId).getElementsByTagName("img")[1].getAttribute("src");
		 if (lockUnlockImg.search('imgUnlock.png') != -1) {	
			 $("#" +divId).children(".ui-resizable-se").css("display","block");
				$("#" +divId).draggable( "option", "disabled", false );
				$("#" +divId).resizable( "option", "disabled", false );
		 } else {
			 $("#" +divId).draggable( "option", "disabled", false );
		 }
	}
}

function setTaskOnPullOut() {
	$("#pageWrap").children('.before_sketch_task:visible').each(function() {		
        var element = $( this );       
        elementId = $(this).attr('id');	  	   
        var position = element.position();	       
        $("#" + elementId).css({ position : "relative"});       
        $("#" + elementId).css({ left : "10px"}); 
        $("#" + elementId).css({ top : "0px"}); 
    });
	//setOuterDivWhileDelete("increasePageWrap");
}

/**
 * This Function save the data in session 
 * while the page refresh
 */
function saveBeforeSketchWhileRefresh() {
	var outerDivHeight = $('#pageWrap').height();
	sessionStorage.setItem("bsDivHeight", outerDivHeight);		    	
   	nextPageNavigation = 1;
   	var cnt = 1;
   	var totalJson = [];
   	var idx = 0;
   	var divValue = document.getElementById("divImg_" + cnt);
   	
   	var mainDivWidth = $("#pageWrap").width();
   	if (!(navigator.userAgent.match(/iPhone|iPad|iPod/i)) && !(navigator.userAgent.match(/Android/i))) {
   		if(mainDivWidth < 1000) {
  			mainDivWidth = 1021;
   		}
  	}
	var mainDivHeight = $("#pageWrap").height();	
	    	    
	var gcd=calc(mainDivWidth,mainDivHeight);
	var r1=mainDivWidth/gcd;
	var r2=mainDivHeight/gcd;
	var ratio=r1+":"+r2;
		    	     	   		    	
	while(divValue != null) {
		var unitJson = {};					
		unitJson["divId"] = document.getElementById("divImg_" + cnt).id;
		unitJson["divIdx"] = document.getElementById("divImg_" + cnt +"_x").id;
		unitJson["energyId"] = document.getElementById("energyId_"+ cnt).id;
		unitJson["energyValue"] = document.getElementById("energyId_"+ cnt).value;
		unitJson["style"] = document.getElementById("divImg_" + cnt).getAttribute("style");
		unitJson["areaId"] = document.getElementById("areaId_"+ cnt).id;
		unitJson["areaValue"] = document.getElementById("areaId_"+ cnt).value.replace(/,/g, "`").replace(/#/g, "-_-").replace(/\?/g, ";_;");
		unitJson["positionLeft"] = document.getElementById("divImg_" + cnt).style.left;		    					    			
		unitJson["positionTop"] = document.getElementById("divImg_" + cnt).style.top;
		unitJson["positionWidth"] = document.getElementById("divImg_" + cnt).style.width;
		unitJson["positionHeight"] = document.getElementById("divImg_" + cnt).style.height;
		var positionLeftPixel = document.getElementById("divImg_" + cnt).style.left;				 
		unitJson["positionLeftPixel"] = positionLeftPixel;
		unitJson["positionLeftPercent"] = getPercentageValue(mainDivWidth, positionLeftPixel);
		var positionTopPixel = document.getElementById("divImg_" + cnt).style.top;	
		unitJson["positionTopPixel"] = positionTopPixel;
		unitJson["positionTopPercent"] = getPercentageValue(mainDivHeight, positionTopPixel);
		var positionWidthPixel = document.getElementById("divImg_" + cnt).style.width;	
		var positionWidth = parseInt(positionWidthPixel.replace("px", ''));
		unitJson["positionWidthPercent"] = getPercentageValue(mainDivWidth, positionWidth);
		unitJson["mainDivWidth"] = mainDivWidth;
		unitJson["mainDivHeight"] = mainDivHeight;
		unitJson["aspectRatio"] = ratio;		    		
		var position = $("#divImg_" + cnt).position();
		var positionLeftNewPercent = getPercentageValue(mainDivWidth, position.left);
		var positionTopNewPercent = getPercentageValue(mainDivWidth, position.top);
		unitJson["positionLeftNew"] = positionLeftNewPercent;		    					    			
		unitJson["positionTopNew"] = positionTopNewPercent;
		
		/***** Added for lock and unlock feature *****/
		unitJson["divLockId"] = document.getElementById("divImg_" + cnt +"_lock").id;
		unitJson["lockImage"] = document.getElementById("divImg_" + cnt+ "_lock").getElementsByTagName("img")[0].getAttribute("src");
				
		totalJson[idx++] = unitJson;
		divValue = document.getElementById("divImg_" + ++cnt);
		}		
		if (sessionStorage.getItem("restartExcersize") == 0) {
			ClientSession.set("total_json_add_task",JSON.stringify(totalJson));
		} else {
			sessionStorage.removeItem("total_json_add_task");
			sessionStorage.setItem("restartExcersize",0);
		}
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
	$("#" + divId).addClass("single_task_item");
	$("#" + divId).removeClass("lock_single_task_item");
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
	$("#" + divId).addClass("single_task_item");
	$("#" + divId).removeClass("lock_single_task_item");
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
	$("#" + divId).removeClass("single_task_item");
	$("#" + divId).addClass("lock_single_task_item");
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
	var unlockDivCount = $('#pageWrap').children('.single_task_item:visible').length;	//count the unlock div
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