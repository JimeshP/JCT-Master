/** 
 * This function is call for before sketch only
 * This function add to fix the css 
 * for each task block on resize or add or delete any task 
 * @param null
 */
function fixDesignBS() {	
	var divCountLoop = $('#pageWrap').children('.before_sketch_task').length;
	var energyValue = "";
	for(var i=1; i<=divCountLoop;i++){
		var style = document.getElementById("divImg_" +i).getAttribute("style");	
		if((style.search('block') != -1)){			
			energyValue = document.getElementById("energyId_" +i).value;			
			var divIdImg = document.getElementById("divImg_" +i).id;
			if(navigator.userAgent.match(/Android/i)){ // Android 
				if (windowSize>550) { // s3, nexus etc window width above 550                  
				if(energyValue<=9 && energyValue >=5) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");						
					$("#" + divIdImg).removeClass("uptoBs_10");
									
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").addClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_smaller");												

				} else if (energyValue == 10) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					$("#" + divIdImg).removeClass("uptoBs_10");
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").addClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_small");	
				} else if(energyValue<=25 && energyValue >10) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					$("#" + divIdImg).removeClass("uptoBs_10");
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".energy").addClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_small");
				} else if(energyValue<=40 && energyValue >25) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");	
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");				
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).removeClass("uptoBs_10");
				} else if(energyValue<=60 && energyValue >40) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					$("#" + divIdImg).removeClass("uptoBs_10");
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_middle");
				} else if(energyValue<=75 && energyValue >60) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					$("#" + divIdImg).removeClass("uptoBs_10");
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_large");
				} else if(energyValue >75) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).removeClass("uptoBs_10");
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_larger");
				}	
				/********** ISSUE# 207 START *******/
				if (energyValue<=10 && energyValue >=5) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '-22px'});
				} else if(energyValue<=15 && energyValue >10) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '-18px'});
				} else if(energyValue<=20 && energyValue >15) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '-16px'});
				} else if(energyValue<=35 && energyValue >20) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '-12px'});
				} else if (energyValue<=40 && energyValue >35) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '-10px'});
				} else if (energyValue<=50 && energyValue >40) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '-7px'});
				} else if (energyValue<=55 && energyValue >50) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '-4px'});
				} else if (energyValue<=60 && energyValue >55) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '-2px'});
				} else {
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0px'});
				}
				/********** ISSUE# 207 END *******/	
						
				} else{ // htc windowwidth below 550			                  
				if(energyValue<=9 && energyValue >=5) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");					
					
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").addClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_smaller");	
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_smaller");
					
					$("#" + divIdImg).addClass("bs_upto_9");
					
				} else if (energyValue == 10) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").addClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_smaller");
					
					$("#" + divIdImg).addClass("bs_upto_10");
				} else if(energyValue<=15 && energyValue >10) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".energy").addClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_small");
					
					$("#" + divIdImg).addClass("bs_upto_15");
				} else if(energyValue<=20 && energyValue >15) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".energy").addClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_small");
					$("#" + divIdImg).addClass("bs_upto_20");
				} else if(energyValue<=25 && energyValue >20) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
			
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".energy").addClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_small");
					
					$("#" + divIdImg).addClass("bs_upto_25");
				} else if(energyValue<=30 && energyValue >25) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");	
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
				
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					$("#" + divIdImg).addClass("bs_upto_30");					
				} else if(energyValue<=35 && energyValue >30) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");	
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");

					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
										
					$("#" + divIdImg).addClass("bs_upto_35");
				} else if(energyValue<=40 && energyValue >35) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");	
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");		
					
					$("#" + divIdImg).addClass("bs_upto_40");					
				} else if(energyValue<=45 && energyValue >40) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");

					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_middle");
					
					$("#" + divIdImg).addClass("bs_upto_45");
				} else if(energyValue<=50 && energyValue >45) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
				
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_middle");
					
					$("#" + divIdImg).addClass("bs_upto_50");
				} else if(energyValue<=55 && energyValue >50) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
	
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_middle");
					
					$("#" + divIdImg).addClass("bs_upto_55");
				} else if(energyValue<=60 && energyValue >55) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");

					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_9");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_middle");
					
					$("#" + divIdImg).addClass("bs_upto_60");
				} else if(energyValue<=65 && energyValue >60) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");

					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_large");
					
					$("#" + divIdImg).addClass("bs_upto_65");
				} else if(energyValue<=70 && energyValue >65) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");

					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_large");
					
					$("#" + divIdImg).addClass("bs_upto_70");
				} else if(energyValue<=75 && energyValue >70) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");

					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_large");
					
					$("#" + divIdImg).addClass("bs_upto_75");
				} else if(energyValue<=80 && energyValue >75) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");

					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_larger");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_larger");
					
					$("#" + divIdImg).addClass("bs_upto_80");
				} else if(energyValue<=84 && energyValue >80) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");

					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_larger");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_larger");
					
					$("#" + divIdImg).addClass("bs_upto_85");
				} else if(energyValue >84) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");

					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_9");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_larger");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_larger");
					
					$("#" + divIdImg).addClass("bs_upto_90");
				}
				if (energyValue<=10 && energyValue >=5) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '-5px'});					
				} else if(energyValue<=15 && energyValue >10) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '-5px'});
				}  else if(energyValue<=20 && energyValue >15) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '-2px'});
				} else if(energyValue<=35 && energyValue >20) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '-1px'});
				} else if (energyValue<=40 && energyValue >35) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '-2px'});
				} else if (energyValue<=50 && energyValue >40) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0px'});
				} else if (energyValue<=55 && energyValue >50) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0px'});
				} else if (energyValue<=60 && energyValue >55) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0px'});
				} else if (energyValue<=65 && energyValue >60) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0px'});
				} else {
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0px'});
				}
				/********** ISSUE# 207 END *******/ 
				}
			}else if(navigator.userAgent.match(/iPhone/i)){ // iPhone
				if(energyValue<=9 && energyValue >=5) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");					
					
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").addClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_smaller");	
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_smaller");
					
					$("#" + divIdImg).addClass("bs_upto_9");					
				} else if (energyValue == 10) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").addClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_smaller");
					
					$("#" + divIdImg).addClass("bs_upto_10");
				} else if(energyValue<=15 && energyValue >10) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".energy").addClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_small");
					
					$("#" + divIdImg).addClass("bs_upto_15");
				} else if(energyValue<=20 && energyValue >15) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".energy").addClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_small");
					$("#" + divIdImg).addClass("bs_upto_20");
				} else if(energyValue<=25 && energyValue >20) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
			
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".energy").addClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_small");
					
					$("#" + divIdImg).addClass("bs_upto_25");
				} else if(energyValue<=30 && energyValue >25) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");	
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
				
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					$("#" + divIdImg).addClass("bs_upto_30");					
				} else if(energyValue<=35 && energyValue >30) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");	
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");

					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
										
					$("#" + divIdImg).addClass("bs_upto_35");
				} else if(energyValue<=40 && energyValue >35) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");	
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");		
					
					$("#" + divIdImg).addClass("bs_upto_40");					
				} else if(energyValue<=45 && energyValue >40) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");

					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_middle");
					
					$("#" + divIdImg).addClass("bs_upto_45");
				} else if(energyValue<=50 && energyValue >45) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
				
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_middle");
					
					$("#" + divIdImg).addClass("bs_upto_50");
				} else if(energyValue<=55 && energyValue >50) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
	
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_middle");
					
					$("#" + divIdImg).addClass("bs_upto_55");
				} else if(energyValue<=60 && energyValue >55) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");

					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_9");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_middle");
					
					$("#" + divIdImg).addClass("bs_upto_60");
				} else if(energyValue<=65 && energyValue >60) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");

					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_large");
					
					$("#" + divIdImg).addClass("bs_upto_65");
				} else if(energyValue<=70 && energyValue >65) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");

					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_large");
					
					$("#" + divIdImg).addClass("bs_upto_70");
				} else if(energyValue<=75 && energyValue >70) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");

					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_large");
					
					$("#" + divIdImg).addClass("bs_upto_75");
				} else if(energyValue<=80 && energyValue >75) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");

					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_larger");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_larger");
					
					$("#" + divIdImg).addClass("bs_upto_80");
				} else if(energyValue<=84 && energyValue >80) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");

					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_larger");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_larger");
					
					$("#" + divIdImg).addClass("bs_upto_85");
				} else if(energyValue >84) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");

					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_9");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_larger");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_larger");
					
					$("#" + divIdImg).addClass("bs_upto_90");
				}
				if (energyValue<=10 && energyValue >=5) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '-5px'});
				} else if(energyValue<=15 && energyValue >10) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '-5px'});
				}  else if(energyValue<=20 && energyValue >15) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '-2px'});
				} else if(energyValue<=35 && energyValue >20) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '-1px'});
				} else if (energyValue<=40 && energyValue >35) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '-2px'});
				} else if (energyValue<=50 && energyValue >40) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0px'});
				} else if (energyValue<=55 && energyValue >50) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0px'});
				} else if (energyValue<=60 && energyValue >55) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0px'});
				} else if (energyValue<=65 && energyValue >60) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0px'});
				} else {
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0px'});
				}
				/********** ISSUE# 207 END *******/
			}else if(navigator.userAgent.match(/iPad/i)){ // IPAD
				if(energyValue<=9 && energyValue >=5) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");					
					
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").addClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_smaller");	
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_smaller");
					
					$("#" + divIdImg).addClass("bs_upto_9");					
				} else if (energyValue == 10) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").addClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_smaller");
					
					$("#" + divIdImg).addClass("bs_upto_10");
				} else if(energyValue<=15 && energyValue >10) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".energy").addClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_small");
					
					$("#" + divIdImg).addClass("bs_upto_15");
				} else if(energyValue<=20 && energyValue >15) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".energy").addClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_small");
					$("#" + divIdImg).addClass("bs_upto_20");
				} else if(energyValue<=25 && energyValue >20) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
			        
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".energy").addClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_small");
					
					$("#" + divIdImg).addClass("bs_upto_25");
				} else if(energyValue<=30 && energyValue >25) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");	
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
				
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");
										
					$("#" + divIdImg).addClass("bs_upto_30");					
				} else if(energyValue<=35 && energyValue >30) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");	
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");

					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
										
					$("#" + divIdImg).addClass("bs_upto_35");
				} else if(energyValue<=40 && energyValue >35) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");	
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");		
					
					$("#" + divIdImg).addClass("bs_upto_40");					
				} else if(energyValue<=45 && energyValue >40) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
                    
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_middle");
					
					$("#" + divIdImg).addClass("bs_upto_45");
				} else if(energyValue<=50 && energyValue >45) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_middle");
					
					$("#" + divIdImg).addClass("bs_upto_50");
				} else if(energyValue<=55 && energyValue >50) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
	
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_middle");
					
					$("#" + divIdImg).addClass("bs_upto_55");
				} else if(energyValue<=60 && energyValue >55) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");

					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_9");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_middle");
					
					$("#" + divIdImg).addClass("bs_upto_60");
				} else if(energyValue<=65 && energyValue >60) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");

					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_large");
					
					$("#" + divIdImg).addClass("bs_upto_65");
				} else if(energyValue<=70 && energyValue >65) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");

					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_large");
					
					$("#" + divIdImg).addClass("bs_upto_70");
				} else if(energyValue<=75 && energyValue >70) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");

					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_large");
					
					$("#" + divIdImg).addClass("bs_upto_75");
				} else if(energyValue<=80 && energyValue >75) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");

					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_larger");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_larger");
					
					$("#" + divIdImg).addClass("bs_upto_80");
				} else if(energyValue<=84 && energyValue >80) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");

					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_larger");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_larger");
					
					$("#" + divIdImg).addClass("bs_upto_85");
				} else if(energyValue >84) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");

					$("#" + divIdImg).removeClass("bs_upto_9");
					$("#" + divIdImg).removeClass("bs_upto_10");
					$("#" + divIdImg).removeClass("bs_upto_15");
					$("#" + divIdImg).removeClass("bs_upto_20");
					$("#" + divIdImg).removeClass("bs_upto_25");
					$("#" + divIdImg).removeClass("bs_upto_30");
					$("#" + divIdImg).removeClass("bs_upto_35");
					$("#" + divIdImg).removeClass("bs_upto_40");
					$("#" + divIdImg).removeClass("bs_upto_45");
					$("#" + divIdImg).removeClass("bs_upto_50");
					$("#" + divIdImg).removeClass("bs_upto_55");
					$("#" + divIdImg).removeClass("bs_upto_60");					
					$("#" + divIdImg).removeClass("bs_upto_65");
					$("#" + divIdImg).removeClass("bs_upto_70");
					$("#" + divIdImg).removeClass("bs_upto_75");
					$("#" + divIdImg).removeClass("bs_upto_80");
					$("#" + divIdImg).removeClass("bs_upto_85");
					$("#" + divIdImg).removeClass("bs_upto_9");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_larger");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_larger");
					
					$("#" + divIdImg).addClass("bs_upto_90");
				}
				if (energyValue<=10 && energyValue >=5) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '-1px'});
				} else if(energyValue<=15 && energyValue >10) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '-1px'});
				}  else if(energyValue<=20 && energyValue >15) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '-1px'});
				} else if(energyValue<=35 && energyValue >20) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '-1px'});
				} else if (energyValue<=40 && energyValue >35) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '-2px'});
				} else if (energyValue<=50 && energyValue >40) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0px'});
				} else if (energyValue<=55 && energyValue >50) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0px'});
				} else if (energyValue<=60 && energyValue >55) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0px'});
				} else if (energyValue<=65 && energyValue >60) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0px'});
				} else {
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0px'});
				}
				/********** ISSUE# 207 END *******/				
								
			}else {  // desktop 				
				if(energyValue<=9 && energyValue >=5) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").addClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_smaller");
					
					$("#" + divIdImg).addClass("uptoDescBS_9");
					
					$("#" + divIdImg).removeClass("uptoDescBS_10");					
					$("#" + divIdImg).removeClass("uptoDescBS_15");					
					$("#" + divIdImg).removeClass("uptoDescBS_20");					
					$("#" + divIdImg).removeClass("uptoDescBS_25");
					$("#" + divIdImg).removeClass("uptoDescBS_30");
					$("#" + divIdImg).removeClass("uptoDescBS_35");
					$("#" + divIdImg).removeClass("uptoDescBS_40");
					$("#" + divIdImg).removeClass("uptoDescBS_45");
					$("#" + divIdImg).removeClass("uptoDescBS_50");
					$("#" + divIdImg).removeClass("uptoDescBS_55");
					$("#" + divIdImg).removeClass("uptoDescBS_60");					
					$("#" + divIdImg).removeClass("uptoDescBS_65");
					$("#" + divIdImg).removeClass("uptoDescBS_70");
					$("#" + divIdImg).removeClass("uptoDescBS_75");
					$("#" + divIdImg).removeClass("uptoDescBS_80");
					$("#" + divIdImg).removeClass("uptoDescBS_90");			
				} else if (energyValue == 10) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").addClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_small");	
					
					$("#" + divIdImg).addClass("uptoDescBS_10");
					
					$("#" + divIdImg).removeClass("uptoDescBS_9");					
					$("#" + divIdImg).removeClass("uptoDescBS_15");	
					$("#" + divIdImg).removeClass("uptoDescBS_20");					
					$("#" + divIdImg).removeClass("uptoDescBS_25");
					$("#" + divIdImg).removeClass("uptoDescBS_30");
					$("#" + divIdImg).removeClass("uptoDescBS_35");
					$("#" + divIdImg).removeClass("uptoDescBS_40");
					$("#" + divIdImg).removeClass("uptoDescBS_45");
					$("#" + divIdImg).removeClass("uptoDescBS_50");
					$("#" + divIdImg).removeClass("uptoDescBS_55");
					$("#" + divIdImg).removeClass("uptoDescBS_60");					
					$("#" + divIdImg).removeClass("uptoDescBS_65");
					$("#" + divIdImg).removeClass("uptoDescBS_70");
					$("#" + divIdImg).removeClass("uptoDescBS_75");
					$("#" + divIdImg).removeClass("uptoDescBS_80");
					$("#" + divIdImg).removeClass("uptoDescBS_90");
				} else if(energyValue<=15 && energyValue >10) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".energy").addClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_small");
					
					$("#" + divIdImg).addClass("uptoDescBS_15");
					
					$("#" + divIdImg).removeClass("uptoDescBS_9");					
					$("#" + divIdImg).removeClass("uptoDescBS_10");						
					$("#" + divIdImg).removeClass("uptoDescBS_20");					
					$("#" + divIdImg).removeClass("uptoDescBS_25");
					$("#" + divIdImg).removeClass("uptoDescBS_30");
					$("#" + divIdImg).removeClass("uptoDescBS_35");
					$("#" + divIdImg).removeClass("uptoDescBS_40");
					$("#" + divIdImg).removeClass("uptoDescBS_45");
					$("#" + divIdImg).removeClass("uptoDescBS_50");
					$("#" + divIdImg).removeClass("uptoDescBS_55");
					$("#" + divIdImg).removeClass("uptoDescBS_60");					
					$("#" + divIdImg).removeClass("uptoDescBS_65");
					$("#" + divIdImg).removeClass("uptoDescBS_70");
					$("#" + divIdImg).removeClass("uptoDescBS_75");
					$("#" + divIdImg).removeClass("uptoDescBS_80");
					$("#" + divIdImg).removeClass("uptoDescBS_90");
				} else if(energyValue<=20 && energyValue >15) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".energy").addClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_small");
					
					$("#" + divIdImg).addClass("uptoDescBS_20");	
					
					$("#" + divIdImg).removeClass("uptoDescBS_9");
					$("#" + divIdImg).removeClass("uptoDescBS_10");					
					$("#" + divIdImg).removeClass("uptoDescBS_15");
					$("#" + divIdImg).removeClass("uptoDescBS_25");
					$("#" + divIdImg).removeClass("uptoDescBS_30");
					$("#" + divIdImg).removeClass("uptoDescBS_35");
					$("#" + divIdImg).removeClass("uptoDescBS_40");
					$("#" + divIdImg).removeClass("uptoDescBS_45");
					$("#" + divIdImg).removeClass("uptoDescBS_50");
					$("#" + divIdImg).removeClass("uptoDescBS_55");
					$("#" + divIdImg).removeClass("uptoDescBS_60");					
					$("#" + divIdImg).removeClass("uptoDescBS_65");
					$("#" + divIdImg).removeClass("uptoDescBS_70");
					$("#" + divIdImg).removeClass("uptoDescBS_75");
					$("#" + divIdImg).removeClass("uptoDescBS_80");
					$("#" + divIdImg).removeClass("uptoDescBS_90");
				} else if(energyValue<=25 && energyValue >20) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".energy").addClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_small");
					
					$("#" + divIdImg).addClass("uptoDescBS_25");	
					
					$("#" + divIdImg).removeClass("uptoDescBS_9");
					$("#" + divIdImg).removeClass("uptoDescBS_10");					
					$("#" + divIdImg).removeClass("uptoDescBS_15");
					$("#" + divIdImg).removeClass("uptoDescBS_20");
					$("#" + divIdImg).removeClass("uptoDescBS_30");
					$("#" + divIdImg).removeClass("uptoDescBS_35");
					$("#" + divIdImg).removeClass("uptoDescBS_40");
					$("#" + divIdImg).removeClass("uptoDescBS_45");
					$("#" + divIdImg).removeClass("uptoDescBS_50");
					$("#" + divIdImg).removeClass("uptoDescBS_55");
					$("#" + divIdImg).removeClass("uptoDescBS_60");					
					$("#" + divIdImg).removeClass("uptoDescBS_65");
					$("#" + divIdImg).removeClass("uptoDescBS_70");
					$("#" + divIdImg).removeClass("uptoDescBS_75");
					$("#" + divIdImg).removeClass("uptoDescBS_80");
					$("#" + divIdImg).removeClass("uptoDescBS_90");
				} else if(energyValue<=30 && energyValue >25) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");	
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");				
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					
					$("#" + divIdImg).addClass("uptoDescBS_30");	
					
					$("#" + divIdImg).removeClass("uptoDescBS_9");
					$("#" + divIdImg).removeClass("uptoDescBS_10");					
					$("#" + divIdImg).removeClass("uptoDescBS_15");
					$("#" + divIdImg).removeClass("uptoDescBS_20");
					$("#" + divIdImg).removeClass("uptoDescBS_25");
					$("#" + divIdImg).removeClass("uptoDescBS_35");
					$("#" + divIdImg).removeClass("uptoDescBS_40");
					$("#" + divIdImg).removeClass("uptoDescBS_45");
					$("#" + divIdImg).removeClass("uptoDescBS_50");
					$("#" + divIdImg).removeClass("uptoDescBS_55");
					$("#" + divIdImg).removeClass("uptoDescBS_60");					
					$("#" + divIdImg).removeClass("uptoDescBS_65");
					$("#" + divIdImg).removeClass("uptoDescBS_70");
					$("#" + divIdImg).removeClass("uptoDescBS_75");
					$("#" + divIdImg).removeClass("uptoDescBS_80");
					$("#" + divIdImg).removeClass("uptoDescBS_90");
				} else if(energyValue<=35 && energyValue >30) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");	
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");				
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					
					$("#" + divIdImg).addClass("uptoDescBS_35");	
					
					$("#" + divIdImg).removeClass("uptoDescBS_9");
					$("#" + divIdImg).removeClass("uptoDescBS_10");					
					$("#" + divIdImg).removeClass("uptoDescBS_15");
					$("#" + divIdImg).removeClass("uptoDescBS_20");
					$("#" + divIdImg).removeClass("uptoDescBS_25");
					$("#" + divIdImg).removeClass("uptoDescBS_30");
					$("#" + divIdImg).removeClass("uptoDescBS_40");
					$("#" + divIdImg).removeClass("uptoDescBS_45");
					$("#" + divIdImg).removeClass("uptoDescBS_50");
					$("#" + divIdImg).removeClass("uptoDescBS_55");
					$("#" + divIdImg).removeClass("uptoDescBS_60");					
					$("#" + divIdImg).removeClass("uptoDescBS_65");
					$("#" + divIdImg).removeClass("uptoDescBS_70");
					$("#" + divIdImg).removeClass("uptoDescBS_75");
					$("#" + divIdImg).removeClass("uptoDescBS_80");
					$("#" + divIdImg).removeClass("uptoDescBS_90");	
				} else if(energyValue<=40 && energyValue >35) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");	
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");				
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					
					$("#" + divIdImg).addClass("uptoDescBS_40");	
					
					$("#" + divIdImg).removeClass("uptoDescBS_9");
					$("#" + divIdImg).removeClass("uptoDescBS_10");					
					$("#" + divIdImg).removeClass("uptoDescBS_15");
					$("#" + divIdImg).removeClass("uptoDescBS_20");
					$("#" + divIdImg).removeClass("uptoDescBS_25");
					$("#" + divIdImg).removeClass("uptoDescBS_30");
					$("#" + divIdImg).removeClass("uptoDescBS_35");
					$("#" + divIdImg).removeClass("uptoDescBS_45");
					$("#" + divIdImg).removeClass("uptoDescBS_50");
					$("#" + divIdImg).removeClass("uptoDescBS_55");
					$("#" + divIdImg).removeClass("uptoDescBS_60");					
					$("#" + divIdImg).removeClass("uptoDescBS_65");
					$("#" + divIdImg).removeClass("uptoDescBS_70");
					$("#" + divIdImg).removeClass("uptoDescBS_75");
					$("#" + divIdImg).removeClass("uptoDescBS_80");
					$("#" + divIdImg).removeClass("uptoDescBS_90");	
				} else if(energyValue<=45 && energyValue >40) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					
					$("#" + divIdImg).addClass("uptoDescBS_45");	
					
					$("#" + divIdImg).removeClass("uptoDescBS_9");
					$("#" + divIdImg).removeClass("uptoDescBS_10");					
					$("#" + divIdImg).removeClass("uptoDescBS_15");
					$("#" + divIdImg).removeClass("uptoDescBS_20");
					$("#" + divIdImg).removeClass("uptoDescBS_25");
					$("#" + divIdImg).removeClass("uptoDescBS_30");
					$("#" + divIdImg).removeClass("uptoDescBS_35");
					$("#" + divIdImg).removeClass("uptoDescBS_40");
					$("#" + divIdImg).removeClass("uptoDescBS_50");
					$("#" + divIdImg).removeClass("uptoDescBS_55");
					$("#" + divIdImg).removeClass("uptoDescBS_60");					
					$("#" + divIdImg).removeClass("uptoDescBS_65");
					$("#" + divIdImg).removeClass("uptoDescBS_70");
					$("#" + divIdImg).removeClass("uptoDescBS_75");
					$("#" + divIdImg).removeClass("uptoDescBS_80");
					$("#" + divIdImg).removeClass("uptoDescBS_90");	
				} else if(energyValue<=50 && energyValue >45) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					
					$("#" + divIdImg).addClass("uptoDescBS_50");	
					
					$("#" + divIdImg).removeClass("uptoDescBS_9");
					$("#" + divIdImg).removeClass("uptoDescBS_10");					
					$("#" + divIdImg).removeClass("uptoDescBS_15");
					$("#" + divIdImg).removeClass("uptoDescBS_20");
					$("#" + divIdImg).removeClass("uptoDescBS_25");
					$("#" + divIdImg).removeClass("uptoDescBS_30");
					$("#" + divIdImg).removeClass("uptoDescBS_35");
					$("#" + divIdImg).removeClass("uptoDescBS_40");
					$("#" + divIdImg).removeClass("uptoDescBS_45");
					$("#" + divIdImg).removeClass("uptoDescBS_55");
					$("#" + divIdImg).removeClass("uptoDescBS_60");					
					$("#" + divIdImg).removeClass("uptoDescBS_65");
					$("#" + divIdImg).removeClass("uptoDescBS_70");
					$("#" + divIdImg).removeClass("uptoDescBS_75");
					$("#" + divIdImg).removeClass("uptoDescBS_80");
					$("#" + divIdImg).removeClass("uptoDescBS_90");
				} else if(energyValue<=55 && energyValue >50) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					
					$("#" + divIdImg).addClass("uptoDescBS_55");	
					
					$("#" + divIdImg).removeClass("uptoDescBS_9");
					$("#" + divIdImg).removeClass("uptoDescBS_10");					
					$("#" + divIdImg).removeClass("uptoDescBS_15");
					$("#" + divIdImg).removeClass("uptoDescBS_20");
					$("#" + divIdImg).removeClass("uptoDescBS_25");
					$("#" + divIdImg).removeClass("uptoDescBS_30");
					$("#" + divIdImg).removeClass("uptoDescBS_35");
					$("#" + divIdImg).removeClass("uptoDescBS_40");
					$("#" + divIdImg).removeClass("uptoDescBS_45");
					$("#" + divIdImg).removeClass("uptoDescBS_50");
					$("#" + divIdImg).removeClass("uptoDescBS_60");					
					$("#" + divIdImg).removeClass("uptoDescBS_65");
					$("#" + divIdImg).removeClass("uptoDescBS_70");
					$("#" + divIdImg).removeClass("uptoDescBS_75");
					$("#" + divIdImg).removeClass("uptoDescBS_80");
					$("#" + divIdImg).removeClass("uptoDescBS_90");
				} else if(energyValue<=60 && energyValue >55) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					
					$("#" + divIdImg).addClass("uptoDescBS_60");	
					
					$("#" + divIdImg).removeClass("uptoDescBS_9");
					$("#" + divIdImg).removeClass("uptoDescBS_10");					
					$("#" + divIdImg).removeClass("uptoDescBS_15");
					$("#" + divIdImg).removeClass("uptoDescBS_20");
					$("#" + divIdImg).removeClass("uptoDescBS_25");
					$("#" + divIdImg).removeClass("uptoDescBS_30");
					$("#" + divIdImg).removeClass("uptoDescBS_35");
					$("#" + divIdImg).removeClass("uptoDescBS_40");
					$("#" + divIdImg).removeClass("uptoDescBS_45");
					$("#" + divIdImg).removeClass("uptoDescBS_50");
					$("#" + divIdImg).removeClass("uptoDescBS_55");					
					$("#" + divIdImg).removeClass("uptoDescBS_65");
					$("#" + divIdImg).removeClass("uptoDescBS_70");
					$("#" + divIdImg).removeClass("uptoDescBS_75");
					$("#" + divIdImg).removeClass("uptoDescBS_80");
					$("#" + divIdImg).removeClass("uptoDescBS_90");
				} else if(energyValue<=65 && energyValue >60) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					
					$("#" + divIdImg).addClass("uptoDescBS_65");	
					
					$("#" + divIdImg).removeClass("uptoDescBS_9");
					$("#" + divIdImg).removeClass("uptoDescBS_10");					
					$("#" + divIdImg).removeClass("uptoDescBS_15");
					$("#" + divIdImg).removeClass("uptoDescBS_20");
					$("#" + divIdImg).removeClass("uptoDescBS_25");
					$("#" + divIdImg).removeClass("uptoDescBS_30");
					$("#" + divIdImg).removeClass("uptoDescBS_35");
					$("#" + divIdImg).removeClass("uptoDescBS_40");
					$("#" + divIdImg).removeClass("uptoDescBS_45");
					$("#" + divIdImg).removeClass("uptoDescBS_50");
					$("#" + divIdImg).removeClass("uptoDescBS_55");					
					$("#" + divIdImg).removeClass("uptoDescBS_60");
					$("#" + divIdImg).removeClass("uptoDescBS_70");
					$("#" + divIdImg).removeClass("uptoDescBS_75");
					$("#" + divIdImg).removeClass("uptoDescBS_80");
					$("#" + divIdImg).removeClass("uptoDescBS_90");
				} else if(energyValue<=70 && energyValue >65) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					
					$("#" + divIdImg).addClass("uptoDescBS_70");	
					
					$("#" + divIdImg).removeClass("uptoDescBS_9");
					$("#" + divIdImg).removeClass("uptoDescBS_10");					
					$("#" + divIdImg).removeClass("uptoDescBS_15");
					$("#" + divIdImg).removeClass("uptoDescBS_20");
					$("#" + divIdImg).removeClass("uptoDescBS_25");
					$("#" + divIdImg).removeClass("uptoDescBS_30");
					$("#" + divIdImg).removeClass("uptoDescBS_35");
					$("#" + divIdImg).removeClass("uptoDescBS_40");
					$("#" + divIdImg).removeClass("uptoDescBS_45");
					$("#" + divIdImg).removeClass("uptoDescBS_50");
					$("#" + divIdImg).removeClass("uptoDescBS_55");					
					$("#" + divIdImg).removeClass("uptoDescBS_60");
					$("#" + divIdImg).removeClass("uptoDescBS_65");
					$("#" + divIdImg).removeClass("uptoDescBS_75");
					$("#" + divIdImg).removeClass("uptoDescBS_80");
					$("#" + divIdImg).removeClass("uptoDescBS_90");
				} else if(energyValue<=75 && energyValue >70) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					
					$("#" + divIdImg).addClass("uptoDescBS_75");	
					
					$("#" + divIdImg).removeClass("uptoDescBS_9");
					$("#" + divIdImg).removeClass("uptoDescBS_10");					
					$("#" + divIdImg).removeClass("uptoDescBS_15");
					$("#" + divIdImg).removeClass("uptoDescBS_20");
					$("#" + divIdImg).removeClass("uptoDescBS_25");
					$("#" + divIdImg).removeClass("uptoDescBS_30");
					$("#" + divIdImg).removeClass("uptoDescBS_35");
					$("#" + divIdImg).removeClass("uptoDescBS_40");
					$("#" + divIdImg).removeClass("uptoDescBS_45");
					$("#" + divIdImg).removeClass("uptoDescBS_50");
					$("#" + divIdImg).removeClass("uptoDescBS_55");					
					$("#" + divIdImg).removeClass("uptoDescBS_60");
					$("#" + divIdImg).removeClass("uptoDescBS_65");
					$("#" + divIdImg).removeClass("uptoDescBS_70");
					$("#" + divIdImg).removeClass("uptoDescBS_80");
					$("#" + divIdImg).removeClass("uptoDescBS_90");
				} else if (energyValue<=80 && energyValue >75) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					
					$("#" + divIdImg).addClass("uptoDescBS_80");	
					
					$("#" + divIdImg).removeClass("uptoDescBS_9");
					$("#" + divIdImg).removeClass("uptoDescBS_10");					
					$("#" + divIdImg).removeClass("uptoDescBS_15");
					$("#" + divIdImg).removeClass("uptoDescBS_20");
					$("#" + divIdImg).removeClass("uptoDescBS_25");
					$("#" + divIdImg).removeClass("uptoDescBS_30");
					$("#" + divIdImg).removeClass("uptoDescBS_35");
					$("#" + divIdImg).removeClass("uptoDescBS_40");
					$("#" + divIdImg).removeClass("uptoDescBS_45");
					$("#" + divIdImg).removeClass("uptoDescBS_50");
					$("#" + divIdImg).removeClass("uptoDescBS_55");					
					$("#" + divIdImg).removeClass("uptoDescBS_60");
					$("#" + divIdImg).removeClass("uptoDescBS_65");
					$("#" + divIdImg).removeClass("uptoDescBS_70");
					$("#" + divIdImg).removeClass("uptoDescBS_75");
					$("#" + divIdImg).removeClass("uptoDescBS_90");
				} else if(energyValue >80) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					
					$("#" + divIdImg).addClass("uptoDescBS_90");	
					
					$("#" + divIdImg).removeClass("uptoDescBS_9");
					$("#" + divIdImg).removeClass("uptoDescBS_10");					
					$("#" + divIdImg).removeClass("uptoDescBS_15");
					$("#" + divIdImg).removeClass("uptoDescBS_20");
					$("#" + divIdImg).removeClass("uptoDescBS_25");
					$("#" + divIdImg).removeClass("uptoDescBS_30");
					$("#" + divIdImg).removeClass("uptoDescBS_35");
					$("#" + divIdImg).removeClass("uptoDescBS_40");
					$("#" + divIdImg).removeClass("uptoDescBS_45");
					$("#" + divIdImg).removeClass("uptoDescBS_50");
					$("#" + divIdImg).removeClass("uptoDescBS_55");					
					$("#" + divIdImg).removeClass("uptoDescBS_60");
					$("#" + divIdImg).removeClass("uptoDescBS_65");
					$("#" + divIdImg).removeClass("uptoDescBS_70");
					$("#" + divIdImg).removeClass("uptoDescBS_75");
					$("#" + divIdImg).removeClass("uptoDescBS_80");
				}	
				/********** ISSUE# 207 START *******/
				if(energyValue<=20 && energyValue >=15) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '-2px'});
				} else if (energyValue<=14 && energyValue >=5) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '-1px'});
				} else {
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '5px'});
				}
				/********** ISSUE# 207 END *******/	
			}
		}
	}	
	
	if(navigator.userAgent.match(/iPad/i)) { // IPAD
		for(var j=1; j<=divCountLoop;j++) {
			var style = document.getElementById("divImg_" +j).getAttribute("style");	
			if((style.search('block') != -1)){
				var textAreaPos = $("#areaId_"+j).position();
				if(textAreaPos.top != 5) {				
					$('#areaId_'+j).css('margin-top','-15px');
				} else {
					$('#areaId_'+j).css('margin-top','0px');
				}
			}
		}
	}
}


/** 
 * This fuction is call for after diagram only
 * This function add to fix the css 
 * for each task block on resize or add or delete any task 
 * @param null
 */
function fixDesignAS() {
	var energyValue = "";
	var divCountLoop = $('#outderDiv').children('.after_sketch_task').length;		
	for(var i=1; i<=divCountLoop;i++){
		var style = document.getElementById("divImg_" +i).getAttribute("style");
		if((style.search('block') != -1)){
			energyValue = document.getElementById("energyId_" +i).value;
			var divIdImg = document.getElementById("divImg_" +i).id;		
			if(navigator.userAgent.match(/Android/i) || navigator.userAgent.match(/iPhone/i)){ // Android and iPhone
				if(energyValue<=9 && energyValue >=5) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_small");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_middle");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_large");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_larger");					
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").addClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_smaller");	
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_smaller");
					
					$("#" + divIdImg).addClass("upto_9");					
				} else if (energyValue == 10) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_small");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_middle");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_large");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_larger");
					
					$("#" + divIdImg).removeClass("upto_9");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").addClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_smaller");
					
					$("#" + divIdImg).addClass("upto_10");
				} else if(energyValue<=15 && energyValue >10) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_middle");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_large");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_larger");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_9");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".energy").addClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_small");
					
					$("#" + divIdImg).addClass("upto_15");
				} else if(energyValue<=20 && energyValue >15) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_middle");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_large");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_larger");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".energy").addClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_small");
					$("#" + divIdImg).addClass("upto_20");
				} else if(energyValue<=25 && energyValue >20) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_middle");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_large");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_larger");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".energy").addClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_small");
					
					$("#" + divIdImg).addClass("upto_25");
				} else if(energyValue<=30 && energyValue >25) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");	
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_small");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_middle");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_large");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_larger");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_9");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");					
					
					$("#" + divIdImg).addClass("upto_30");					
				} else if(energyValue<=35 && energyValue >30) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");	
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_small");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_middle");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_large");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_larger");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_9");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");					
									
					$("#" + divIdImg).addClass("upto_35");
				} else if(energyValue<=40 && energyValue >35) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");	
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_small");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_middle");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_large");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_larger");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_9");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");		
					
					$("#" + divIdImg).addClass("upto_40");					
				} else if(energyValue<=45 && energyValue >40) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_small");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_large");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_larger");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_9");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_middle");
					
					$("#" + divIdImg).addClass("upto_45");
				} else if(energyValue<=50 && energyValue >45) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_small");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_large");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_larger");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_9");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_middle");
					
					$("#" + divIdImg).addClass("upto_50");
				} else if(energyValue<=55 && energyValue >50) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_small");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_large");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_larger");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_9");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_middle");
					
					$("#" + divIdImg).addClass("upto_55");
				} else if(energyValue<=60 && energyValue >55) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_small");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_large");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_larger");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_9");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_middle");
					
					$("#" + divIdImg).addClass("upto_60");
				} else if(energyValue<=65 && energyValue >60) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_small");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_middle");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_larger");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_9");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_large");
					
					$("#" + divIdImg).addClass("upto_65");
				} else if(energyValue<=70 && energyValue >65) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_small");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_middle");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_larger");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_9");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_large");
					
					$("#" + divIdImg).addClass("upto_70");
				} else if(energyValue<=75 && energyValue >70) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_small");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_middle");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_larger");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_9");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_large");
					
					$("#" + divIdImg).addClass("upto_75");
				} else if(energyValue<=80 && energyValue >75) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_small");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_middle");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_large");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_9");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_larger");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_larger");
					
					$("#" + divIdImg).addClass("upto_80");
				} else if(energyValue<=84 && energyValue >80) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_small");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_middle");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_large");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_9");
					$("#" + divIdImg).removeClass("upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_larger");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_larger");
					
					$("#" + divIdImg).addClass("upto_85");
				} else if(energyValue >84) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_small");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_middle");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_large");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_9");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_larger");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_larger");
					
					$("#" + divIdImg).addClass("upto_90");
				}
				/********** ISSUE# 207 START *******/
				/*if (energyValue<=10 && energyValue >=5) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0'});
				} else if(energyValue<=20 && energyValue >10) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0'});
				} else if(energyValue<=35 && energyValue >20) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0'});
				} else if (energyValue<=40 && energyValue >35) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0'});
				} else if (energyValue<=50 && energyValue >40) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0'});
				} else if (energyValue<=55 && energyValue >50) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0'});
				} else if (energyValue<=60 && energyValue >55) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0'});
				} else if (energyValue<=65 && energyValue >60) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0px'});
				} else {
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0px'});
				}*/
				$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0'});
				/********** ISSUE# 207 END *******/
			} 

			else if(navigator.userAgent.match(/iPad/i)){ // IPAD
				if(energyValue<=9 && energyValue >=5) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_small");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_middle");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_large");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_larger");					
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").addClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_smaller");	
									
					$("#" + divIdImg).addClass("upto_9");					
				} else if (energyValue == 10) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_small");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_middle");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_large");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_larger");
					
					$("#" + divIdImg).removeClass("upto_9");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").addClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_small");
					
					$("#" + divIdImg).addClass("upto_10");
				} else if(energyValue<=15 && energyValue >10) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_middle");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_large");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_larger");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_9");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".energy").addClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_small");
					
					$("#" + divIdImg).addClass("upto_15");
				} else if(energyValue<=20 && energyValue >15) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_middle");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_large");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_larger");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".energy").addClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_small");					
					$("#" + divIdImg).addClass("upto_20");
				} else if(energyValue<=25 && energyValue >20) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_middle");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_large");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_larger");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".energy").addClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".form-control-additional").addClass("form-control-additional_small");
					
					$("#" + divIdImg).addClass("upto_25");
				} else if(energyValue<=30 && energyValue >25) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");	
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_small");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_middle");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_large");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_larger");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_9");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");
										
					$("#" + divIdImg).addClass("upto_30");					
				} else if(energyValue<=35 && energyValue >30) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");	
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_small");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_middle");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_large");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_larger");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_9");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");					
										
					$("#" + divIdImg).addClass("upto_35");
				} else if(energyValue<=40 && energyValue >35) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");	
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_small");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_middle");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_large");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_larger");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_9");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");		
					
					$("#" + divIdImg).addClass("upto_40");					
				} else if(energyValue<=45 && energyValue >40) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_small");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_large");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_larger");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_9");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_middle");
					
					$("#" + divIdImg).addClass("upto_45");
				} else if(energyValue<=50 && energyValue >45) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_small");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_large");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_larger");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_9");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_middle");
					
					$("#" + divIdImg).addClass("upto_50");
				} else if(energyValue<=55 && energyValue >50) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_small");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_large");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_larger");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_9");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_middle");
					
					$("#" + divIdImg).addClass("upto_55");
				} else if(energyValue<=60 && energyValue >55) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_small");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_large");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_larger");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_9");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_middle");
					
					$("#" + divIdImg).addClass("upto_60");
				} else if(energyValue<=65 && energyValue >60) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_small");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_middle");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_larger");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_9");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_large");
					
					$("#" + divIdImg).addClass("upto_65");
				} else if(energyValue<=70 && energyValue >65) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_small");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_middle");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_larger");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_9");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_large");
					
					$("#" + divIdImg).addClass("upto_70");
				} else if(energyValue<=75 && energyValue >70) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_small");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_middle");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_larger");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_9");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_large");
					
					$("#" + divIdImg).addClass("upto_75");
				} else if(energyValue<=80 && energyValue >75) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_small");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_middle");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_large");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_9");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_larger");
					
					$("#" + divIdImg).addClass("upto_80");
				} else if(energyValue<=84 && energyValue >80) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_small");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_middle");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_large");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_9");
					$("#" + divIdImg).removeClass("upto_90");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_larger");
					
					$("#" + divIdImg).addClass("upto_85");
				} else if(energyValue >84) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_small");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_smaller");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_middle");
					$("#" + divIdImg).children().find(".form-control-additional").removeClass("form-control-additional_large");
					
					$("#" + divIdImg).removeClass("upto_10");
					$("#" + divIdImg).removeClass("upto_15");
					$("#" + divIdImg).removeClass("upto_20");
					$("#" + divIdImg).removeClass("upto_25");
					$("#" + divIdImg).removeClass("upto_30");
					$("#" + divIdImg).removeClass("upto_35");
					$("#" + divIdImg).removeClass("upto_40");
					$("#" + divIdImg).removeClass("upto_45");
					$("#" + divIdImg).removeClass("upto_50");
					$("#" + divIdImg).removeClass("upto_55");
					$("#" + divIdImg).removeClass("upto_60");					
					$("#" + divIdImg).removeClass("upto_65");
					$("#" + divIdImg).removeClass("upto_70");
					$("#" + divIdImg).removeClass("upto_75");
					$("#" + divIdImg).removeClass("upto_80");
					$("#" + divIdImg).removeClass("upto_85");
					$("#" + divIdImg).removeClass("upto_9");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_bigger");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_larger");
						
					$("#" + divIdImg).addClass("upto_90");
				}
				if (energyValue<=10 && energyValue >=5) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0px'});								
				} else if(energyValue<=15 && energyValue >10) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0px'});
				}  else if(energyValue<=20 && energyValue >15) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '-3px'});
				} else if(energyValue<=35 && energyValue >20) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '-1px'});
				} else if (energyValue<=40 && energyValue >35) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '-2px'});
				} else if (energyValue<=50 && energyValue >40) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0px'});
				} else if (energyValue<=55 && energyValue >50) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0px'});
				} else if (energyValue<=60 && energyValue >55) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0px'});
				} else if (energyValue<=65 && energyValue >60) {	
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0px'});
				} else {
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '0px'});
				}
				/********** ISSUE# 207 END *******/
			}else {  // desktop 
				if(energyValue<=9 && energyValue >=5) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").addClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_smaller");	
					
					$("#" + divIdImg).addClass("uptoDesc_9");
					
					$("#" + divIdImg).removeClass("uptoDesc_10");
					$("#" + divIdImg).removeClass("uptoDesc_15");
					$("#" + divIdImg).removeClass("uptoDesc_20");					
					$("#" + divIdImg).removeClass("uptoDesc_25");
					$("#" + divIdImg).removeClass("uptoDesc_30");
					$("#" + divIdImg).removeClass("uptoDesc_35");
					$("#" + divIdImg).removeClass("uptoDesc_40");
					$("#" + divIdImg).removeClass("uptoDesc_45");
					$("#" + divIdImg).removeClass("uptoDesc_50");
					$("#" + divIdImg).removeClass("uptoDesc_55");
					$("#" + divIdImg).removeClass("uptoDesc_60");					
					$("#" + divIdImg).removeClass("uptoDesc_65");
					$("#" + divIdImg).removeClass("uptoDesc_70");
					$("#" + divIdImg).removeClass("uptoDesc_75");
					$("#" + divIdImg).removeClass("uptoDesc_80");
					$("#" + divIdImg).removeClass("uptoDesc_90");
				} else if (energyValue == 10) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").addClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_small");
					
					$("#" + divIdImg).addClass("uptoDesc_10");
					
					$("#" + divIdImg).removeClass("uptoDesc_9");
					$("#" + divIdImg).removeClass("uptoDesc_15");
					$("#" + divIdImg).removeClass("uptoDesc_20");					
					$("#" + divIdImg).removeClass("uptoDesc_25");
					$("#" + divIdImg).removeClass("uptoDesc_30");
					$("#" + divIdImg).removeClass("uptoDesc_35");
					$("#" + divIdImg).removeClass("uptoDesc_40");
					$("#" + divIdImg).removeClass("uptoDesc_45");
					$("#" + divIdImg).removeClass("uptoDesc_50");
					$("#" + divIdImg).removeClass("uptoDesc_55");
					$("#" + divIdImg).removeClass("uptoDesc_60");					
					$("#" + divIdImg).removeClass("uptoDesc_65");
					$("#" + divIdImg).removeClass("uptoDesc_70");
					$("#" + divIdImg).removeClass("uptoDesc_75");
					$("#" + divIdImg).removeClass("uptoDesc_80");
					$("#" + divIdImg).removeClass("uptoDesc_90");
				} else if(energyValue<=15 && energyValue >10) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
								
					/******** Add Class********/
					$("#" + divIdImg).children().find(".energy").addClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_small");
					
					$("#" + divIdImg).addClass("uptoDesc_15");
										
					$("#" + divIdImg).removeClass("uptoDesc_9");
					$("#" + divIdImg).removeClass("uptoDesc_10");
					$("#" + divIdImg).removeClass("uptoDesc_20");					
					$("#" + divIdImg).removeClass("uptoDesc_25");
					$("#" + divIdImg).removeClass("uptoDesc_30");
					$("#" + divIdImg).removeClass("uptoDesc_35");
					$("#" + divIdImg).removeClass("uptoDesc_40");
					$("#" + divIdImg).removeClass("uptoDesc_45");
					$("#" + divIdImg).removeClass("uptoDesc_50");
					$("#" + divIdImg).removeClass("uptoDesc_55");
					$("#" + divIdImg).removeClass("uptoDesc_60");					
					$("#" + divIdImg).removeClass("uptoDesc_65");
					$("#" + divIdImg).removeClass("uptoDesc_70");
					$("#" + divIdImg).removeClass("uptoDesc_75");
					$("#" + divIdImg).removeClass("uptoDesc_80");
					$("#" + divIdImg).removeClass("uptoDesc_90");
				} else if(energyValue<=20 && energyValue >15) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
								
					/******** Add Class********/
					$("#" + divIdImg).children().find(".energy").addClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().addClass("form-control-sketch_small");
					
					$("#" + divIdImg).addClass("uptoDesc_20");
					
					$("#" + divIdImg).removeClass("uptoDesc_9");
					$("#" + divIdImg).removeClass("uptoDesc_10");					
					$("#" + divIdImg).removeClass("uptoDesc_15");					
					$("#" + divIdImg).removeClass("uptoDesc_25");
					$("#" + divIdImg).removeClass("uptoDesc_30");
					$("#" + divIdImg).removeClass("uptoDesc_35");
					$("#" + divIdImg).removeClass("uptoDesc_40");
					$("#" + divIdImg).removeClass("uptoDesc_45");
					$("#" + divIdImg).removeClass("uptoDesc_50");
					$("#" + divIdImg).removeClass("uptoDesc_55");
					$("#" + divIdImg).removeClass("uptoDesc_60");					
					$("#" + divIdImg).removeClass("uptoDesc_65");
					$("#" + divIdImg).removeClass("uptoDesc_70");
					$("#" + divIdImg).removeClass("uptoDesc_75");
					$("#" + divIdImg).removeClass("uptoDesc_80");
					$("#" + divIdImg).removeClass("uptoDesc_90");					
				} else if(energyValue<=25 && energyValue >20) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");										
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".energy").addClass("energy_small");
					
					$("#" + divIdImg).addClass("uptoDesc_25");
					
					$("#" + divIdImg).removeClass("uptoDesc_9");
					$("#" + divIdImg).removeClass("uptoDesc_10");					
					$("#" + divIdImg).removeClass("uptoDesc_15");
					$("#" + divIdImg).removeClass("uptoDesc_20");
					$("#" + divIdImg).removeClass("uptoDesc_30");
					$("#" + divIdImg).removeClass("uptoDesc_35");
					$("#" + divIdImg).removeClass("uptoDesc_40");
					$("#" + divIdImg).removeClass("uptoDesc_45");
					$("#" + divIdImg).removeClass("uptoDesc_50");
					$("#" + divIdImg).removeClass("uptoDesc_55");
					$("#" + divIdImg).removeClass("uptoDesc_60");					
					$("#" + divIdImg).removeClass("uptoDesc_65");
					$("#" + divIdImg).removeClass("uptoDesc_70");
					$("#" + divIdImg).removeClass("uptoDesc_75");
					$("#" + divIdImg).removeClass("uptoDesc_80");
					$("#" + divIdImg).removeClass("uptoDesc_90");				
				} else if(energyValue<=30 && energyValue >25) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");	
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).addClass("uptoDesc_30");
					
					$("#" + divIdImg).removeClass("uptoDesc_9");
					$("#" + divIdImg).removeClass("uptoDesc_10");					
					$("#" + divIdImg).removeClass("uptoDesc_15");
					$("#" + divIdImg).removeClass("uptoDesc_20");
					$("#" + divIdImg).removeClass("uptoDesc_25");					
					$("#" + divIdImg).removeClass("uptoDesc_35");
					$("#" + divIdImg).removeClass("uptoDesc_40");
					$("#" + divIdImg).removeClass("uptoDesc_45");
					$("#" + divIdImg).removeClass("uptoDesc_50");
					$("#" + divIdImg).removeClass("uptoDesc_55");
					$("#" + divIdImg).removeClass("uptoDesc_60");					
					$("#" + divIdImg).removeClass("uptoDesc_65");
					$("#" + divIdImg).removeClass("uptoDesc_70");
					$("#" + divIdImg).removeClass("uptoDesc_75");
					$("#" + divIdImg).removeClass("uptoDesc_80");
					$("#" + divIdImg).removeClass("uptoDesc_90");										
				} else if(energyValue<=35 && energyValue >30) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");	
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).addClass("uptoDesc_35");
					
					$("#" + divIdImg).removeClass("uptoDesc_9");
					$("#" + divIdImg).removeClass("uptoDesc_10");					
					$("#" + divIdImg).removeClass("uptoDesc_15");
					$("#" + divIdImg).removeClass("uptoDesc_20");
					$("#" + divIdImg).removeClass("uptoDesc_25");					
					$("#" + divIdImg).removeClass("uptoDesc_30");
					$("#" + divIdImg).removeClass("uptoDesc_40");
					$("#" + divIdImg).removeClass("uptoDesc_45");
					$("#" + divIdImg).removeClass("uptoDesc_50");
					$("#" + divIdImg).removeClass("uptoDesc_55");
					$("#" + divIdImg).removeClass("uptoDesc_60");					
					$("#" + divIdImg).removeClass("uptoDesc_65");
					$("#" + divIdImg).removeClass("uptoDesc_70");
					$("#" + divIdImg).removeClass("uptoDesc_75");
					$("#" + divIdImg).removeClass("uptoDesc_80");
					$("#" + divIdImg).removeClass("uptoDesc_90");										
				} else if(energyValue<=40 && energyValue >35) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_middle");
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".task").removeClass("task_middle");	
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");
					
					$("#" + divIdImg).addClass("uptoDesc_40");
					
					$("#" + divIdImg).removeClass("uptoDesc_9");
					$("#" + divIdImg).removeClass("uptoDesc_10");					
					$("#" + divIdImg).removeClass("uptoDesc_15");
					$("#" + divIdImg).removeClass("uptoDesc_20");
					$("#" + divIdImg).removeClass("uptoDesc_25");					
					$("#" + divIdImg).removeClass("uptoDesc_30");
					$("#" + divIdImg).removeClass("uptoDesc_35");
					$("#" + divIdImg).removeClass("uptoDesc_45");
					$("#" + divIdImg).removeClass("uptoDesc_50");
					$("#" + divIdImg).removeClass("uptoDesc_55");
					$("#" + divIdImg).removeClass("uptoDesc_60");					
					$("#" + divIdImg).removeClass("uptoDesc_65");
					$("#" + divIdImg).removeClass("uptoDesc_70");
					$("#" + divIdImg).removeClass("uptoDesc_75");
					$("#" + divIdImg).removeClass("uptoDesc_80");
					$("#" + divIdImg).removeClass("uptoDesc_90");									
				} else if(energyValue<=45 && energyValue >40) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");									
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					
					$("#" + divIdImg).addClass("uptoDesc_45");
					
					$("#" + divIdImg).removeClass("uptoDesc_9");
					$("#" + divIdImg).removeClass("uptoDesc_10");					
					$("#" + divIdImg).removeClass("uptoDesc_15");
					$("#" + divIdImg).removeClass("uptoDesc_20");
					$("#" + divIdImg).removeClass("uptoDesc_25");					
					$("#" + divIdImg).removeClass("uptoDesc_30");
					$("#" + divIdImg).removeClass("uptoDesc_35");
					$("#" + divIdImg).removeClass("uptoDesc_40");
					$("#" + divIdImg).removeClass("uptoDesc_50");
					$("#" + divIdImg).removeClass("uptoDesc_55");
					$("#" + divIdImg).removeClass("uptoDesc_60");					
					$("#" + divIdImg).removeClass("uptoDesc_65");
					$("#" + divIdImg).removeClass("uptoDesc_70");
					$("#" + divIdImg).removeClass("uptoDesc_75");
					$("#" + divIdImg).removeClass("uptoDesc_80");
					$("#" + divIdImg).removeClass("uptoDesc_90");
				} else if(energyValue<=50 && energyValue >45) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");									
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					
					$("#" + divIdImg).addClass("uptoDesc_50");
					
					$("#" + divIdImg).removeClass("uptoDesc_9");
					$("#" + divIdImg).removeClass("uptoDesc_10");					
					$("#" + divIdImg).removeClass("uptoDesc_15");
					$("#" + divIdImg).removeClass("uptoDesc_20");
					$("#" + divIdImg).removeClass("uptoDesc_25");					
					$("#" + divIdImg).removeClass("uptoDesc_30");
					$("#" + divIdImg).removeClass("uptoDesc_35");
					$("#" + divIdImg).removeClass("uptoDesc_40");
					$("#" + divIdImg).removeClass("uptoDesc_45");
					$("#" + divIdImg).removeClass("uptoDesc_55");
					$("#" + divIdImg).removeClass("uptoDesc_60");					
					$("#" + divIdImg).removeClass("uptoDesc_65");
					$("#" + divIdImg).removeClass("uptoDesc_70");
					$("#" + divIdImg).removeClass("uptoDesc_75");
					$("#" + divIdImg).removeClass("uptoDesc_80");
					$("#" + divIdImg).removeClass("uptoDesc_90");
				} else if(energyValue<=55 && energyValue >50) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");									
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					
					$("#" + divIdImg).addClass("uptoDesc_55");
					
					$("#" + divIdImg).removeClass("uptoDesc_9");
					$("#" + divIdImg).removeClass("uptoDesc_10");					
					$("#" + divIdImg).removeClass("uptoDesc_15");
					$("#" + divIdImg).removeClass("uptoDesc_20");
					$("#" + divIdImg).removeClass("uptoDesc_25");					
					$("#" + divIdImg).removeClass("uptoDesc_30");
					$("#" + divIdImg).removeClass("uptoDesc_35");
					$("#" + divIdImg).removeClass("uptoDesc_40");
					$("#" + divIdImg).removeClass("uptoDesc_45");
					$("#" + divIdImg).removeClass("uptoDesc_50");
					$("#" + divIdImg).removeClass("uptoDesc_60");					
					$("#" + divIdImg).removeClass("uptoDesc_65");
					$("#" + divIdImg).removeClass("uptoDesc_70");
					$("#" + divIdImg).removeClass("uptoDesc_75");
					$("#" + divIdImg).removeClass("uptoDesc_80");
					$("#" + divIdImg).removeClass("uptoDesc_90");
				} else if(energyValue<=60 && energyValue >55) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");									
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					
					$("#" + divIdImg).addClass("uptoDesc_60");
					
					$("#" + divIdImg).removeClass("uptoDesc_9");
					$("#" + divIdImg).removeClass("uptoDesc_10");					
					$("#" + divIdImg).removeClass("uptoDesc_15");
					$("#" + divIdImg).removeClass("uptoDesc_20");
					$("#" + divIdImg).removeClass("uptoDesc_25");					
					$("#" + divIdImg).removeClass("uptoDesc_30");
					$("#" + divIdImg).removeClass("uptoDesc_35");
					$("#" + divIdImg).removeClass("uptoDesc_40");
					$("#" + divIdImg).removeClass("uptoDesc_45");
					$("#" + divIdImg).removeClass("uptoDesc_50");
					$("#" + divIdImg).removeClass("uptoDesc_55");					
					$("#" + divIdImg).removeClass("uptoDesc_65");
					$("#" + divIdImg).removeClass("uptoDesc_70");
					$("#" + divIdImg).removeClass("uptoDesc_75");
					$("#" + divIdImg).removeClass("uptoDesc_80");
					$("#" + divIdImg).removeClass("uptoDesc_90");
				} else if(energyValue<=65 && energyValue >60) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					
					$("#" + divIdImg).addClass("uptoDesc_65");
					
					$("#" + divIdImg).removeClass("uptoDesc_9");
					$("#" + divIdImg).removeClass("uptoDesc_10");					
					$("#" + divIdImg).removeClass("uptoDesc_15");
					$("#" + divIdImg).removeClass("uptoDesc_20");
					$("#" + divIdImg).removeClass("uptoDesc_25");					
					$("#" + divIdImg).removeClass("uptoDesc_30");
					$("#" + divIdImg).removeClass("uptoDesc_35");
					$("#" + divIdImg).removeClass("uptoDesc_40");
					$("#" + divIdImg).removeClass("uptoDesc_45");
					$("#" + divIdImg).removeClass("uptoDesc_50");
					$("#" + divIdImg).removeClass("uptoDesc_55");					
					$("#" + divIdImg).removeClass("uptoDesc_60");
					$("#" + divIdImg).removeClass("uptoDesc_70");
					$("#" + divIdImg).removeClass("uptoDesc_75");
					$("#" + divIdImg).removeClass("uptoDesc_80");
					$("#" + divIdImg).removeClass("uptoDesc_90");
				} else if(energyValue<=70 && energyValue >65) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					
					$("#" + divIdImg).addClass("uptoDesc_70");
					
					$("#" + divIdImg).removeClass("uptoDesc_9");
					$("#" + divIdImg).removeClass("uptoDesc_10");					
					$("#" + divIdImg).removeClass("uptoDesc_15");
					$("#" + divIdImg).removeClass("uptoDesc_20");
					$("#" + divIdImg).removeClass("uptoDesc_25");					
					$("#" + divIdImg).removeClass("uptoDesc_30");
					$("#" + divIdImg).removeClass("uptoDesc_35");
					$("#" + divIdImg).removeClass("uptoDesc_40");
					$("#" + divIdImg).removeClass("uptoDesc_45");
					$("#" + divIdImg).removeClass("uptoDesc_50");
					$("#" + divIdImg).removeClass("uptoDesc_55");					
					$("#" + divIdImg).removeClass("uptoDesc_60");
					$("#" + divIdImg).removeClass("uptoDesc_65");
					$("#" + divIdImg).removeClass("uptoDesc_75");
					$("#" + divIdImg).removeClass("uptoDesc_80");
					$("#" + divIdImg).removeClass("uptoDesc_90");
				} else if(energyValue<=75 && energyValue >70) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_larger");					
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					
					$("#" + divIdImg).addClass("uptoDesc_75");
					
					$("#" + divIdImg).removeClass("uptoDesc_9");
					$("#" + divIdImg).removeClass("uptoDesc_10");					
					$("#" + divIdImg).removeClass("uptoDesc_15");
					$("#" + divIdImg).removeClass("uptoDesc_20");
					$("#" + divIdImg).removeClass("uptoDesc_25");					
					$("#" + divIdImg).removeClass("uptoDesc_30");
					$("#" + divIdImg).removeClass("uptoDesc_35");
					$("#" + divIdImg).removeClass("uptoDesc_40");
					$("#" + divIdImg).removeClass("uptoDesc_45");
					$("#" + divIdImg).removeClass("uptoDesc_50");
					$("#" + divIdImg).removeClass("uptoDesc_55");					
					$("#" + divIdImg).removeClass("uptoDesc_60");
					$("#" + divIdImg).removeClass("uptoDesc_65");
					$("#" + divIdImg).removeClass("uptoDesc_70");
					$("#" + divIdImg).removeClass("uptoDesc_80");
					$("#" + divIdImg).removeClass("uptoDesc_90");
				} else if (energyValue<=80 && energyValue >75) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");										
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					
					$("#" + divIdImg).addClass("uptoDesc_80");
					
					$("#" + divIdImg).removeClass("uptoDesc_9");
					$("#" + divIdImg).removeClass("uptoDesc_10");					
					$("#" + divIdImg).removeClass("uptoDesc_15");
					$("#" + divIdImg).removeClass("uptoDesc_20");
					$("#" + divIdImg).removeClass("uptoDesc_25");					
					$("#" + divIdImg).removeClass("uptoDesc_30");
					$("#" + divIdImg).removeClass("uptoDesc_35");
					$("#" + divIdImg).removeClass("uptoDesc_40");
					$("#" + divIdImg).removeClass("uptoDesc_45");
					$("#" + divIdImg).removeClass("uptoDesc_50");
					$("#" + divIdImg).removeClass("uptoDesc_55");					
					$("#" + divIdImg).removeClass("uptoDesc_60");
					$("#" + divIdImg).removeClass("uptoDesc_65");
					$("#" + divIdImg).removeClass("uptoDesc_70");
					$("#" + divIdImg).removeClass("uptoDesc_75");
					$("#" + divIdImg).removeClass("uptoDesc_90");
				} else if(energyValue >80) {
					/******** Remove Class********/
					$("#" + divIdImg).children().find(".task").removeClass("task_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_smaller");
					$("#" + divIdImg).children().find(".energy").removeClass("energy_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_smaller");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_small");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_middle");
					$("#" + divIdImg).children().find(".energy_field").children().removeClass("form-control-sketch_large");										
					
					/******** Add Class********/
					$("#" + divIdImg).children().find(".task").addClass("task_middle");
					$("#" + divIdImg).children().find(".energy").addClass("energy_middle");
					
					$("#" + divIdImg).addClass("uptoDesc_90");
					
					$("#" + divIdImg).removeClass("uptoDesc_9");
					$("#" + divIdImg).removeClass("uptoDesc_10");					
					$("#" + divIdImg).removeClass("uptoDesc_15");
					$("#" + divIdImg).removeClass("uptoDesc_20");
					$("#" + divIdImg).removeClass("uptoDesc_25");					
					$("#" + divIdImg).removeClass("uptoDesc_30");
					$("#" + divIdImg).removeClass("uptoDesc_35");
					$("#" + divIdImg).removeClass("uptoDesc_40");
					$("#" + divIdImg).removeClass("uptoDesc_45");
					$("#" + divIdImg).removeClass("uptoDesc_50");
					$("#" + divIdImg).removeClass("uptoDesc_55");					
					$("#" + divIdImg).removeClass("uptoDesc_60");
					$("#" + divIdImg).removeClass("uptoDesc_65");
					$("#" + divIdImg).removeClass("uptoDesc_70");
					$("#" + divIdImg).removeClass("uptoDesc_75");
					$("#" + divIdImg).removeClass("uptoDesc_80");
				}			
				/********** ISSUE# 207 START *******/
				if(energyValue<=20 && energyValue >=15) {
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '-1px'});
				} else if (energyValue<=14 && energyValue >=10) {
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '-1px'});
				} else if (energyValue<=9 && energyValue >=5) {
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '-1px'});
				} else {
					$("#" + divIdImg).children(".ui-resizable-handle").css({'bottom' : '1px'});
				}
				/************ Added for Additional Areas *************/                     				
				$("#" + divIdImg).children().find(".form-control-additional").css({'width' : '77%'});

			}		
		}		
	}
}