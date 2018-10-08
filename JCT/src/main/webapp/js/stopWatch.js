/*This goes in the HEAD of the html file */

var sec = 0;
var min = 0;
var hour = 0;
var theResult = "";
var totalTime = "";
var sec2 = 0;
var min2 = 0;
var hour2 = 0;
var sessionTime = "";
var sec3 = 0;
var min3 = 0;
var hour3 = 0;
function stopwatch() {
   sec++;
  if (sec == 60) {
   sec = 0;
   min = min + 1; }
  else {
   min = min; }
  if (min == 60) {
   min = 0; 
   hour += 1; }

if (sec<=9) { sec = "0" + sec; }
   document.clock.stwa.value = ((hour<=9) ? "0"+hour : hour) + " : " + ((min<=9) ? "0" + min : min) + " : " + sec;
   SD=window.setTimeout("stopwatch();", 1000);
   theResult = document.clock.stwa.value;
   
   //document.clock.stwa2.value = ((hour<=9) ? "0"+hour : hour) + " : " + ((min<=9) ? "0" + min : min) + " : " + sec;
   //theResult = document.clock.stwa2.value;
   updateTotalTime();
}

function updateSessionTime() {
	if (!sessionStorage.getItem("sessionTime")) {
		sessionTime = "00:00:00";
	} else {
		sessionTime = sessionStorage.getItem("sessionTime");
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
	/*var totalTime3 = hour3+" : "+min3+" : "+sec3;
	sessionStorage.setItem("sessionTime", totalTime3);
	//document.getElementById("timeId").innerHTML = totalTime2;
	document.clock.stwa2.value = totalTime3;*/
	
	
	
	var totalTimeDisp = hour3+" Hrs "+min3+" Mins ";
	var totalTime3 = hour3+" : "+min3+" : "+sec3;
	sessionStorage.setItem("sessionTime", totalTime3);
	document.clock.stwa2.value = totalTimeDisp;
	//document.getElementById("stwa2").innerHTML = totalTimeDisp;
	//console.log("totalTimeDisp"+totalTimeDisp);
	//console.log("totalTime3"+totalTime3);
}

function updateTotalTime() {
	updateSessionTime();
	totalTime = sessionStorage.getItem("totalTime");
	sessionStorage.setItem("started", "Y");
	var totalTimeArr = totalTime.split(":");
	var sec2 = parseInt(totalTimeArr[2]);
	sec2 = parseInt(sec2) + 1; 
	min2 = parseInt(totalTimeArr[1]);
	hour2 = parseInt(totalTimeArr[0]);
	if (parseInt(sec2) == 60) {
		sec2 = 0;
		min2 = parseInt(min2) + 1;
	} else {
		min2 = min2; 
	} if (parseInt(min2) == 60) {
		   min2 = 0; 
		   hour2 = parseInt(hour2) + 1; 
	}
	if (sec2<=9) { sec2 = "0" + sec2; }
	hour2 = ""+hour2;
	min2 = ""+min2;
	//console.log(hour2.length);
	if (""+hour2.length < 2) {
		hour2 = "0"+hour2;
	}
	if (""+min2.length < 2) {
		min2 = "0"+min2;
	}
	/*var totalTime2 = hour2+" : "+min2+" : "+sec2;
	sessionStorage.setItem("totalTime", totalTime2);
	document.getElementById("timeId").innerHTML = totalTime2;*/
	
	var totalTimeDisp2 = hour2+" Hrs "+min2+" Mins ";
	var totalTime2 = hour2+" : "+min2+" : "+sec2;
	sessionStorage.setItem("totalTime", totalTime2);
	document.getElementById("timeId").innerHTML = totalTimeDisp2;
	
}

window.onload = stopwatch;

