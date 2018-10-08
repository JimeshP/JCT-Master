/*This goes in the HEAD of the html file */

var sec = 0;
var min = 0;
var hour = 0;
var theResult = "";

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

/*  if (text == "Start") { document.clock.theButton.value = "Stop "; }
  if (text == "Stop ") { document.clock.theButton.value = "Start"; sec = sec-1; sec--; }

  if (document.clock.theButton.value == "Start") {
   return true; }*/
SD=window.setTimeout("stopwatch();", 1000);

 theResult = document.clock.stwa.value;
}

window.onload = stopwatch;

function saveIt() {
 if (document.clock.saver.value == "Save Time 1") {
   document.clock.saver.value = "Save Time 2";
 document.getElementById('time01').innerHTML = "<b>Time 1 =</b> " + theResult + "<br />";
 return;
 }
 if (document.clock.saver.value == "Save Time 2") {
   document.clock.saver.value = "Save Time 3";
 document.getElementById('time02').innerHTML = "<b>Time 2 =</b> " + theResult + "<br />";
 return;
 }
 if (document.clock.saver.value == "Save Time 3") {
   document.clock.saver.value = "Reset Times!";
 document.getElementById('time03').innerHTML = "<b>Time 3 =</b> " + theResult;
 return;
 }
 if (document.clock.saver.value == "Reset Times!") {
  document.getElementById('time01').innerHTML = "";
  document.getElementById('time02').innerHTML = "";
  document.getElementById('time03').innerHTML = "";
  document.clock.saver.value = "Save Time 1";
 }
}
