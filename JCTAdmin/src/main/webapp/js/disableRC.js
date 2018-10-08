////////// F12 & F5 disable code////////////////////////


$(document).on('keydown keyup', function(e) {	//	newly added_ARINDAM
    if(e.which === 116) {
       //console.log('Refresh blocked');
       return false;
    }
    if(e.which === 82 && e.ctrlKey) {
       //console.log('Refresh blocked');
       return false;
    }
    if(e.which === 123) {
        //console.log('F12 blocked');
        return false;
     }
});

//Disable right click script 
//visit http://www.rainbow.arch.scriptmania.com/scripts/ 
var message="Sorry, right-click has been disabled"; 
/////////////////////////////////// 
function clickIE() {if (document.all) {(message);return false;}} 
function clickNS(e) {if 
(document.layers||(document.getElementById&&!document.all)) { 
if (e.which==2||e.which==3) {(message);return false;}}} 
if (document.layers) 
{document.captureEvents(Event.MOUSEDOWN);document.onmousedown=clickNS;} 
else{document.onmouseup=clickNS;document.oncontextmenu=clickIE;} 
document.oncontextmenu=new Function("return false") 
// -->