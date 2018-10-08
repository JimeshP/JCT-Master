$(function(){
	
	$(".loader_bg").fadeIn();
	$(".loader").fadeIn();
	var navigation = Spine.Model.sub();
	navigation.configure("/user/search/fetchOldPdf", "jctUserId");
	navigation.extend( Spine.Model.Ajax );	
	//Populate the model
	var replyModel = new navigation({
		jctUserId: sessionStorage.getItem("jctUserId")
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
		console.log(obj.statusMsg);
		if(obj.statusMsg == "Success"){
			plotDiagrams(obj.dto, obj.statuscode);			
		} else {
			alertify.alert("<img src='../img/alert-icon.png'><br /><p>Unable to fetch your history, please try again later.</p>");
			return false;		
		}
		$(".loader_bg").fadeOut();
	    $(".loader").hide();
	});
});



function plotDiagrams(list, statusCode) {
	var plottingDiv = document.getElementById("plotPdfDiv");
	if (statusCode == 200) {	
		var trColor = "";
		var headingString = "<table width='94%' id='refQtnTableId' border=1 align='center' bordercolor='#FAAC58' class='tablesorter' style='margin-top: 5px;'><thead class='tab_header'><tr><th width='10%' ><b>#</b></th><th width='41%'><b>Date Created</b></th><th width='11%' class='header_disable'><b>View</b></th></tr></thead><tbody>";
		for (var index = 0; index < list.length; index++) {
			if(index % 2 == 0) {
				trColor = "#F5D0A9";
			} else {
				trColor = "#F8ECE0";
			}
			createdTs = new Date((list[index].jctCreatedTimestamp)).toDateString();
			formattedDate = dateformat(new Date (createdTs));
			//var pdfName = "Results PDF#"+(index + 1)+ "-" +formattedDate;
			var fnName="showPdf(\""+list[index].jctFileName+"\")";
			headingString = headingString + "<tr align='center' class='row_width' bgcolor='"+trColor+"'><td align='center'>"+(parseInt(index)+1)+".</td><td>"+formattedDate+"</td><td style='text-align: center;'><span class='pdf_view' onclick="+fnName+">Download PDF</span></td></tr>";
		}
	
		headingString = headingString + "</tbody></table>";
		plottingDiv.innerHTML = headingString;
		$(function(){
			  $("table").tablesorter({
			    headers: {
			      2: { sorter: false }      // disable first column			     
			    }
			  });
			});
	} else if (statusCode == 201) {
		plottingDiv.innerHTML = "<div align='center'><img src='../img/no-record.png'><br /><div class='textStyleNoExist' align='center'>User has not shared any diagram.</div></div>";
	} else {
		plottingDiv.innerHTML = "<div align='center'>Server encountered error.</div>";
	}
}
function showPdf(jctFileName){		
	var link = document.createElement("a");
	var pdfLink = "../../user/search/showIndividualPdf/";
	pdfLink = pdfLink + jctFileName.substring(jctFileName.lastIndexOf("/")) ;
	link.setAttribute("href", pdfLink);
	window.open(pdfLink, "", "width=700, height=600");
}
/**
 * This function gives the date in format
 * @param inputdate
 * @returns formatedDate
 */
function dateformat(inputdate) {
	var dateStr = inputdate.toDateString();	
	//var newchar = '-';
	//formatedDate = formatedDate.split(' ').join(newchar);
	return dateStr;
}