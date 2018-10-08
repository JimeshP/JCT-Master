<!DOCTYPE html>
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=Edge">
<meta charset="utf-8">

<title>bootstrap-wysihtml5</title>


<link rel="stylesheet" type="text/css" href="../css/bootstrap.min.css"></link>
<link rel="stylesheet" type="text/css" href="../css/prettify.css"></link>
<link rel="stylesheet" type="text/css" href="../css/bootstrap-wysihtml5.css"></link>

<style type="text/css" media="screen">
	.btn.jumbo {
		font-size: 20px;
		font-weight: normal;
		padding: 14px 24px;
		margin-right: 10px;
		-webkit-border-radius: 6px;
		-moz-border-radius: 6px;
		border-radius: 6px;
	}
</style>

<script>
function test(){
 var htmlTxt = $('#myTextAreaId').val();
 alert(htmlTxt);
 
 document.getElementById('testDiv').innerHTML = htmlTxt;
}
</script>
</head>
<body>
<div class="main-cont-wrapper">
	<div class="hero-unit" style="margin-top:40px">
		<hr/>
		<textarea class="textarea" id='myTextAreaId' placeholder="Enter text ..." style="width: 810px; height: 200px"></textarea>
	</div>
	
	<div class="row">
		<div class="span6">
	        
			
			
		</div>
	</div>
</div>
<input type="button" onClick="test()" value="my html" />
<div id='testDiv'>

</div>

<script src="../lib/wysihtml5-0.3.0.js"></script>
<script src="../lib/jquery-1.7.2.min.js"></script>
<script src="../lib/prettify.js"></script>
<script src="../lib/bootstrap.min.js"></script>
<script src="../lib/bootstrap-wysihtml5.js"></script>

<script>
	//$('.textarea').wysihtml5();
	$('.textarea').wysihtml5({
	"font-styles": false, //Font styling, e.g. h1, h2, etc. Default true
	"emphasis": true, //Italics, bold, etc. Default true
	"lists": true, //(Un)ordered lists, e.g. Bullets, Numbers. Default true
	"html": false, //Button which allows you to edit the generated HTML. Default false
	"link": false, //Button to insert a link. Default true
	"image": false, //Button to insert an image. Default true,
	"color": false //Button to change color of font  
});
</script>

</body>
</html>
