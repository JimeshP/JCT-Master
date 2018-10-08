<!DOCTYPE html>
<%@taglib uri="http://www.springframework.org/tags" prefix="spring"%>
<html>
<head>
    <title><spring:message code="label.landing.page.title"/></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link rel="shortcut icon" href="/user/img/crafting_ico.ico" />
    <!-- Bootstrap -->
    <link href="../css/bootstrap.css" rel="stylesheet">
    <link href="../css/style.css" rel="stylesheet">
    <link href="https://s3-us-west-2.amazonaws.com/jobcrafting/css/commonStyle.css" rel="stylesheet">
    <link rel="stylesheet" href="../css/alertify.core.css" />
	<link rel="stylesheet" href="../css/alertify.default.css" id="toggleCSS" />
	<script src="../js/alertify.min.js"></script>
    <script type="text/javascript" src="../lib/jquery.js"></script>
    <script type="text/javascript" src="../lib/spine.js"></script>
	<script type="text/javascript" src="../lib/ajax.js"></script>
	
    <!-- HTML5 Shim and Respond.js IE8 support of HTML5 elements and media queries -->
    <!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
      <script src="https://oss.maxcdn.com/libs/respond.js/1.3.0/respond.min.js"></script>
    <![endif]-->

    </head>
<body>
    <div class="main_warp">

      <!-- Header area start -->
        <!-- <div class="row-fluid header">
            <div class="header_wrap_area container" >
            <div class="col-md-3">
            <h5 class="heading_main"><img src="../img/logo.png" alt=""></h5>
            </div>
              <div class="col-md-9">
              <h5 class="heading_main">Header Section</h5>
              </div>              
            </div>
        </div> -->
        <!-- Header area start -->
        	<jsp:include page="headerLanding.jsp"/>
     	<!-- Header area end -->
       

        <!-- Form area start -->
        
        <div class="container-fluid">  
           <div class="container" >
           <div class="row" >
             <div class="cont-middle">

               <div class="video_testimo_area">
                  <div class="video-section">
                    <h3 class="page-title"><spring:message code="label.watch.video"/></h3>
                    <video width="75%" poster="../img/frame.jpg" controls preload="auto">
					    <!-- <source src="../../JCT_VIDEO/JobCrafting_MP4.mp4" type="video/mp4">
					    <source src="../../JCT_VIDEO/JobCrafting_WEBM.webm" type="video/webm">
					    <source src="../../JCT_VIDEO/JobCrafting_OGV.ogv" type="video/ogv">
					    <object data="../../JCT_VIDEO/JobCrafting_SWF.swf">
    						<embed src="../../JCT_VIDEO/JobCrafting_SWF.swf">
 						</object> -->
 						<source src="https://s3-us-west-2.amazonaws.com/jobcrafting/video/JobCrafting_MP4.mp4" type="video/mp4">
					    <source src="https://s3-us-west-2.amazonaws.com/jobcrafting/video/JobCrafting_WEBM.webm" type="video/webm">
					    <source src="https://s3-us-west-2.amazonaws.com/jobcrafting/video/JobCrafting_OGV.ogv" type="video/ogv">
					    <object data="https://s3-us-west-2.amazonaws.com/jobcrafting/video/JobCrafting_SWF.swf">
    						<embed src="https://s3-us-west-2.amazonaws.com/jobcrafting/video/JobCrafting_SWF.swf">
 						</object>
					</video> 
                  </div>
                  
                  <%-- <div class="testimonial">
                    <h3 class="page-title"><spring:message code="label.testimonials"/></h3>
                    <ul class="bxslider">
                      <li>
                        <blockquote>Eye-opening. 
                        <p style="text-align:right;margin-right:20px;">- Management Analyst</p>  
                        </blockquote>
                      </li>
                      <li>
                        <blockquote>Rejuvenating.
                        <p style="text-align:right;margin-right:20px;">- Product Engineer</p>
                        </blockquote>
                      </li>
                      <li>
                        <blockquote>Reveals a truckload of information and insight.
                        <p style="text-align:right;margin-right:20px;">- Project Manager</p>
                        </blockquote>
                      </li>
                      <li>
                        <blockquote>Enjoyable, informational, and inspiring.
                        <p style="text-align:right;margin-right:20px;">- Retirement Consultant</p>
                        </blockquote>
                      </li>
                      <li>
                        <blockquote>Very useful and systematic.
                        <p style="text-align:right;margin-right:20px;">- Marketing Coordinator</p>
                        </blockquote>
                      </li>
                      <li>
                        <blockquote>Inspired real, lasting change in my work life.
                        <p style="text-align:right;margin-right:20px;">- Administrative Assistant</p>
                        </blockquote>
                      </li>
                      <li>
                        <blockquote>Ingenious...helped me see how to improve my job in ways I would have never thought of otherwise.
                        <p style="text-align:right;margin-right:20px;">- Software Engineer</p>
                        </blockquote>
                      </li>
                      <li>
                        <blockquote>Matters to what I'll do tomorrow, not ten years and two jobs later.
                        <p style="text-align:right;margin-right:20px;">- Marketing Analyst</p>
                        </blockquote>
                      </li>
                      <li>
                        <blockquote>Enlightening...made me realize what I liked and disliked about my job in a clearer way than I thought possible.
                        <p style="text-align:right;margin-right:20px;">- HR Manager</p>
                        </blockquote>
                      </li>
                    </ul>
                  </div> --%>
               </div>
               <div class="clearfix"></div>
               
               <div class="landing_page_note"><spring:message code="label.landing.page.note"/></div>
               
               <div class="start-btn">
                 <a class="multi-line-button green" href="#" style="width:14em" id="proceed" onClick="start()">
                  <span class="title"><spring:message code="label.start.now"/></span>
                 </a>
               </div>
             </div>
             </div>
           </div>
        </div>
         <!-- Form area end -->
        <!-- Footer area start -->
        <div class="row-fluid footer footer-wrapper footer_landing">
          <div class="container">
            <div class="copyright">
            	<spring:message code="label.footer"/><br>
            	<spring:message code="label.footer.permission"/>
            </div>
          </div>
        </div>
        <!-- Footer area end -->
    </div>

<!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->

<script src="https://code.jquery.com/jquery.js"></script>    
<script src="../js/bootstrap.min.js"></script>


<script src="../js/jquery.bxslider.min.js" type="text/javascript"></script>
<link href="../css/jquery.bxslider.css" rel="stylesheet" type="text/css" />
<!-- <link href="../css/bootstrap.min.css" rel="stylesheet" type="text/css" /> -->
<script src="../js/common.js"></script>
<script type="text/javascript">

		/**
		 * Function add to disable browser back button
		 * while the page is loaded
		 * @param null
		 */
		window.location.hash="";
		window.location.hash="Again-No-back-button";//again because google chrome don't insert first hash into history
		window.onhashchange=function(){window.location.hash="";};


         $(document).ready(function () {           
             $('.bxslider').bxSlider({
                 mode: 'vertical',
                 slideMargin: 3,
                 auto:true
             });             
         });
    </script>
           


 
</body>

</html>