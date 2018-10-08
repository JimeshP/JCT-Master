function equalHeight(group) {
   tallest = 0;
   mmh = 15;
   group.each(function() {
      thisHeight = $(this).height();
      //alert(thisHeight);
      if(thisHeight > tallest) {
         tallest = thisHeight;
      }
   });
   group.height(tallest);
}
$(document).ready(function() {
   equalHeight($(".col"));
});