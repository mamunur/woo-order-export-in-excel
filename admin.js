(function($) {
  "use strict";

 jQuery(document).ready( function($){

	$(document).on("click", ".generate_download_excel", function(e){
		e.preventDefault();
//$preparingFileModal.dialog({ modal: true });
       var shop_order_id = $(this).attr('data-shop-id');
	  //alert(shop_id);
	 
		jQuery.ajax({
			url : ajax_object.ajaxurl,
			type : 'post',
			data : {
				action : 'order_export_in_excel',
				post_id : shop_order_id
			},
			success : function( response ) {
				//alert(response);
				window.location.href=response;
			}
		 });
	});
	
	
});
})(jQuery);