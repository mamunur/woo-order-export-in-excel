<?php
//********************************************
//	Custom meta boxes
//***********************************************************
if(!function_exists("add_custom_boxes")){
	function add_custom_boxes(){
		add_meta_box( "page_options", __("Save order in Excel", "educare"), "page_options", "shop_order", "side", "core", null );
		
	}
}

function page_options(){
	global $post;
	?>
    <p><b><?php _e("Order Excel Generating", "educare"); ?></b></p>
    
    <button class="generate_download_excel button button-primary" data-uploader-title="<?php _e("Select a header image", "educare"); ?>" data-uploader-button-text="<?php _e("Select Image", "educare"); ?>" name="generate_excel" data-shop-id="<?php echo $post->ID; ?>" >Generate & Download Excel</button>
    <input type="hidden" class="header_image_input" name="header_image" value="" >

<?php
}
add_action( 'add_meta_boxes', 'add_custom_boxes' );


add_action( 'manage_shop_order_posts_custom_column' , 'custom_orders_list_column_content', 10, 2 );
function custom_orders_list_column_content($column_name, $post_id)
{
    
   global $post, $woocommerce, $the_order;

    if($column_name=="order_actions"){
      echo "<a class='button tips view generate_download_excel' data-shop-id=".$post_id." >".__( 'Generate Excel', 'woocommerce' )."</a>";

    }
}
// Set Here the WooCommerce icon for your action button
add_action( 'admin_head', 'add_custom_order_status_actions_button_css' );
function add_custom_order_status_actions_button_css() {
    echo '<style>.view.generate_download_excel::after { font-family: woocommerce; content: "\e005" !important; }</style>';
}