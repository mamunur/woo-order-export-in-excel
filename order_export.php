<?php
/*
Plugin Name: Woo Order Export-in-excel
Plugin URI: http://ciphercoin.com/
Description: This generates and export woocommerce single order into Excel format from order column and order details.
Author: MD. Mamunur Rahman
Author URI: http://ciphercoin.com/
Text Domain: welldone
Version: 1.0.2
*/
register_activation_hook( __FILE__, 'oe_plugin_activate' );
function oe_plugin_activate(){
   add_option('product_id','What on');
}
register_deactivation_hook( __FILE__, 'oe_plugin_deactivate' );
function oe_plugin_deactivate(){

}

add_action( 'admin_menu', 'neod_plugin_admin_menu' );
function oe_plugin_admin_menu() {
    add_menu_page(
        'Order Export',
        'Order Export',
        'manage_options',
        'order-export',
        'order_export_fuction', // Callback, leave empty
        'dashicons-calendar',
        6 // Position
    );
}

add_action( 'admin_menu', 'oe_plugin_admin_menu' );
function order_export_fuction(){
   $order = wc_get_order( 2869 );  
//echo "<pre>".print_r($data, 1)."</pre>";
echo '<table border="1">';
echo '<tr><td>Product</td><td>Color</td><td>Stain Top</td><td>Customer Name</td><td>Customer Msg</td><td>Table Size</td><td>Customize Depth</td></tr>';
$order_data = $order->get_data(); // The Order data

   //echo "<pre>".print_r($order_data , 1)."</pre>";
   $order_no = $order_data['number'];
   echo $order_no;
  foreach ($order->get_items() as $item_key => $item_values):
    $item_data = $item_values->get_data();
    $color =$stain_top = $table_size = $customer_name = $customer_msg = $customize_depth = '';
    $product_name = $item_data['name'];
    $product_id = $item_data['product_id'];
    $product_cats_ids = wc_get_product_term_ids( $product_id, 'product_cat' );
    foreach( $product_cats_ids as $cat_id ) {
      $term = get_term_by( 'id', $cat_id, 'product_cat' );
      echo $term->name."</br>";
    }
    $variation_id = $item_data['variation_id'];
    $quantity = $item_data['quantity'];
    $tax_class = $item_data['tax_class'];
    $line_subtotal = $item_data['subtotal'];
    $line_subtotal_tax = $item_data['subtotal_tax'];
    $line_total = $item_data['total'];
    $line_total_tax = $item_data['total_tax'];
    $wc_product = $item_values->get_product();
    $table_size = $wc_product->get_description();
   //echo $product_name." ".$quantity." ".$line_subtotal;
    $data = wc_get_order_item_meta( $item_key, '_tmcartepo_data' );
    foreach ($data as $value) {
        if($value['name'] == 'Color') $color = $value['value'];
        if($value['name'] == 'Stain Top') $stain_top = $value['value'];
        if($value['name'] == 'Customer Name') $customer_name = $value['value'];
        if($value['name'] == 'Notes / Customization') $customer_msg = $value['value'];
        if($value['name'] == 'Table Size') $table_size = $value['value'];
        if($value['name'] == 'Customize Depth') $customize_depth = $value['value'];   
    }
    

    //echo "<pre>".print_r($item_data, 1)."</pre>";
    
   // echo "<pre>".print_r($wc_product , 1)."</pre>";
    
   // echo '<tr><td>'.$product_name.'</td><td>'.$color.'</td><td>'.$stain_top.'</td><td>'.$customer_name.'</td><td>'.$customer_msg.'</td><td>'.$table_size.'</td><td>'.$customize_depth.'</td></tr>';
    

endforeach;
echo '</table>';
$dir = plugin_dir_path( __FILE__ );
echo $dir;
}
add_action('plugins_loaded','oe_load_class_files');
function oe_load_class_files(){
  if ( in_array( 'woocommerce/woocommerce.php', apply_filters( 'active_plugins', get_option( 'active_plugins' ) ) ) ) {
	require_once 'classes/class.export.php';
	require_once 'meta_boxes.php';
  }
}
function oe_admin_scripts($hook_suffix){
  wp_enqueue_script( 'oe_admin', plugins_url( '/js/admin.js', __FILE__ ), array('jquery'), '1.0', true );
	wp_localize_script( 'oe_admin', 'ajax_object', array(
		'ajaxurl' => admin_url( 'admin-ajax.php' )
	));	
		
}
add_action( 'admin_enqueue_scripts', 'oe_admin_scripts' );

add_action( 'wp_ajax_nopriv_order_export_in_excel', 'order_export_in_excel' );
add_action( 'wp_ajax_order_export_in_excel', 'order_export_in_excel' );

function order_export_in_excel() {
	$order_id = $_POST['post_id'];
   require_once( '../wp-load.php' );
  
   $file_url = oe_export::get_order_xls($order_id);

		
	  echo $file_url;
	die();
}
?>