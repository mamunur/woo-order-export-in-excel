<?php
/**
 * teachPress export class
 *
 * @since 3.0.0
 */
class oe_export {
    /**
     * Export course data in xls format
     * @param int $course_ID 
     */
    static function get_order_xls($order_id) {
        $dir = plugin_dir_path( __DIR__ );
        $url = plugin_dir_url( __DIR__ );
        // load course data
        $filename = 'order_'. date('Y-m-d-His').'_'.$order_id.'.xlsx';
		$filepath = $dir .'excel/'. $filename;
        $file_url  = stripslashes( trim( $url .'excel/'. $filename ) );
    
include 'PHPExcel.php';
include 'PHPExcel/Writer/Excel2007.php';

$objPHPExcel = new PHPExcel();
$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->getProperties()->setCreator("Me")->setLastModifiedBy("Me")->setTitle("My Excel Sheet")->setSubject("My Excel Sheet")->setDescription("Excel Sheet")->setKeywords("Excel Sheet")->setCategory("Me");
$order = wc_get_order( $order_id );

$order_data = $order->get_data(); // The Order data
$orderid = $order_data['number'];
$order_date_created = $order_data['date_created']->date('Y-m-d');
$customer_email = $order_data['billing']['email'];

$total = $order_data['total'];
// Add column headers
$objPHPExcel->getActiveSheet()
			->setCellValue('B1', 'Order No:')
			->setCellValue('C1', 'Order Date:')
            ->getStyle('B1:C1')->getFont()->setBold(true)
			;
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setAutoSize(false)->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setAutoSize(false)->setWidth(25);
$objPHPExcel->getActiveSheet()
			->setCellValue('A4', 'Category:')
            ->setCellValue('B4', 'Color:')
			->setCellValue('C4', 'Stained Top:')
			->setCellValue('D4', 'Qty:')
			->setCellValue('E4', 'Product Name:')
			->setCellValue('F4', 'Customize Depth:')
			->setCellValue('G4', 'Customer Message:')
			->setCellValue('H4', 'Customer Name:')
			->setCellValue('I4', 'Price:')
			->setCellValue('J4', 'Line Total:')
            ->getStyle('A4:J4')->getFont()->setBold(true)
            ;
//Put each record in a new cell
	$objPHPExcel->getActiveSheet()->setCellValue('B2', $orderid)->getStyle('B2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
	$objPHPExcel->getActiveSheet()->setCellValue('C2', $order_date_created);
	//$objPHPExcel->getActiveSheet()->setCellValue('D4', 'Qty:')->getStyle('D40')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$i = 4;
foreach ($order->get_items() as $item_key => $item_values):
    $item_data = $item_values->get_data();
    $i++;
    $color =$stain_top = $table_size = $customer_name = $customer_msg = $customize_depth = $price = $drawing = $category = '';
    $product_name = $item_data['name'];
    $product_id = $item_data['product_id'];
    $variation_id = $item_data['variation_id'];
    $product_cats_ids = wc_get_product_term_ids( $product_id, 'product_cat' );
    foreach( $product_cats_ids as $cat_id ) {
      $term = get_term_by( 'id', $cat_id, 'product_cat' );
      $category .= $term->name.", ";
    }
    $category = htmlspecialchars_decode($category);
    $category = rtrim($category,', ');
    $quantity = $item_data['quantity'];
    $tax_class = $item_data['tax_class'];
    $line_subtotal = $item_data['subtotal'];
    $price = wc_format_decimal($line_subtotal / $quantity);
    $line_subtotal_tax = $item_data['subtotal_tax'];
    $line_total = $item_data['total'];
    $line_total_tax = $item_data['total_tax'];
    $wc_product = $item_values->get_product();
    $table_size = html_entity_decode($wc_product->get_description());
   //echo $product_name." ".$quantity." ".$line_subtotal;
    $data = wc_get_order_item_meta( $item_key, '_tmcartepo_data' );
    foreach ($data as $value) {
        if($value['name'] == 'Color') $color = $value['value'];
        if($value['name'] == 'Stain Top') $stain_top = $value['value'];
        if($value['name'] == 'Customer Name') $customer_name = $value['value'];
        if($value['name'] == 'Notes / Customization') $customer_msg = $value['value']."\n ";
        if($value['name'] == 'Table Size') $table_size = $value['value'];
        if($value['name'] == 'Customize Depth') $customize_depth = $value['value'];
        if($value['name'] == 'Bench Size') $customize_depth = $value['value'];
        if($value['name'] == 'Choose Size') $customize_depth = $value['value'];
        if($value['name'] == 'Table Size') $customize_depth = $value['value'];    
        
		if($value['name'] == 'Attach Drawing (If Custom)') {
		         $drawing = $value['display'];  
				 $customer_msg = htmlspecialchars_decode( $customer_msg . $drawing) ;
		}
    }
    $objPHPExcel->getActiveSheet()->setCellValue('A'.$i, $category)
                ->setCellValue('B'.$i, $color)
                ->setCellValue('C'.$i, $stain_top)
                ->setCellValue('D'.$i, $quantity)
                ->setCellValue('E'.$i, $product_name)  
                ->setCellValue('F'.$i, $customize_depth)
                ->setCellValue('G'.$i, trim($customer_msg))
                ->setCellValue('H'.$i, $customer_name)
                ->setCellValue('I'.$i, $price)
                ->setCellValue('J'.$i, trim($line_subtotal));
 endforeach;
 $j = $i +2;
 $k = $j +1;
 $objPHPExcel->getActiveSheet()->setCellValue('J'.$j, 'Total Cost:')->setCellValue('J'.$k, $total);
 // Set worksheet title
 $objPHPExcel->getActiveSheet()->setTitle('order info');

 $objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
 $objWriter->save($filepath);
 return $file_url;
    }
} //class close