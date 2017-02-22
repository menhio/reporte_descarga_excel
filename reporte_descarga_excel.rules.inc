<?php
/*
 * Rules actions - Reporte de Descarga en Excel
 * Implements hook_rules_action_info()
 */
function reporte_descarga_excel_rules_action_info() {
  $actions = array(
    'reporte_descarga_action' => array(
      'label' => t('Generates a Excel file based on Descarga Content Type'),
      'group' => t('Mazpesca Rules'),
      'parameter' => array(
        'account' => array(
          'type' => 'user',
          'label' => t('User that gerenates the file'),
        ),
        'nodes' => array(
          'type' => 'node',
          'label' => t('Select Node'),
        ),
      ),
    ),
  );
  
  return $actions;
}

/*
 * Queries
 */
function descargaMeta($nid, $field) {
  $query = db_select('descarga_metadata_view', 'd');
  $query->addField('d', $field, $field);
  $query->addField('d', 'viaje_nid', 'viaje_nid');
  $query->condition('d.viaje_nid', $nid, '=');
  $exeResults = $query->execute();
  $results = $exeResults->fetchAll();

  return $results[0]->$field;
}

function toneEstimadas($nid, $field) {
  $query = db_select('toneladas_estimadas_view3', 'd');
  $query->addField('d', $field, $field);
  $query->addField('d', 'barco_viaje_nid', 'barco_viaje_nid');
  $query->condition('d.barco_viaje_nid', $nid, '=');
  $exeResults = $query->execute();
  $results = $exeResults->fetchAll();

  return $results[0]->$field;
}

function descargaCal1($nid, $field) {
  $query = db_select('descarga_cal1_view', 'd');
  $query->addField('d', $field, $field);
  $query->addField('d', 'barco_viaje_tid', 'barco_viaje_tid');
  $query->condition('d.barco_viaje_tid', $nid, '=');
  $exeResults = $query->execute();
  $results = $exeResults->fetchAll();

  return $results[0]->$field;
}

function descargaCal2($nid, $field) {
  $query = db_select('descarga_cal2_view', 'd');
  $query->addField('d', $field, $field);
  $query->addField('d', 'barco_viaje_tid', 'barco_viaje_tid');
  $query->condition('d.barco_viaje_tid', $nid, '=');
  $exeResults = $query->execute();
  $results = $exeResults->fetchAll();

  return $results[0]->$field;
}

function pescado2daCalidad($nid, $field) {
  $query = db_select('pescado_2da_calidad_view', 'd');
  $query->addField('d', $field, $field);
  $query->addField('d', 'barco_viaje_nid', 'barco_viaje_nid');
  $query->condition('d.barco_viaje_nid', $nid, '=');
  $exeResults = $query->execute();
  $results = $exeResults->fetchAll();

  return $results[0]->$field;
}

function pescadoRechazo($nid, $field) {
  $query = db_select('pescado_rechazo_view', 'd');
  $query->addField('d', $field, $field);
  $query->addField('d', 'barco_viaje_nid', 'barco_viaje_nid');
  $query->condition('d.barco_viaje_nid', $nid, '=');
  $exeResults = $query->execute();
  $results = $exeResults->fetchAll();

  return $results[0]->$field;
}

/*
 * Action function for reporte_descarga_action
 */
function reporte_descarga_action($account, $nodes) {
  /*
   * Excel Generation
   * Template Path: /public_html/sites/default/files/descarga_final_template.xls
   */
  $nid = $nodes->nid;
  include_once(libraries_get_path('PHPExcel1_8') . '/PHPExcel.php');
  $objPHPExcel = PHPExcel_IOFactory::load("sites/default/files/reporte_descarga_final_tem.xlsx");
  $active_sheet = $objPHPExcel->getActiveSheet();
  
  $objDrawing = new PHPExcel_Worksheet_Drawing();
  $objDrawing->setName('Green Bar');
  $objDrawing->setDescription('Green Bar');
  $objDrawing->setPath('sites/default/files/green_bar.png');
  $objDrawing->setWidth(1250);
  $objDrawing->setCoordinates('B5');
  $objDrawing->setWorksheet($active_sheet);
  
  $objDrawing2 = new PHPExcel_Worksheet_Drawing();
  $objDrawing2->setName('Green Bar');
  $objDrawing2->setDescription('Green Bar');
  $objDrawing2->setPath('sites/default/files/green_bar.png');
  $objDrawing2->setWidth(1250);
  $objDrawing2->setCoordinates('B11');
  $objDrawing2->setWorksheet($active_sheet);
  
  /*
   * Excel Data
   */
  $meses = array("ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO","JULIO",
    "AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE");
  $fecha_inicio = new DateTime(descargaMeta($nid, 'inicio'));
  $fecha_fin = new DateTime(descargaMeta($nid, 'fin'));
  $inicio = $fecha_inicio->format('d') . ' ' . 'DE' . ' ' . 
      $meses[$fecha_inicio->format('n') - 1] . ' ' . $fecha_inicio->format('Y');
  
  $fin = $fecha_fin->format('d') . ' ' . 'DE' . ' ' . 
      $meses[$fecha_fin->format('n') - 1] . ' ' . $fecha_fin->format('Y');
  
  $active_sheet->getCell('A3')->setValue('B/M '. strtoupper(descargaMeta($nid, 'viaje')));
  $active_sheet->setCellValue('G7', $inicio);
  $active_sheet->setCellValue('G8', $fin);
  $active_sheet->setCellValue('P8', descargaMeta($nid, 'bascula'));
  $active_sheet->setCellValue('E10', descargaMeta($nid, 'jefe_maq'));
  /*
   * Toneladas Estimadas
   */
  $active_sheet->setCellValue('AB7', toneEstimadas($nid, 'ton_estimadas'));
  $active_sheet->setCellValue('E17', toneEstimadas($nid, 'aaa_40'));
  $active_sheet->setCellValue('E18', toneEstimadas($nid, 'aaa_20_40'));
  $active_sheet->setCellValue('E19', toneEstimadas($nid, 'aaa_12_20'));
  $active_sheet->setCellValue('E20', toneEstimadas($nid, 'aaa_5_12'));
  $active_sheet->setCellValue('E21', toneEstimadas($nid, 'aaa_2_5'));
  $active_sheet->setCellValue('E22', toneEstimadas($nid, 'bb_3_5'));
  $active_sheet->setCellValue('E23', toneEstimadas($nid, 'bb_2_3_5'));
  $active_sheet->setCellValue('E24', toneEstimadas($nid, 'aaa_1_2'));
  $active_sheet->setCellValue('E25', toneEstimadas($nid, 'bb_1_2'));
  $active_sheet->setCellValue('E26', toneEstimadas($nid, 'bta_2_5'));
  /*
   * Calidad 1
   */
  $active_sheet->setCellValue('O16', descargaCal1($nid, 'aaa_60_cal1'));
  $active_sheet->setCellValue('O17', descargaCal1($nid, 'aaa_40_60_cal1'));
  $active_sheet->setCellValue('O18', descargaCal1($nid, 'aaa_20_40_cal1'));
  $active_sheet->setCellValue('O19', descargaCal1($nid, 'aaaa_12_20_cal1'));
  $active_sheet->setCellValue('O20', descargaCal1($nid, 'aaa_5_12_cal1'));
  $active_sheet->setCellValue('O21', descargaCal1($nid, 'aaa_2_5_cal1'));
  $active_sheet->setCellValue('O22', descargaCal1($nid, 'aaa_1_2_cal1'));
  $active_sheet->setCellValue('O23', descargaCal1($nid, 'aaa_1_cal1'));
  $active_sheet->setCellValue('O24', descargaCal1($nid, 'bte_5_cal1'));
  $active_sheet->setCellValue('O25', descargaCal1($nid, 'bte_3_5a5_cal1'));
  $active_sheet->setCellValue('O26', descargaCal1($nid, 'bte_2a_3_5_cal1'));
  $active_sheet->setCellValue('O27', descargaCal1($nid, 'bte_1a2_cal1'));
  $active_sheet->setCellValue('O28', descargaCal1($nid, 'bte_1_cal1'));
  $active_sheet->setCellValue('O29', descargaCal1($nid, 'bte_ngo_1a2_cal1'));
  /*
   * Calidad 2
   */
  $active_sheet->setCellValue('S16', descargaCal2($nid, 'aaa_60_cal2'));
  $active_sheet->setCellValue('S17', descargaCal2($nid, 'aaa_40_60_cal2'));
  $active_sheet->setCellValue('S18', descargaCal2($nid, 'aaa_20_40_cal2'));
  $active_sheet->setCellValue('S19', descargaCal2($nid, 'aaaa_12_20_cal2'));
  $active_sheet->setCellValue('S20', descargaCal2($nid, 'aaa_5_12_cal2'));
  $active_sheet->setCellValue('S21', descargaCal2($nid, 'aaa_2_5_cal2'));
  $active_sheet->setCellValue('S22', descargaCal2($nid, 'aaa_1_2_cal2'));
  $active_sheet->setCellValue('S23', descargaCal2($nid, 'aaa_1_cal2'));
  $active_sheet->setCellValue('S24', descargaCal2($nid, 'bte_5_cal2'));
  $active_sheet->setCellValue('S25', descargaCal2($nid, 'bte_3_5a5_cal2'));
  $active_sheet->setCellValue('S26', descargaCal2($nid, 'bte_2a_3_5_cal2'));
  $active_sheet->setCellValue('S27', descargaCal2($nid, 'bte_1a2_cal2'));
  $active_sheet->setCellValue('S28', descargaCal2($nid, 'bte_1_cal2'));
  $active_sheet->setCellValue('S29', descargaCal2($nid, 'bte_ngo_1a2_cal2'));
  /*
   * Clasificación de Pescado de Segunda Calidad
   */
  $active_sheet->setCellValue('O40', pescado2daCalidad($nid, 'marcado_cal2'));
  $active_sheet->setCellValue('O41', pescado2daCalidad($nid, 'tallado_cal2'));
  $active_sheet->setCellValue('O42', pescado2daCalidad($nid, 'reventado_cal2'));
  $active_sheet->setCellValue('O43', pescado2daCalidad($nid, 'mal_olor_cal2'));
  /*
   * Pescado de Rechazo
   */
  $active_sheet->setCellValue('Z40', pescadoRechazo($nid, 'aaa_bte_recha'));
  $active_sheet->setCellValue('Z41', pescadoRechazo($nid, 'bule_recha'));
  $active_sheet->setCellValue('Z42', pescadoRechazo($nid, 'fauna_recha'));
  //$active_sheet->setCellValue('Z43', pescado2daCalidad($nid, 'mal_olor_cal2'));
  /*
   * Excel output metadata
   */
  header('Content-Type: application/vnd.ms-excel');
  header('Content-Disposition: attachment;filename="Reporte_Descarga_Final.xlsx"');
  header('Cache-Control: max-age=0');
  // If you're serving to IE 9, then the following may be needed
  header('Cache-Control: max-age=1');
  // If you're serving to IE over SSL, then the following may be needed
  header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
  header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
  header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
  header ('Pragma: public'); // HTTP/1.0
  $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
  //$objWriter->setIncludeCharts(TRUE);
  $objWriter->save('php://output');
  exit;
  
}

