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

function fechasTurnos($nid, $row, $field) {
  $query = db_select('fechas_turnos_view', 'd');
  $query->addField('d', $field, $field);
  $query->addField('d', 'barco_viaje_nid', 'barco_viaje_nid');
  $query->condition('d.barco_viaje_nid', $nid, '=');
  $exeResults = $query->execute();
  $results = $exeResults->fetchAll();

  return $results[$row]->$field;
}

function descargaPorTurno($nid, $row, $field, $table) {
  $query = db_select($table, 'd');
  $query->addField('d', $field, $field);
  $query->addField('d', 'barco_viaje_nid', 'barco_viaje_nid');
  $query->addField('d', 'fecha', 'fecha');
  $query->condition('d.barco_viaje_nid', $nid, '=');
  $query->orderBy('fecha');
  $exeResults = $query->execute();
  $results = $exeResults->fetchAll();

  return $results[$row]->$field;
}

function countTableRows($table) {
  $query = db_select($table, 'd');
  //$query->addField('d', 'eid', 'eid');
  $num_rows = $query->countQuery()->execute()->fetchField();
  return $num_rows;
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
  $active_sheet->getSheetView()->setView(PHPExcel_Worksheet_SheetView::SHEETVIEW_NORMAL);
  $objDrawing = new PHPExcel_Worksheet_Drawing();
  $objDrawing->setName('Logo');
  $objDrawing->setDescription('Logo');
  $objDrawing->setPath('sites/default/files/logo_excel.png');
  $objDrawing->setOffsetX(160);                      
  $objDrawing->setOffsetY(10);
  $objDrawing->setWidth(237.35);
  $objDrawing->setHeight(58.58); 
  $objDrawing->setWorksheet($active_sheet);
  
  $objDrawing2 = new PHPExcel_Worksheet_Drawing();
  $objDrawing2->setName('Green Bar');
  $objDrawing2->setDescription('Green Bar');
  $objDrawing2->setPath('sites/default/files/green_bar.png');
  $objDrawing2->setHeight(16.62); 
  $objDrawing2->setWidth(1360);
  $objDrawing2->setCoordinates('B5');
  $objDrawing2->setWorksheet($active_sheet);
  
  $objDrawing3 = new PHPExcel_Worksheet_Drawing();
  $objDrawing3->setName('Green Bar');
  $objDrawing3->setDescription('Green Bar');
  $objDrawing3->setPath('sites/default/files/green_bar.png');
  $objDrawing3->setHeight(16.62);
  $objDrawing3->setWidth(1360);
  $objDrawing3->setCoordinates('B11');
  $objDrawing3->setWorksheet($active_sheet);
  
  $objDrawing4 = new PHPExcel_Worksheet_Drawing();
  $objDrawing4->setName('Green Bar');
  $objDrawing4->setDescription('Green Bar');
  $objDrawing4->setPath('sites/default/files/green_bar.png');
  $objDrawing4->setHeight(16.62);
  $objDrawing4->setWidth(1360);
  $objDrawing4->setCoordinates('B48');
  $objDrawing4->setWorksheet($active_sheet);
  
  $objDrawing5 = new PHPExcel_Worksheet_Drawing();
  $objDrawing5->setName('Green Bar');
  $objDrawing5->setDescription('Green Bar');
  $objDrawing5->setPath('sites/default/files/green_bar.png');
  $objDrawing5->setHeight(16.62);
  $objDrawing5->setWidth(1360);
  $objDrawing5->setCoordinates('B66');
  $objDrawing5->setWorksheet($active_sheet);
  
  $objDrawing6 = new PHPExcel_Worksheet_Drawing();
  $objDrawing6->setName('Green Bar');
  $objDrawing6->setDescription('Green Bar');
  $objDrawing6->setPath('sites/default/files/green_bar.png');
  $objDrawing6->setHeight(16.62);
  $objDrawing6->setWidth(1360);
  $objDrawing6->setCoordinates('B94');
  $objDrawing6->setWorksheet($active_sheet);
  
  $objDrawing7 = new PHPExcel_Worksheet_Drawing();
  $objDrawing7->setName('Green Bar');
  $objDrawing7->setDescription('Green Bar');
  $objDrawing7->setPath('sites/default/files/green_bar.png');
  $objDrawing7->setHeight(16.62);
  $objDrawing7->setWidth(1360);
  $objDrawing7->setCoordinates('B103');
  $objDrawing7->setWorksheet($active_sheet);
  
  $objDrawing8 = new PHPExcel_Worksheet_Drawing();
  $objDrawing8->setName('Green Bar');
  $objDrawing8->setDescription('Green Bar');
  $objDrawing8->setPath('sites/default/files/green_bar.png');
  $objDrawing8->setHeight(16.62);
  $objDrawing8->setWidth(1360);
  $objDrawing8->setCoordinates('B48');
  $objDrawing8->setWorksheet($active_sheet);
  
  /*
   * Cascos
   */
  function createCasco($w, $y, $x, $coord, $active_sheet, $image) {
    $casco1 = new PHPExcel_Worksheet_Drawing();
    $casco1->setName('Casco');
    $casco1->setDescription('Casco');
    $casco1->setPath($image);
    //$casco1->setHeight(16.62);
    $casco1->setWidth($w);
    $casco1->setOffsetY($y);                      
    $casco1->setOffsetX($x);                      
    $casco1->setCoordinates($coord);
    $casco1->setWorksheet($active_sheet);
  }
  
  createCasco(138, 10, 161, 'D68', $active_sheet, 'sites/default/files/casco.png');
  createCasco(138, 0, 159, 'D82', $active_sheet, 'sites/default/files/casco_back.png');
  
  createCasco(138, 10, 159, 'L68', $active_sheet, 'sites/default/files/casco.png');
  createCasco(138, 0, 159, 'L82', $active_sheet, 'sites/default/files/casco_back.png');
  
  createCasco(138, 10, 160, 'S68', $active_sheet, 'sites/default/files/casco.png');
  createCasco(138, 0, 159, 'S82', $active_sheet, 'sites/default/files/casco_back.png');
  
  createCasco(138, 0, 159, 'A106', $active_sheet, 'sites/default/files/casco.png');
  createCasco(138, 0, 159, 'A121', $active_sheet, 'sites/default/files/casco_back.png');
  
  createCasco(138, 0, 159, 'E106', $active_sheet, 'sites/default/files/casco.png');
  createCasco(138, 0, 159, 'E121', $active_sheet, 'sites/default/files/casco_back.png');
  
  createCasco(138, 0, 159, 'I106', $active_sheet, 'sites/default/files/casco.png');
  createCasco(138, 0, 159, 'I121', $active_sheet, 'sites/default/files/casco_back.png');
  
  createCasco(138, 0, 159, 'M106', $active_sheet, 'sites/default/files/casco.png');
  createCasco(138, 0, 159, 'M121', $active_sheet, 'sites/default/files/casco_back.png');
  
  createCasco(138, 0, 159, 'Q106', $active_sheet, 'sites/default/files/casco.png');
  createCasco(138, 0, 159, 'Q121', $active_sheet, 'sites/default/files/casco_back.png');
  
  createCasco(138, 0, 159, 'U106', $active_sheet, 'sites/default/files/casco.png');
  createCasco(138, 0, 159, 'U121', $active_sheet, 'sites/default/files/casco_back.png');
  
  createCasco(138, 0, 159, 'Y106', $active_sheet, 'sites/default/files/casco.png');
  createCasco(138, 0, 159, 'Y121', $active_sheet, 'sites/default/files/casco_back.png');
  
  createCasco(138, 0, 159, 'AC106', $active_sheet, 'sites/default/files/casco.png');
  createCasco(138, 0, 159, 'AC121', $active_sheet, 'sites/default/files/casco_back.png');
  
  createCasco(138, 0, 159, 'A124', $active_sheet, 'sites/default/files/casco.png');
  createCasco(138, 0, 159, 'A139', $active_sheet, 'sites/default/files/casco_back.png');
  
  createCasco(138, 0, 159, 'E124', $active_sheet, 'sites/default/files/casco.png');
  createCasco(138, 0, 159, 'E139', $active_sheet, 'sites/default/files/casco_back.png');
  
  createCasco(138, 0, 159, 'I124', $active_sheet, 'sites/default/files/casco.png');
  createCasco(138, 0, 159, 'I139', $active_sheet, 'sites/default/files/casco_back.png');
  
  createCasco(138, 0, 159, 'M124', $active_sheet, 'sites/default/files/casco.png');
  createCasco(138, 0, 159, 'M139', $active_sheet, 'sites/default/files/casco_back.png');
  
  createCasco(138, 0, 159, 'Q124', $active_sheet, 'sites/default/files/casco.png');
  createCasco(138, 0, 159, 'Q139', $active_sheet, 'sites/default/files/casco_back.png');
  
  createCasco(138, 0, 159, 'U124', $active_sheet, 'sites/default/files/casco.png');
  createCasco(138, 0, 159, 'U139', $active_sheet, 'sites/default/files/casco_back.png');
  
  createCasco(138, 0, 159, 'Y124', $active_sheet, 'sites/default/files/casco.png');
  createCasco(138, 0, 159, 'Y139', $active_sheet, 'sites/default/files/casco_back.png');
  
  createCasco(138, 0, 159, 'AC124', $active_sheet, 'sites/default/files/casco.png');
  createCasco(138, 0, 159, 'AC139', $active_sheet, 'sites/default/files/casco_back.png');
  
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
   * Fechas de los turnos
   */
  $col = 53;
  for ($i = 0; $i < countTableRows('fechas_turnos_view'); $i++) {
    $fecha_turno = new DateTime(fechasTurnos($nid, $i, 'fecha'));
    $dia_turno = $fecha_turno->format('d');
    $mes_turno = substr($meses[$fecha_turno->format('n') - 1], 0, 3);
    $active_sheet->setCellValue('A' . $col, $dia_turno . '-' . $mes_turno);
    $col++;
  }
  /*
   * Primer Turno
   */
  $col = 53;
  for ($i = 0; $i < countTableRows('descarga_primer_turno_view'); $i++) {
    $active_sheet->setCellValue('C' . $col, descargaPorTurno($nid, $i, 
        'toneladas', 'descarga_primer_turno_view'));
    $col++;
  }
  /*
   * Segundo Turno
   */
  $col = 53;
  for ($i = 0; $i < countTableRows('descarga_segundo_turno_view'); $i++) {
    $active_sheet->setCellValue('F' . $col, descargaPorTurno($nid, $i, 
        'toneladas', 'descarga_segundo_turno_view'));
    $col++;
  }
  /*
   * Tercer Turno
   */
  $col = 53;
  for ($i = 0; $i < countTableRows('descarga_tercer_turno_view'); $i++) {
    $active_sheet->setCellValue('I' . $col, descargaPorTurno($nid, $i, 
        'toneladas', 'descarga_tercer_turno_view'));
    $col++;
  }
  /*
   * Charts
   */
  /* Labels */
  $sheetTitle = $active_sheet->getTitle();
  $dsl = array(
    new PHPExcel_Chart_DataSeriesValues('String', '', NULL, 1),
  );
  /* X Axis Value Label */
  $xsl = array(
    new PHPExcel_Chart_DataSeriesValues('String', "'" . $sheetTitle . "'" . '!$K$16:$K$31', NULL, 90)
  );
  /* Data Series Values */
  $dsv = array(
    new PHPExcel_Chart_DataSeriesValues('Number', "'" . $sheetTitle . "'" . '!$W$16:$W$31', NULL, 90)
  );
  $ds = new PHPExcel_Chart_DataSeries(
      PHPExcel_Chart_DataSeries::TYPE_BARCHART, 
      PHPExcel_Chart_DataSeries::GROUPING_STANDARD,
      range(0, count($dsv)-1),
      $dsl, $xsl, $dsv
  );
  $ds->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);
  
  $pa=new PHPExcel_Chart_PlotArea(NULL, array($ds));
  $legend=new PHPExcel_Chart_Legend(\PHPExcel_Chart_Legend::POSITION_RIGHT, 
      NULL, false);
  $title=new PHPExcel_Chart_Title('CLASIFICACIÓN GENERAL DEL PESCADO');
  $chart= new PHPExcel_Chart('chart1', $title, $legend, $pa, true, 0, NULL, NULL);
  $chart->setTopLeftPosition('Q84');
  $chart->setBottomRightPosition('AF92');
  $active_sheet->addChart($chart);
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
  $objWriter->setIncludeCharts(TRUE);
  $objWriter->save('php://output');
  exit;
  
}

