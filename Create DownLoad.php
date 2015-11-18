<?php
    /**
     * PHPExcel
     *
     * Copyright (C) 2006 - 2014 PHPExcel
     *
     * This library is free software; you can redistribute it and/or
     * modify it under the terms of the GNU Lesser General Public
     * License as published by the Free Software Foundation; either
     * version 2.1 of the License, or (at your option) any later version.
     *
     * This library is distributed in the hope that it will be useful,
     * but WITHOUT ANY WARRANTY; without even the implied warranty of
     * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
     * Lesser General Public License for more details.
     *
     * You should have received a copy of the GNU Lesser General Public
     * License along with this library; if not, write to the Free Software
     * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
     *
     * @category   PHPExcel
     * @package    PHPExcel
     * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
     * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
     * @version    1.8.0, 2014-03-02
     */
    

    /** Error reporting */
    error_reporting(E_ALL);
    ini_set('display_errors', TRUE);
    ini_set('display_startup_errors', TRUE);
    date_default_timezone_set('Europe/London');

    define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

    /** Include PHPExcel */
    require_once dirname(__FILE__) . './Classes/PHPExcel.php';


// Create new PHPExcel object

    $objPHPExcel = new PHPExcel();

// Set document propertiese

    $objPHPExcel->getProperties()->setCreator("Maarten Balliauw")
        ->setLastModifiedBy("Maarten Balliauw")
        ->setTitle("PHPExcel Test Document")
        ->setSubject("PHPExcel Test Document")
        ->setDescription("Test document for PHPExcel, generated using PHP classes.")
        ->setKeywords("office PHPExcel php")
        ->setCategory("Test result file");


// Add some data

    $objPHPExcel->setActiveSheetIndex(0)
        ->setCellValue('A1', 'Helloasdfasdfasdf')
        ->setCellValue('B2', 'world!')
        ->setCellValue('C1', 'Hello')
        ->setCellValue('D2', 'world!');

// Miscellaneous glyphs, UTF-8
    $objPHPExcel->setActiveSheetIndex(0)
        ->setCellValue('A4', 'Miscellaneous glyphs')
        ->setCellValue('A5', 'éàèùâêîôûëïüÿäöüç');


    $objPHPExcel->getActiveSheet()->setCellValue('A8',"Hello".PHP_EOL."World");
    $objPHPExcel->getActiveSheet()->getRowDimension(8)->setRowHeight(-1);
    $objPHPExcel->getActiveSheet()->getStyle('A8')->getAlignment()->setWrapText(true);


    $objPHPExcel->setActiveSheetIndex(0);


    $callStartTime = microtime(true);


//    echo PHP_EOL;
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
    $objWriter->save(str_replace('.php', '.xls', __FILE__));

//    var_dump(str_replace('.php', '.xls', __FILE__));

//    header('Content-Type: text/html');
    header('Content-Disposition: attachment;filename="01simple.xls"');
//    header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
//    header('Cache-Control: max-age=1');

// If you're serving to IE over SSL, then the following may be needed
//    header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
//    header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
//    header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
//    header ('Pragma: public'); // HTTP/1.0

    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
    $objWriter->save('php://output');