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

error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

date_default_timezone_set('Europe/London');

/** Include PHPExcel_IOFactory */
require_once dirname(__FILE__) . './Classes/PHPExcel/IOFactory.php';


	define('__YK__FILE__','Create xls.xls');

if (!file_exists(__YK__FILE__)) {
	exit("Please run 05featuredemo.php first." . EOL);
}

//echo date('H-i-s') , "Load from Excel2007 file" , EOL;
$callStartTime = microtime(true);

    $objPHPExcel = PHPExcel_IOFactory::load(__YK__FILE__);
	$sheet = $objPHPExcel->getSheet(0); // 读取第一個工作表
	$highestRow = $sheet->getHighestRow(); // 取得总行数
	
	
	$highestColumn = $sheet->getHighestColumn(); // 取得总列数
	
	
	$highestColumn= PHPExcel_Cell::columnIndexFromString($highestColumn); //字母列转换为数字列 如:AA变为27

//	$sheet->setCellValue('B5',2341243);

//	for ($i = 0; $i < 5; $i++) {
//		echo $i;
//	}


	//循环所有
	for ($i = 1; $i <= $highestRow; $i++) {
		for ($col = 0; $col < $highestColumn; $col++) {
			$values = $sheet->getCellByColumnAndRow($col,$i)->getValue();

			//这里写判断; 获取需要修改的列
			$column_s= PHPExcel_Cell::columnIndexFromString('B');

			//需要修改的列
			if(($col+1) == $column_s){
//				echo $values.'--';
				//保存列的位置...修改这里
					$sheet->setCellValue('B'.$i, $values+5);
			}




		}
		echo "<br/>";
	}

	$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
//$objWriter->save(str_replace('.php', '.xls',  __FILE__));
$objWriter->save(__DIR__.DIRECTORY_SEPARATOR.date('H-i-s').'.xls');

//	$sheet = $PHPExcel->getSheet(0); // 读取第一個工作表
//	$highestRow = $sheet->getHighestRow(); // 取得总行数



	echo __DIR__.DIRECTORY_SEPARATOR.date('H-i-s').'.xls';

