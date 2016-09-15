package com.aspose.cells.examples.articles;

public class WorkingWithPHP {

	<?php
			require_once("java/Java.inc");
			require("AsposeCells.php");
			$workbook = ClassFactory::createWorkbook();
			$workbook->open5("t1.xls");
			$cell = $workbook->getWorksheets()->get(0)->getCells()->getCell(0, 0);
			$cell->setValue6("Hello World!"); 
			$workbook->save5("t.xls");
			?>

}
