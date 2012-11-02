Integrate OSExcelBundle in your project
----------------------------------------

Add this line to the require option in your composer.json file:

     "os/excel-bundle": "dev-master"

Add autoloader for PHPExcel in app/autoloader.php

     require __DIR__.'/../vendor/os/php-excel/PHPExcel/PHPExcel.php';

Execute this command line

     php composer.phar install

USING:
------

Call service:

     $excel = $this->get('os.excel');

Select sheet

     $excel->setActiveSheet([INDEX_OF_SHEET]);

Get count of sheet

     $excel->getSheetCount();

Get sheet names

     $excel->getSheetNames();

Get count of rows of selected sheet

     $excel->getRowCount();

Get Count of columns of selected sheet

     $excel->getRowCount();

Get index of last column

     $excel->getHighestColumn();

Get row data

     $excel->getRowData([INDEX_OF_ROW]);

Get cell data

     $excel->getCellData([INDEX_OF_ROW], [INDEX_OF_COLUMN]);

GET sheet data as php array

     $excel->getSheetData();

Enjoy!
