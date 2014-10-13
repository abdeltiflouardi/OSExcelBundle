Integrate OSExcelBundle in your project
----------------------------------------

Add this line to the require option in your composer.json file:

     "os/excel-bundle": "dev-master"

Execute this command line

     php composer.phar update os/excel-bundle

In your app/AppKernel.php

     new OS\ExcelBundle\OSExcelBundle

USING:
------

Call service

     $excel = $this->get('os.excel');

Open file

     $excel->loadFile([FILE_PATHNAME]);

Select sheet

     $excel->setActiveSheet([INDEX_OF_SHEET]);

Get count of sheet

     $excel->getSheetCount();

Get sheet names

     $excel->getSheetNames();

Get count of rows of selected sheet

     $excel->getRowCount();

Get Count of columns of selected sheet

     $excel->getColumnCount();

Get index of last column

     $excel->getHighestColumn();

Get row data

     $excel->getRowData([INDEX_OF_ROW]);

Get cell data

     $excel->getCellData([INDEX_OF_ROW], [INDEX_OF_COLUMN]);

GET sheet data as php array

     $excel->getSheetData();

Enjoy!
