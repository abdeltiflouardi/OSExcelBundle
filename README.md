Integrate OSExcelBundle in your project
----------------------------------------

Add this lines in require option in your composer.json file:

     "os/excel-bundle": "dev-master",
     "os/php-excel": "dev-master"

Add autoloader for PHPExcel in app/autoloader.php

     require __DIR__.'/../vendor/os/php-excel/PHPExcel/PHPExcel.php';

Execute this command line

     php composer.phar install


Enjoy!
