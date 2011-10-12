<?php
namespace OS\ExcelBundle\Excel;

require str_replace('bundles/OS/ExcelBundle/Excel', '',__DIR__).'/PHPExcel/PHPExcel.php';

use PHPExcel_IOFactory;

class Excel {
    
    private $options = array('readOnly' => true);
    private $reader;
    private $phpExcel;
    private $currentSheet;
    
    public function loadFile($file = null, $options = array()) {
        $this->options += $options;
        
        if (!$file)
            throw new \Symfony\Component\HttpKernel\Exception\NotFoundHttpException('Missing arguments!');

        extract($this->options);

        $this->reader = PHPExcel_IOFactory::createReader(PHPExcel_IOFactory::identify($file));  
        $this->reader->setReadDataOnly($readOnly);
        $this->phpExcel = $this->reader->load($file);
    }
    
    public function getSheetCount() {
        return $this->phpExcel->getSheetCount();
    }
    
    public function getSheetNames() {
        return $this->phpExcel->getSheetNames();
    }
    
    public function getSheet($indice = 0) {
        $this->currentSheet = $this->phpExcel->setActiveSheetIndex($indice);
        
        return $this->currentSheet;
    }
    
    public function getCellData($row = 0, $col = 0) {
        if (!$this->currentSheet)
            $this->getSheet();
        
        return $this->currentSheet->getCellByColumnAndRow($col, $row)->getValue();
    }
    
}
?>
