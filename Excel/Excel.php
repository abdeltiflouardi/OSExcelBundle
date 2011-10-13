<?php

namespace OS\ExcelBundle\Excel;

require str_replace('bundles/OS/ExcelBundle/Excel', '', __DIR__) . '/PHPExcel/PHPExcel.php';

use PHPExcel_IOFactory,
    PHPExcel_Cell;

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

        $this->setActiveSheet();
    }

    public function getSheetCount() {
        return $this->phpExcel->getSheetCount();
    }

    public function getSheetNames() {
        return $this->phpExcel->getSheetNames();
    }

    public function setActiveSheet($indice = 0) {
        $this->currentSheet = $this->phpExcel->setActiveSheetIndex($indice);
    }

    public function getSheet() {
        return $this->currentSheet;
    }

    public function getRowCount() {
        return $this->getSheet()->getHighestRow();
    }

    public function getHighestColumn() {
        return $this->getSheet()->getHighestColumn();
    }

    public function getColumnCount() {
        return PHPExcel_Cell::columnIndexFromString($this->getHighestColumn());
    }

    public function getCellData($row = 0, $col = 0) {
        return $this->getSheet()->getCellByColumnAndRow($col, $row)->getValue();
    }

    public function getSheetData() {
        $data = array();
        for ($row = 1; $row <= $this->getRowCount(); ++$row) {
            for ($col = 0; $col <= $this->getColumnCount(); ++$col) {
                $value = $this->getCellData($row, $col);

                if (is_array($data)) {
                    $data[$row - 1][$col] = $value;
                }
            }
        }

        return $data;
    }

}

?>
