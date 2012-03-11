<?php

namespace OS\ExcelBundle\Excel;

require preg_replace ('/(.{1})bundles(.{1})OS(.{1})ExcelBundle(.{1})Excel/', '', DIR) . '/PHPExcel/PHPExcel.php';

use PHPExcel_IOFactory,
    PHPExcel_Cell;

class Excel {

    private $options = array('readOnly' => true);
    private $reader;
    private $phpExcel;
    private $currentSheet;

    /**
     * Parse and load and spreadsheet file
     * 
     * @param string $file
     * @param array $options 
     */
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

    /**
     * Number of sheets
     * 
     * @return int
     */
    public function getSheetCount() {
        return $this->phpExcel->getSheetCount();
    }

    /**
     * Sheet names
     * 
     * @return array
     */
    public function getSheetNames() {
        return $this->phpExcel->getSheetNames();
    }

    /**
     * Active a sheet by indix
     *
     * @param int $index
     */
    public function setActiveSheet($index = 0) {
        $this->currentSheet = $this->phpExcel->setActiveSheetIndex($index);
    }

    /**
     * Return activated sheet
     * 
     * @return \PHPExcel_Worksheet
     */
    public function getSheet() {
        return $this->currentSheet;
    }

    /**
     * Number of lines
     * 
     * @return int
     */
    public function getRowCount() {
        return $this->getSheet()->getHighestRow();
    }

    /**
     * Char of highest column (A, B, C, ...)
     * 
     * @return string
     */
    public function getHighestColumn() {
        return $this->getSheet()->getHighestColumn();
    }

    /**
     * Number of columns
     * 
     * @return int
     */
    public function getColumnCount() {
        return PHPExcel_Cell::columnIndexFromString($this->getHighestColumn());
    }

    /**
     * Value of defined cellule
     *
     * @param int $row
     * @param int $col
     * @return mixed 
     */
    public function getCellData($row = 0, $col = 0) {
        return $this->getSheet()->getCellByColumnAndRow($col, $row)->getValue();
    }

    /**
     * Return row data
     * 
     * @param int $index
     * @return array $row_data 
     */
    public function getRowData($index = 0) {
        $row_data = array();
        for ($col = 0, $row = $index; $col < $this->getColumnCount(); $col++)
            $row_data[] = $this->getCellData($row, $col);

        return $row_data;
    }

    /**
     * Return current sheet data
     * 
     * @return array $data
     */
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
