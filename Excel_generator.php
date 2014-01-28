<?php

/**
 * Excel Generator for CodeIgniter
 * 
 * @author Dida Nurwanda <didanurwanda@gmail.com>
 * @link http://didanurwanda.blogspot.com 
 * 
 */

require_once dirname(__FILE__) . '/PHPExcel/PHPExcel.php';

class Excel_generator extends PHPExcel {

    /**
     * @var CI_DB_result
     */
    private $query;
    private $column = array();
    private $header = array();
    private $width = array();
    private $set_bold = TRUE;

    /**
     * Diisi dengan query Anda
     * <pre>
     * $query = $this->db->get('users');
     * $this->excel_generator->set_query($query);
     * </pre>
     * 
     * @access public
     * @param CI_DB_result $query
     * @return void
     */
    public function set_query(CI_DB_result $query) {
        $this->query = $query;
    }

    /**
     * Diisi sesuai dengan field pada table
     * <pre>
     * $this->excel_generator->set_column(array('name', 'address', 'email'));
     * </pre>
     * 
     * @access public
     * @param array $column
     * @return void
     */
    public function set_column($column = array()) {
        $this->column = $column;
    }

    /**
     * Untuk mengisi header pada table excel
     * <pre>
     * $this->excel_generator->set_header(array('Name', 'Address', 'Email'));
     * </pre>
     * Jika ingin tulisannya tidak dalam bentuk bold
     * <pre>
     * $this->excel_generator->set_header(array('...'), FALSE);
     * </pre>
     * 
     * @access public
     * @param array $header
     * @param bool $set_bold
     */
    public function set_header($header = array(), $set_bold = TRUE) {
        $this->header = $header;
        $this->set_bold = $set_bold;
    }
    
    public function set_width($width = array()) {
        $this->width = $width;
    }

    /**
     * Untuk menghasilkan data excel
     * 
     * @access public
     * @return void
     */
    public function generate() {
        $start = 1;
        if (count($this->header) > 0) {
            foreach ($this->header as $row) {
                $this->getActiveSheet()->setCellValue($this->columnName($start) . '1', $row);
                $this->getActiveSheet()->getStyle($this->columnName($start) . '1')->getFont()->setBold(TRUE);
                $start++;
            }
            $start = 2;
        }

        foreach ($this->query->result_array() as $result_db) {
            $index = 1;
            foreach ($this->column as $row) {
                if(count($this->width) > 0) {
                    $this->getActiveSheet()->getColumnDimension($this->columnName($index))->setWidth($this->width[$index-1]);
                }
                
                $this->getActiveSheet()->setCellValue($this->columnName($index) . $start, $result_db[$row]);
                $index++;
            }
            $start++;
        }
    }

    private function columnName($index) {
        --$index;
        if ($index >= 0 && $index < 26)
            return chr(ord('A') + $index);
        else if ($index > 25)
            return ($this->columnName($index / 26)) . ($this->columnName($index % 26 + 1));
        else
            show_error("Invalid Column # " . ($index + 1));
    }

    /**
     * Untuk membuat file excel
     * 
     * @param string $filename
     * @param string $writerType
     * @param string $mimes
     */
    private function writeToFile($filename = 'doc', $writerType = 'Excel5', $mimes = 'application/vnd.ms-excel') {
        $this->generate();
        header("Content-Type: $mimes");
        header("Content-Disposition: attachment;filename=\"$filename\"");
        header("Cache-Control: max-age=0");
        $objWriter = PHPExcel_IOFactory::createWriter($this, $writerType);
        $objWriter->save('php://output');
    }

    /**
     * @param string $filename
     */
    public function exportTo2003($filename = 'doc') {
        $this->writeToFile($filename . '.xls');
    }

    /**
     * @param string $filename
     */
    public function exportTo2007($filename = 'doc') {
        $this->writeToFile($filename . '.xlsx', 'Excel2007', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    }

}