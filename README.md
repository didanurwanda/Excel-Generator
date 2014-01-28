Excel Generator
===============

PHPExcel Generator for CodeIgniter

#### Example
```php
class Welcome extends CI_Controller {

    public function __construct() {
        $this->load->database();
        $this->load->library('Excel_generator');
    }
    
    public function index() {
        $query = $this->db->get('users');
        $this->excel_generator->set_query($query);
        $this->excel_generator->set_header(array('Name', 'Address', 'Email'));
        $this->excel_generator->set_column(array('name', 'address', 'email'));
        $this->excel_generator->set_width(array(25, 30, 15));
        $this->excel_generator->exportTo2007('report users');
	}
}
```