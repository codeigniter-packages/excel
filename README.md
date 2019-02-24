# Excel

## How to use

```bash

composer require codeigniter-packages/installer
composer require codeigniter-packages/loader
composer require codeigniter-packages/excel

```

```php

$this->load->package('codeigniter-packages/excel');

$data =array(
	array('field1'=>'value1','field2'=>'value2')
);
$this->excel->query_to_excel(array('field1','field2'),$data,'data.xls');

$this->excel->execl_to_array('./data.xls',array('excel_field1'=>'field1','excel_field2'=>'filed2'));



```