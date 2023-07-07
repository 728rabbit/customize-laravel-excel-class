<?php
/* 
Install maatwebsite/excel, version 3.1

1.  composer require maatwebsite/excel

2.  add to config/app.php 
    
    'providers' => [
        ...
        Maatwebsite\Excel\ExcelServiceProvider::class,
    ]

    'aliases' => [
        ...
        'Excel' => Maatwebsite\Excel\Facades\Excel::class,
    ]

3.  composer update

How to use

1.  export excel with single sheet
    
    return (new \App\Http\Controllers\ExcelSheet())->setData([
        [
            'id' => 1, 
            'name' => 'Chan Tai Man'
        ],
        [
            'id' => 2, 
            'name' => 'Lily Ho'
        ]
    ])->setHeader([
        'ID',
        'Name'
    ])->doExport('user.xlsx');

2.  export excel with multi sheet
    
    return (new \App\Http\Controllers\ExcelSheet())->setData([
        'my_sheet_1' => 
        [
            [
                'id' => 1, 
                'name' => 'Chan Tai Man'
            ],
            [
                'id' => 2, 
                'name' => 'Lily Ho'
            ]
        ],
        'my_sheet_2' => 
        [
            [
                'id' => 3, 
                'name' => 'Petter'
            ],
            [
                'id' => 4, 
                'name' => 'Jim'
            ]
        ],
    ])->setHeader([
        'header_1' => 
        [
            'ID',
            'Name'
        ],
        'header_2' => 
        [
            'ID',
            'Name'
        ]
    ])->doExport('user.xlsx', true);
*/

namespace App\Http\Controllers;

use Maatwebsite\Excel\Concerns\Exportable;
use Maatwebsite\Excel\Concerns\WithMultipleSheets;
use Maatwebsite\Excel\Concerns\WithTitle;

use Maatwebsite\Excel\Concerns\ShouldAutoSize;
use Maatwebsite\Excel\Concerns\WithHeadings;
use Maatwebsite\Excel\Concerns\FromArray;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Events\AfterSheet;

use Maatwebsite\Excel\Concerns\WithStyles;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class ExcelSheet {
    use Exportable;
    
    protected $_data = [];
    protected $_header = [];

    public function setData($data) {
        $this->_data = $data;
        return $this;
    }

    public function setHeader($header) {
        $this->_header = $header;
        return $this;
    }

    public function doExport($file_name, $multi = false) {
        if(empty($multi)) {
            return \Maatwebsite\Excel\Facades\Excel::download(new ExcelData($this->_data,$this->_header), $file_name);
        }
        else {
            return \Maatwebsite\Excel\Facades\Excel::download(new ExcelMultiData($this->_data,$this->_header), $file_name);
        }
    }
}

class ExcelMultiData implements WithMultipleSheets {
    use Exportable;
    
    protected $_sheet_header = [];
    protected $_sheet_data = [];

    public function __construct($data = [], $header = []){
        $this->_sheet_data = (array)$data;
        $this->_sheet_header = (array)$header;
    }
    
    public function sheets(): array {
        $loop_data = $this->_sheet_data;
        
        sort($this->_sheet_data);
        sort($this->_sheet_header);

        $sheets = [];
        $k = 0;
        foreach ($loop_data as $key => $data) {
            $sheets[] = new ExcelData(
                ((!empty($this->_sheet_data[$k]))?$this->_sheet_data[$k]:[]),
                ((!empty($this->_sheet_header[$k]))?$this->_sheet_header[$k]:[]),
                $key
            );
            $k++;
        }
        return $sheets;
    }
}

class ExcelData implements WithHeadings,FromArray,WithTitle,WithEvents,ShouldAutoSize,WithStyles {
    protected $_sheet_header = [];
    protected $_sheet_data = [];
    protected $_sheet_title = '';

    public function __construct($data = [], $header = [], $title = ''){
        $this->_sheet_data = (array)$data;
        $this->_sheet_header = (array)$header;
        $this->_sheet_title = (string)$title;
    }
    
    public function headings(): array {
        return $this->_sheet_header;
    }

    public function array(): array {
        return $this->_sheet_data;
    }
    
    public function title(): string{
        return (!empty($this->_sheet_title))?$this->_sheet_title:'Worksheet';
    }
    
    public function registerEvents(): array {
        return 
        [
            AfterSheet::class    => function(AfterSheet $event) {
                if(!empty($this->_sheet_header)) {
                    $event->sheet->getDelegate()->freezePane('A2')->getStyle('A1')->getFont()->setBold(true);
                }
            }
        ];
    }
    
    public function styles(Worksheet $sheet){
        if(!empty($this->_sheet_header)) {
            return [
                // style the first row as bold text.
                1    => 
                [
                    'font' => 
                    [
                        'bold' => true    
                    ]
                ] 
            ];
        }
    }
}
