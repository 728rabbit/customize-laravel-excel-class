**Install maatwebsite/excel, version 3.1**

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

------------

**How to use**

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
			]
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
