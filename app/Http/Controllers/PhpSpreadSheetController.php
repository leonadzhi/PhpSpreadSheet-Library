<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class PhpSpreadSheetController extends Controller
{
    public function index()
    {
        $spreadsheet = new Spreadsheet();

        $arrayData = json_decode(file_get_contents('data.json'));
        $sheet = $spreadsheet->getActiveSheet()->fromArray(
            $arrayData,
            NULL,
            'C3');

        $arrayData2 = json_decode(file_get_contents('header.json'));

        $sheet = $spreadsheet->getActiveSheet()->fromArray(
            $arrayData2,
            NULL,
            'C2');
        $writer = new Xlsx($spreadsheet);

        //Дизайн ячеек
        $sheet->getStyle('C3:H26')->applyFromArray([
            'font' => [
                'name' => 'Arial',
                'bold' => true,
                'italic' => false,
                'strikethrough' => false,
                'color' => [
                    'rgb' => '808080'
                ]
            ],
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => [
                        'rgb' => '808080'
                    ]
                ],
            ],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical' => Alignment::VERTICAL_CENTER,
                'wrapText' => true,
            ],

            'fill' => [
                'fillType' =>Fill::FILL_GRADIENT_LINEAR,
                'rotation' => 90,
                'startColor' => [
                    'argb' => 'FFEBCD',
                ],
                'endColor' => [
                    'argb' => 'FFEBCD',
                ],
            ],
        ]);
        $sheet->getStyle('C2:H2')->applyFromArray([
            'font' => [
                'name' => 'Arial',
                'bold' => true,
                'italic' => false,
                'strikethrough' => false,
                'color' => [
                    'rgb' => '8B4513'
                ]
            ],
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => [
                        'rgb' => '808080'
                    ]
                ],
            ],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical' => Alignment::VERTICAL_CENTER,
                'wrapText' => true,
            ],
            'fill' => [
                'fillType' =>Fill::FILL_GRADIENT_LINEAR,
                'rotation' => 90,
                'startColor' => [
                    'argb' => 'CD853F',
                ],
                'endColor' => [
                    'argb' => 'F5DEB3',
                ],
            ],
        ]);
        $sheet->getColumnDimension('H')->setWidth(20);
        $sheet->getRowDimension(2)->setRowHeight(50);


        $writer->save('phpspreadsheet.xlsx');
    }
}
