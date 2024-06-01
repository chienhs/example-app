<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use File;
use Illuminate\Support\Facades\File as FacadesFile;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Font;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class aaaa extends Controller
{
    public function newexportTemplateKeNgang()
    {
        $dataTable = $this->getDataFromJson();
    $nameTemplate = 'BangKeNgang.xlsx';
        $templatePath = storage_path('template/BBNT/') . $nameTemplate;
        $spreadsheet = IOFactory::load($templatePath);
        $activeWorksheet = $spreadsheet->getActiveSheet();

        $ds_ngay = [];
        foreach ($dataTable as $key => $value) {
            $tmp['so_ct'] = rand(10,99);
            $tmp['date'] = $value['date'];
            array_push($ds_ngay, $tmp);
        }
        usort($ds_ngay, fn ($a, $b) => $a['date'] <=> $b['date']);
        $ds_ngay = collect($ds_ngay)->unique('date')->values()->toArray();
        // them phan tu vao dau
        $first_item = [
            'so_ct' =>  'STT',
            'date' =>  '',
        ];
        $second_item = [
            'so_ct' =>  'Tên chất thải',
            'date' =>  '',
        ];
        $third_item = [
            'so_ct' =>  'Mã CTNH',
            'date' =>  '',
        ];
        $fourth_item = [
            'so_ct' =>  'Số CT',
            'date' =>  'Đơn vị tính',
        ];
        array_unshift($ds_ngay, $first_item,$second_item,$third_item,$fourth_item);
        // them phan tu cuoi

        $last_item = [
            'so_ct' => '',
            'date' => 'Tổng khối lượng',
        ];
        array_push($ds_ngay, $last_item);
        // dd($ds_ngay);
         //
         $mang_dinh_danh_ngay_vs_cot = [];
         $startRow = 5;
         for ($i = 0; $i < count($ds_ngay); $i++) {
             $column = chr(65 + $i);
             $activeWorksheet->setCellValue($column . ($startRow + 0), $ds_ngay[$i]['so_ct']);
             $activeWorksheet->setCellValue($column . ($startRow + 1), $ds_ngay[$i]['date']);
             $pair = $ds_ngay[$i]['date'];
             if ($i > 3) {
                 array_push($mang_dinh_danh_ngay_vs_cot, $pair);
             }
         }
         $outputPath = 'bao-cao-ty-le-dap-ung.xlsx';
         $writer = new Xlsx($spreadsheet);
         $writer->save($outputPath);
         return response()->download($outputPath)->deleteFileAfterSend(true);

        // dd($data);
    }
    public function getDataFromJson()
    {
        // Đường dẫn tới tệp JSON, giả sử nó nằm trong cùng thư mục với controller
        $path = __DIR__ . '/sampleData.json';



        // Đọc nội dung của tệp JSON
        $json = FacadesFile::get($path);

        // Chuyển đổi JSON thành mảng PHP
        $data = json_decode($json, true);



        // Trả về dữ liệu dưới dạng JSON hoặc xử lý theo yêu cầu của bạn
        return  $data;
    }
}
