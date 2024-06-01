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
          // xong header
        $startRow =
        $startRow + 2;
    //main
     // Nhóm theo san pham
     $groupedByItemNam = collect($dataTable)->groupBy('item_name');
    //  dd($groupedByItemNam);
     $finalGrouped = $groupedByItemNam->map(function ($items) {

        $groupedByProductId = $items->groupBy('date')->map(function ($products) {


            $dataMerge = array_merge($products[0], [
                //    'item_id' => $products->first()['item_id'],
                'quantity' => $products->sum('quantity'),
            ]);
            return $dataMerge;
        })->values()->sortBy('date'); // .sortBy() để sắp xếp theo product_id

        // Tính tổng qty cho mỗi ngày
        $totalQty = $groupedByProductId->sum('quantity');
        $totalQty = $groupedByProductId->sum('quantity');

        return [
            'date' => $groupedByProductId,
            'total_qty' => $totalQty,
            'extraItem' => [
                'item_name' => $groupedByProductId->first()['item_name'],
                'item_code' => $groupedByProductId->first()['item_code'],
                'unit_name' => $groupedByProductId->first()['unit_name'],
            ],
        ];
    });
    // dd($finalGrouped);
     // dua vao file excel
     $chi_muc = 1;
     foreach ($finalGrouped as $key => $each_product) {


        $activeWorksheet->setCellValue("A{$startRow}", $chi_muc);
        $activeWorksheet->setCellValue("B{$startRow}", $key);
        $activeWorksheet->setCellValue("C{$startRow}", $each_product['extraItem']['item_code']);
        $activeWorksheet->setCellValue("D{$startRow}", $each_product['extraItem']['unit_name']);
        for ($i = 0; $i < count($mang_dinh_danh_ngay_vs_cot); $i++) {
            $found = false;
            $column = chr(65 + $i + 4);
            foreach ($each_product['date'] as $key => $product) {
                if ($product['date'] == $mang_dinh_danh_ngay_vs_cot[$i]) {
                    $activeWorksheet->setCellValue($column . $startRow, $product['quantity']);
                    $found = true;
                    break;
                }
            }
            if (!$found) {
                $activeWorksheet->setCellValue($column . $startRow, 0);;
            }
        }
        // chen du lieu tong
        $columnTong = chr(65 + 3 + count($mang_dinh_danh_ngay_vs_cot));
        // dd($columnTong);
        $activeWorksheet->setCellValue("{$columnTong}{$startRow}", $each_product['total_qty']);

        ++$startRow;
        ++$chi_muc;
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
