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
            $tmp['so_ct'] = rand(10, 99);
            $tmp['date'] = $value['date'];
            array_push($ds_ngay, $tmp);
        }
        usort($ds_ngay, fn ($a, $b) => $a['date'] <=> $b['date']);
        $ds_ngay = collect($ds_ngay)->unique('date')->values()->toArray();
        $ds_ngay_nguyen_ban = $ds_ngay;


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
        array_unshift($ds_ngay, $first_item, $second_item, $third_item, $fourth_item);
        // them phan tu cuoi

        $last_item = [
            'so_ct' => '',
            'date' => 'Tổng khối lượng',
        ];
        $last_item_gia_xy_ly = [
            'so_ct' => '[FCC] trả Thuận Thành	',
            'one_column_merge' => true,
            'date' => 'Đơn giá xử lý (VNĐ)',
        ];
        $last_item_tien_xy_ly = [
            'so_ct' => '',
            'one_column_merge' => true,
            'date' => 'Thành tiền (VNĐ)',
        ];
        $last_item_gia_thu_mua = [
            'so_ct' => 'Thuận Thành trả [FCC]	',
            'one_column_merge' => true,
            'date' => ' Đơn giá thu mua đã bao gồm VAT (VNĐ) ',
        ];
        $last_item_tien_thu_mua = [
            'so_ct' => '',
            'one_column_merge' => true,
            'date' => 'Thành tiền (VNĐ)',
        ];
        array_push($ds_ngay, $last_item, $last_item_gia_xy_ly, $last_item_tien_xy_ly, $last_item_gia_thu_mua, $last_item_tien_thu_mua);
        // dd($ds_ngay);
        //
        $mang_dinh_danh_ngay_vs_cot = [];


        $startRow = 5;
        for ($i = 0; $i < count($ds_ngay); $i++) {
            $column = chr(65 + $i);
            $activeWorksheet->setCellValue($column . ($startRow + 0), $ds_ngay[$i]['so_ct']);
            $activeWorksheet->setCellValue($column . ($startRow + 1), $ds_ngay[$i]['date']);
            $pair = $ds_ngay[$i]['date'];
            if ($i > 3  && $i <= 3 + count($ds_ngay_nguyen_ban)) {
                array_push($mang_dinh_danh_ngay_vs_cot, $pair);
            }
            if ($i > 3 + count($ds_ngay_nguyen_ban)) {
                //style kieu khac
                $activeWorksheet->setCellValue($column . ($startRow + 0), $ds_ngay[$i]['so_ct']);
                $activeWorksheet->setCellValue($column . ($startRow + 1), $ds_ngay[$i]['date']);
            }
        }
        // dd($mang_dinh_danh_ngay_vs_cot);

        // xong header
        $startRow =
            $startRow + 2;
        //main
        // Nhóm theo san pham
        $tong_thue_va_tien_thu_mua_va_xu_ly = [
            'total_vat_all' => 0,
            'total_product_price_all' => 0,
            'total_purchase_price_all' => 0,
        ];
        $groupedByItemNam = collect($dataTable)->groupBy('item_name');
        //  dd($groupedByItemNam);
        $finalGrouped = $groupedByItemNam->map(function ($items) use (&$tong_thue_va_tien_thu_mua_va_xu_ly) {

            $groupedByProductId = $items->groupBy('date')->map(function ($products) {
                $dataMerge = array_merge($products[0], [
                    //    'item_id' => $products->first()['item_id'],
                    'quantity' => $products->sum('quantity'),
                ]);
                return $dataMerge;
            })->values()->sortBy('date'); // .sortBy() để sắp xếp theo product_id
            // Tính tổng qty cho mỗi ngày
            $totalQty = $groupedByProductId->sum('quantity');
            ////////////////
            $production_unit_price = $groupedByProductId->first()['production_unit_price'];
            $production_unit_total = $groupedByProductId->first()['production_unit_price'] * $totalQty;
            $purchasing_unit_price = $groupedByProductId->first()['purchasing_unit_price'];
            $purchasing_unit_total = $groupedByProductId->first()['purchasing_unit_price'] * $totalQty;
            ///////////
            $totalVat = $groupedByProductId->sum('total_vat');
            $tong_thue_va_tien_thu_mua_va_xu_ly['total_vat_all'] += $totalVat;
            $tong_thue_va_tien_thu_mua_va_xu_ly['total_product_price_all'] += $production_unit_total;
            $tong_thue_va_tien_thu_mua_va_xu_ly['total_purchase_price_all'] += $purchasing_unit_total;
            return [
                'date' => $groupedByProductId,
                'total_qty' => $totalQty,
                'production_unit_price' => $production_unit_price,
                'production_unit_total' => $production_unit_total,
                'purchasing_unit_price' => $purchasing_unit_price,
                'purchasing_unit_total' => $purchasing_unit_total,
                'totalVat' => $totalVat,
                'extraItem' => [
                    'item_name' => $groupedByProductId->first()['item_name'],
                    'item_code' => $groupedByProductId->first()['item_code'],
                    'unit_name' => $groupedByProductId->first()['unit_name'],
                ],
            ];
        });
        // dd($finalGrouped);
        // dd($finalGrouped);
        // dua vao file excel

        $chi_muc = 1;
        foreach ($finalGrouped as $key => $each_product) {

            // dd($mang_dinh_danh_ngay_vs_cot);

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
            $columnTong = chr(65 + 4 + count($mang_dinh_danh_ngay_vs_cot));

            $activeWorksheet->setCellValue("{$columnTong}{$startRow}", $each_product['total_qty']);
            // dd($columnTong);
            $columnTong = chr(65 + 5 + count($mang_dinh_danh_ngay_vs_cot));
            $activeWorksheet->setCellValue("{$columnTong}{$startRow}", $each_product['production_unit_price']);
            // dd($columnTong);
            $columnTong = chr(65 + 6 + count($mang_dinh_danh_ngay_vs_cot));
            $activeWorksheet->setCellValue("{$columnTong}{$startRow}", $each_product['production_unit_total']);

            $columnTong = chr(65 + 7 + count($mang_dinh_danh_ngay_vs_cot));
            // dd($columnTong);
            $activeWorksheet->setCellValue("{$columnTong}{$startRow}", $each_product['purchasing_unit_price']);

            $columnTong = chr(65 + 8 + count($mang_dinh_danh_ngay_vs_cot));
            // dd($columnTong);
            $activeWorksheet->setCellValue("{$columnTong}{$startRow}", $each_product['purchasing_unit_total']);

            ++$startRow;
            ++$chi_muc;
        }
        // footer end
        $ds_tien_thanh_toan = [];
        $last_item_one  = [
            'sub_total' => 'Cộng',
            'vat_total' => ' Thuế VAT',
            'total' => 'Tổng cộng',
        ];

        $last_item_two  = [
            'sub_total' => $tong_thue_va_tien_thu_mua_va_xu_ly['total_product_price_all'],
            'vat_total' => $tong_thue_va_tien_thu_mua_va_xu_ly['total_vat_all'],
            'total' => $tong_thue_va_tien_thu_mua_va_xu_ly['total_product_price_all'] * $tong_thue_va_tien_thu_mua_va_xu_ly['total_vat_all'],

        ];
        $last_item_three  = [
            'sub_total' => $tong_thue_va_tien_thu_mua_va_xu_ly['total_purchase_price_all'],
            'vat_total' => $tong_thue_va_tien_thu_mua_va_xu_ly['total_vat_all'],
            'total' => $tong_thue_va_tien_thu_mua_va_xu_ly['total_purchase_price_all'] * $tong_thue_va_tien_thu_mua_va_xu_ly['total_vat_all'],
        ];
        array_push($ds_tien_thanh_toan, $last_item_one, $last_item_two, $last_item_three);
        for ($i = 0; $i < count($ds_tien_thanh_toan); $i++) {
            if ($i == 1) {
                $columnTong = chr(65 + 6 + count($ds_ngay_nguyen_ban));
                $activeWorksheet->setCellValue($columnTong . ($startRow + 0), $ds_tien_thanh_toan[$i]['sub_total']);
                $activeWorksheet->setCellValue($columnTong . ($startRow + 1), $ds_tien_thanh_toan[$i]['vat_total']);
                $activeWorksheet->setCellValue($columnTong . ($startRow + 2), $ds_tien_thanh_toan[$i]['total']);
            } elseif ($i == 2) {
                $columnTong = chr(65 + 8 + count($ds_ngay_nguyen_ban));
                $activeWorksheet->setCellValue($columnTong . ($startRow + 0), $ds_tien_thanh_toan[$i]['sub_total']);
                $activeWorksheet->setCellValue($columnTong . ($startRow + 1), $ds_tien_thanh_toan[$i]['vat_total']);
                $activeWorksheet->setCellValue($columnTong . ($startRow + 2), $ds_tien_thanh_toan[$i]['total']);
            } else {
                $column = 'A';
                $activeWorksheet->setCellValue($column . ($startRow + 0), $ds_tien_thanh_toan[$i]['sub_total']);
                $activeWorksheet->setCellValue($column . ($startRow + 1), $ds_tien_thanh_toan[$i]['vat_total']);
                $activeWorksheet->setCellValue($column . ($startRow + 2), $ds_tien_thanh_toan[$i]['total']);
            }
        }
        // tong thue va tien cua xu ly va thu mua
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
