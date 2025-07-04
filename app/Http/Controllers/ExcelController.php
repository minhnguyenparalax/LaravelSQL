<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use PhpOffice\PhpSpreadsheet\IOFactory;
use Illuminate\Support\Facades\Log;

class ExcelController extends Controller
{
    public function addExcel(Request $request)
    {
        $request->validate([
            'file_path' => 'required|string',
        ]);

        $filePath = trim($request->input('file_path'), '"\'');
        $filePath = str_replace('/', '\\', $filePath);

        if (!file_exists($filePath)) {
            return back()->with('error', 'File không tồn tại tại: ' . $filePath);
        }

        $allowedExtensions = ['xlsx', 'xls'];
        $extension = strtolower(pathinfo($filePath, PATHINFO_EXTENSION));
        if (!in_array($extension, $allowedExtensions)) {
            return back()->with('error', 'Định dạng file không hợp lệ. Chỉ hỗ trợ .xlsx hoặc .xls.');
        }

        try {
            $spreadsheet = IOFactory::load($filePath);
            if (!$spreadsheet->getSheetCount()) {
                return back()->with('error', 'File rỗng hoặc không hợp lệ.');
            }

            $sheetNames = $spreadsheet->getSheetNames();
            $excelFiles = session('excel_files', []);
            $excelFiles[] = [
                'path' => $filePath,
                'name' => basename($filePath),
                'sheets' => $sheetNames,
            ];

            session(['excel_files' => $excelFiles]);

            return redirect()->route('file.index')->with('success', 'Đã thêm file Excel thành công: ' . $filePath);

        } catch (\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
            Log::error('Lỗi khi đọc file Excel: ' . $e->getMessage());
            return back()->with('error', 'Không thể đọc file Excel: Định dạng sai hoặc file hỏng.');
        } catch (\Exception $e) {
            Log::error('Lỗi hệ thống: ' . $e->getMessage());
            return back()->with('error', 'Lỗi không xác định: ' . $e->getMessage());
        }
    }

    public function removeExcel($fileIndex)
    {
        $excelFiles = session('excel_files', []);

        if (!isset($excelFiles[$fileIndex])) {
            return redirect()->route('file.index')->with('error', 'File không tồn tại trong danh sách.');
        }

        $filePath = $excelFiles[$fileIndex]['path'];
        unset($excelFiles[$fileIndex]);
        $excelFiles = array_values($excelFiles);

        session(['excel_files' => $excelFiles]);

        return redirect()->route('file.index')->with('success', 'Đã xóa file Excel: ' . $filePath);
    }

    public function readSheet($fileIndex, $sheetIndex)
    {
        $excelFiles = session('excel_files', []);

        if (!isset($excelFiles[$fileIndex])) {
            return redirect()->route('file.index')->with('error', 'File không tồn tại trong danh sách.');
        }

        $filePath = $excelFiles[$fileIndex]['path'];
        if (!file_exists($filePath)) {
            return redirect()->route('file.index')->with('error', 'File không tồn tại tại: ' . $filePath);
        }

        try {
            $spreadsheet = IOFactory::load($filePath);
            $sheetNames = $spreadsheet->getSheetNames();

            if (!isset($sheetNames[$sheetIndex])) {
                return redirect()->route('file.index')->with('error', 'Sheet không tồn tại.');
            }

            $worksheet = $spreadsheet->getSheet($sheetIndex);
            $highestRow = $worksheet->getHighestRow();
            $highestColumn = $worksheet->getHighestColumn();
            $highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn);

            // Lấy thông tin merged cells
            $mergeCells = $worksheet->getMergeCells();
            $mergeInfo = [];
            foreach ($mergeCells as $mergeRange) {
                [$startCell, $endCell] = explode(':', $mergeRange);
                [$startCol, $startRow] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($startCell);
                [$endCol, $endRow] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($endCell);
                $startColIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($startCol);
                $endColIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($endCol);
                $colspan = $endColIndex - $startColIndex + 1;
                $mergeInfo[$startRow][$startColIndex] = $colspan;
            }

            $data = [];
            $statusColumnIndex = null;

            // Đọc tất cả ô từ sheet
            for ($row = 1; $row <= $highestRow; $row++) {
                $rowData = [];
                for ($col = 1; $col <= $highestColumnIndex; $col++) {
                    // Kiểm tra ô có thuộc vùng gộp không
                    $colspan = isset($mergeInfo[$row][$col]) ? $mergeInfo[$row][$col] : 1;
                    $cell = $worksheet->getCellByColumnAndRow($col, $row);
                    $value = $cell->getCalculatedValue();
                    $value = is_null($value) ? '' : $value; // Chuyển null thành chuỗi rỗng
                    $rowData[$col - 1] = [
                        'value' => $value,
                        'colspan' => $colspan,
                    ];

                    // Tìm cột Status ở hàng đầu tiên
                    if ($row === 1 && strtolower($value) === 'status') {
                        $statusColumnIndex = $col - 1;
                    }

                    // Bỏ qua các cột đã gộp
                    if ($colspan > 1) {
                        for ($i = 1; $i < $colspan; $i++) {
                            $rowData[$col - 1 + $i] = ['value' => '', 'colspan' => 0]; // Đánh dấu cột gộp phụ
                        }
                        $col += $colspan - 1;
                    }
                }
                $data[] = $rowData;
            }

            if (empty($data)) {
                return redirect()->route('file.index')->with('error', 'Không tìm thấy dữ liệu trong sheet.');
            }

            // Tìm hàng cuối cùng có giá trị
            $lastNonEmptyRowIndex = 0;
            foreach ($data as $rowIndex => $rowData) {
                foreach ($rowData as $cell) {
                    if ($cell['value'] !== '' && trim($cell['value']) !== '') {
                        $lastNonEmptyRowIndex = max($lastNonEmptyRowIndex, $rowIndex);
                    }
                }
            }

            // Lọc bỏ các cột hoàn toàn trống
            $nonEmptyColumns = [];
            for ($col = 0; $col < $highestColumnIndex; $col++) {
                $hasData = false;
                foreach ($data as $rowIndex => $rowData) {
                    if ($rowIndex > $lastNonEmptyRowIndex) {
                        continue; // Bỏ qua các hàng sau hàng cuối cùng có giá trị
                    }
                    if (isset($rowData[$col]) && $rowData[$col]['value'] !== '' && $rowData[$col]['colspan'] !== 0) {
                        $hasData = true;
                        break;
                    }
                }
                if ($hasData) {
                    $nonEmptyColumns[] = $col;
                }
            }

            // Tạo dữ liệu mới chỉ chứa các cột không trống và các hàng đến lastNonEmptyRowIndex
            $filteredData = [];
            foreach ($data as $rowIndex => $rowData) {
                if ($rowIndex > $lastNonEmptyRowIndex) {
                    continue; // Bỏ qua các hàng sau hàng cuối cùng có giá trị
                }
                $filteredRow = [];
                foreach ($nonEmptyColumns as $col) {
                    $filteredRow[] = $rowData[$col] ?? ['value' => '', 'colspan' => 1];
                }
                $filteredData[] = $filteredRow;
            }

            // Cập nhật statusColumnIndex cho dữ liệu đã lọc
            if ($statusColumnIndex !== null) {
                $statusColumnIndex = array_search($statusColumnIndex, $nonEmptyColumns);
            }

            return view('file_reader', [
                'excelFiles' => $excelFiles,
                'docFiles' => session('doc_files', []),
                'data' => $filteredData,
                'statusColumnIndex' => $statusColumnIndex,
                'currentFileIndex' => $fileIndex,
                'currentSheetIndex' => $sheetIndex,
                'success' => 'Đã đọc sheet "' . $sheetNames[$sheetIndex] . '" từ file "' . $excelFiles[$fileIndex]['name'] . '" thành công.'
            ]);

        } catch (\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
            Log::error('Lỗi khi đọc sheet: ' . $e->getMessage());
            return redirect()->route('file.index')->with('error', 'Không thể đọc sheet: Định dạng sai hoặc sheet hỏng.');
        } catch (\Exception $e) {
            Log::error('Lỗi hệ thống: ' . $e->getMessage());
            return redirect()->route('file.index')->with('error', 'Lỗi không xác định: ' . $e->getMessage());
        }
    }
}