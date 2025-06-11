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

    public function removeExcel(Request $request)
    {
        $request->validate([
            'file_index' => 'required|integer|min:0',
        ]);

        $fileIndex = $request->input('file_index');
        $excelFiles = session('excel_files', []);

        if (!isset($excelFiles[$fileIndex])) {
            return back()->with('error', 'File không tồn tại trong danh sách.');
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

            $data = [];
            $statusColumnIndex = null;

            for ($row = 1; $row <= $highestRow; $row++) {
                $rowData = [];
                for ($col = 1; $col <= $highestColumnIndex; $col++) {
                    $cell = $worksheet->getCellByColumnAndRow($col, $row);
                    $value = $cell->getCalculatedValue();
                    $rowData[] = $value;

                    if ($row === 1 && strtolower($value) === 'status') {
                        $statusColumnIndex = $col - 1;
                    }
                }
                $data[] = $rowData;
            }

            if (empty($data)) {
                return redirect()->route('file.index')->with('error', 'Không tìm thấy dữ liệu trong sheet.');
            }

            return view('file_reader', [
                'excelFiles' => $excelFiles,
                'docFiles' => session('doc_files', []),
                'data' => $data,
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