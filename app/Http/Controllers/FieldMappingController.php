<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;

class FieldMappingController extends Controller
{
    public function mapVariable(Request $request)
    {
        $docIndex = $request->input('doc_index');
        $variable = $request->input('variable');
        $fileIndex = $request->input('file_index');
        $sheetIndex = $request->input('sheet_index');
        $field = $request->input('field');

        $docFiles = session('doc_files', []);
        $excelFiles = session('excel_files', []);

        // Kiểm tra file tồn tại
        if (!isset($docFiles[$docIndex]) || !isset($excelFiles[$fileIndex])) {
            return redirect()->back()->with('error', 'File không tồn tại.');
        }

        $mappings = session('mappings', []);

        // Kiểm tra 1:1 trong phạm vi cùng doc_index
        foreach ($mappings as $mapping) {
            // Kiểm tra biến đã được mapping trong cùng doc_index
            if ($mapping['doc_index'] == $docIndex && $mapping['variable'] == $variable) {
                return redirect()->back()->with('error', 'Biến "' . $variable . '" đã được mapping trong báo cáo "' . $docFiles[$docIndex]['name'] . '".');
            }
            // Kiểm tra trường đã được mapping trong cùng doc_index
            if ($mapping['doc_index'] == $docIndex &&
                $mapping['field']['file_index'] == $fileIndex && 
                $mapping['field']['sheet_index'] == $sheetIndex && 
                $mapping['field']['field'] == $field) {
                return redirect()->back()->with('error', 'Trường "' . $field . '" đã được mapping trong báo cáo "' . $docFiles[$docIndex]['name'] . '".');
            }
        }

        // Lưu mapping
        $mappings[] = [
            'doc_index' => $docIndex,
            'variable' => $variable,
            'field' => [
                'file_index' => $fileIndex,
                'sheet_index' => $sheetIndex,
                'field' => $field,
            ],
        ];
        session(['mappings' => $mappings]);

        return redirect()->back()->with('success', 'Đã mapping biến "' . $variable . '" với trường "' . $field . '" trong báo cáo "' . $docFiles[$docIndex]['name'] . '".');
    }

    public function removeMapping(Request $request)
    {
        $docIndex = $request->input('doc_index');
        $variable = $request->input('variable');

        $docFiles = session('doc_files', []);
        $mappings = session('mappings', []);

        // Kiểm tra file tồn tại
        if (!isset($docFiles[$docIndex])) {
            return redirect()->back()->with('error', 'File không tồn tại.');
        }

        // Lọc xóa mapping
        $mappings = array_filter($mappings, fn($mapping) => 
            !($mapping['doc_index'] == $docIndex && $mapping['variable'] == $variable)
        );

        session(['mappings' => array_values($mappings)]);

        return redirect()->back()->with('success', 'Đã xóa mapping của biến "' . $variable . '" trong báo cáo "' . $docFiles[$docIndex]['name'] . '".');
    }
}