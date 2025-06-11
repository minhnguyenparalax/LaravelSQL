<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;

class FileController extends Controller
{
    public function index(Request $request)
    {
        $excelFiles = session('excel_files', []);
        $docFiles = session('doc_files', []);

        return view('file_reader', [
            'excelFiles' => $excelFiles,
            'docFiles' => $docFiles,
        ]);
    }
}