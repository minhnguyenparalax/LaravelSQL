<?php

use Illuminate\Support\Facades\Route;
use App\Http\Controllers\FileController;
use App\Http\Controllers\ExcelController;
use App\Http\Controllers\DocController;

Route::get('/', [FileController::class, 'index'])->name('file.index');
Route::post('/add-excel', [ExcelController::class, 'addExcel'])->name('excel.addExcel');
Route::post('/remove-excel', [ExcelController::class, 'removeExcel'])->name('excel.removeExcel');
Route::get('/read-sheet/{fileIndex}/{sheetIndex}', [ExcelController::class, 'readSheet'])->name('excel.readSheet');

//Thêm doc
Route::post('/add-doc', [DocController::class, 'addDoc'])->name('doc.addDoc');
Route::post('/remove-doc', [DocController::class, 'removeDoc'])->name('doc.removeDoc');
Route::get('/read-doc/{docIndex}', [DocController::class, 'readDoc'])->name('doc.readDoc');