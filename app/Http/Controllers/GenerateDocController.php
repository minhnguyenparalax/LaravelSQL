<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use PhpOffice\PhpSpreadsheet\IOFactory as SpreadsheetIOFactory;
use PhpOffice\PhpWord\IOFactory as WordIOFactory;
use PhpOffice\PhpWord\PhpWord;
use Illuminate\Support\Facades\Log;
use Illuminate\Support\Facades\Storage;

class GenerateDocController extends Controller
{
    public function setPrimaryKey(Request $request)
    {
        $docIndex = $request->input('doc_index');
        $variable = $request->input('variable');

        $docFiles = session('doc_files', []);
        $docVariables = session('doc_variables', []);

        if (!isset($docFiles[$docIndex]) || !isset($docVariables[$docIndex])) {
            return redirect()->route('file.index')->with('error', 'File hoặc danh sách biến không tồn tại.');
        }

        // Kiểm tra biến có tồn tại trong danh sách biến của Doc
        if (!in_array($variable, $docVariables[$docIndex]['variables'])) {
            return redirect()->route('file.index')->with('error', 'Biến không tồn tại trong file Doc.');
        }

        // Cập nhật khóa chính trong session
        $docVariables[$docIndex]['primary_key'] = $variable;
        session(['doc_variables' => $docVariables]);

        return redirect()->route('file.index')->with('success', 'Đã đặt biến "' . $variable . '" làm khóa chính cho file "' . $docFiles[$docIndex]['name'] . '".');
    }

    public function setOutputFolder(Request $request)
    {
        $request->validate([
            'output_folder' => 'required|string',
        ]);

        $outputFolder = trim($request->input('output_folder'), '"\'');
        $outputFolder = str_replace('/', DIRECTORY_SEPARATOR, $outputFolder);

        if (!is_dir($outputFolder)) {
            return redirect()->route('file.index')->with('error', 'Thư mục đầu ra không tồn tại: ' . $outputFolder);
        }

        session(['output_folder' => $outputFolder]);

        return redirect()->route('file.index')->with('success', 'Đã đặt thư mục đầu ra: ' . $outputFolder);
    }

    public function generateDoc(Request $request, $docIndex)
    {
        $docFiles = session('doc_files', []);
        $excelFiles = session('excel_files', []);
        $mappings = session('mappings', []);
        $docVariables = session('doc_variables', []);
        $outputFolder = session('output_folder');

        // Kiểm tra file và biến
        if (!isset($docFiles[$docIndex])) {
            return redirect()->route('file.index')->with('error', 'File Doc không tồn tại.');
        }
        if (!isset($docVariables[$docIndex]['primary_key'])) {
            return redirect()->route('file.index')->with('error', 'Chưa đặt khóa chính cho file Doc.');
        }
        if (empty($outputFolder) || !is_dir($outputFolder)) {
            return redirect()->route('file.index')->with('error', 'Chưa đặt thư mục đầu ra hoặc thư mục không tồn tại.');
        }

        $primaryKey = $docVariables[$docIndex]['primary_key'];

        // Tìm mapping của biến khóa chính
        $primaryMapping = collect($mappings)->firstWhere(function ($mapping) use ($docIndex, $primaryKey) {
            return $mapping['doc_index'] == $docIndex && $mapping['variable'] == $primaryKey;
        });

        if (!$primaryMapping) {
            return redirect()->route('file.index')->with('error', 'Biến khóa chính "' . $primaryKey . '" chưa được mapping.');
        }

        $fileIndex = $primaryMapping['field']['file_index'];
        $sheetIndex = $primaryMapping['field']['sheet_index'];
        $field = $primaryMapping['field']['field'];

        if (!isset($excelFiles[$fileIndex])) {
            return redirect()->route('file.index')->with('error', 'File Excel không tồn tại.');
        }

        $filePath = $excelFiles[$fileIndex]['path'];
        if (!file_exists($filePath)) {
            return redirect()->route('file.index')->with('error', 'File Excel không tồn tại tại: ' . $filePath);
        }

        try {
            // Đọc file Excel
            $spreadsheet = SpreadsheetIOFactory::load($filePath);
            $worksheet = $spreadsheet->getSheet($sheetIndex);
            $highestRow = $worksheet->getHighestRow();
            $highestColumn = $worksheet->getHighestColumn();
            $highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn);

            // Tìm cột của trường khóa chính
            $fieldColumnIndex = null;
            for ($col = 1; $col <= $highestColumnIndex; $col++) {
                $cell = $worksheet->getCellByColumnAndRow($col, 1);
                if ($cell->getCalculatedValue() === $field) {
                    $fieldColumnIndex = $col;
                    break;
                }
            }

            if ($fieldColumnIndex === null) {
                return redirect()->route('file.index')->with('error', 'Không tìm thấy trường "' . $field . '" trong sheet.');
            }

            // Đọc giá trị từ cột khóa chính
            $values = [];
            for ($row = 2; $row <= $highestRow; $row++) {
                $cell = $worksheet->getCellByColumnAndRow($fieldColumnIndex, $row);
                $value = $cell->getCalculatedValue();
                if (!is_null($value) && trim($value) !== '') {
                    $values[] = $value;
                }
            }

            if (empty($values)) {
                return redirect()->route('file.index')->with('error', 'Không tìm thấy giá trị hợp lệ trong cột "' . $field . '".');
            }

            // Đọc file Doc gốc
            $docPath = $docFiles[$docIndex]['path'];
            $phpWord = WordIOFactory::load($docPath, 'Word2007');
            $generatedDocs = [];

            // Tạo file Doc mới cho mỗi giá trị
            foreach ($values as $index => $value) {
                $newPhpWord = new PhpWord();
                foreach ($phpWord->getSections() as $section) {
                    $newSection = $newPhpWord->addSection($section->getStyle());
                    foreach ($section->getElements() as $element) {
                        if ($element instanceof \PhpOffice\PhpWord\Element\TextRun) {
                            $newTextRun = $newSection->addTextRun($element->getParagraphStyle());
                            foreach ($element->getElements() as $subElement) {
                                if ($subElement instanceof \PhpOffice\PhpWord\Element\Text) {
                                    $text = $subElement->getText();
                                    $fontStyle = $subElement->getFontStyle();
                                    // Thay thế biến khóa chính
                                    $newText = str_replace('{{' . $primaryKey . '}}', $value, $text);
                                    // Thay thế các biến khác
                                    foreach ($mappings as $mapping) {
                                        if ($mapping['doc_index'] == $docIndex) {
                                            $newText = str_replace('{{' . $mapping['variable'] . '}}', $this->getFieldValue($excelFiles, $mapping, $row - 1, $worksheet), $newText);
                                        }
                                    }
                                    $newTextRun->addText($newText, $fontStyle);
                                } else {
                                    $newSection->addElement($subElement);
                                }
                            }
                        } elseif ($element instanceof \PhpOffice\PhpWord\Element\Table) {
                            $newTable = $newSection->addTable($element->getStyle());
                            foreach ($element->getRows() as $row) {
                                $newRow = $newTable->addRow();
                                foreach ($row->getCells() as $cell) {
                                    $newCell = $newRow->addCell($cell->getWidth(), $cell->getStyle());
                                    foreach ($cell->getElements() as $cellElement) {
                                        if ($cellElement instanceof \PhpOffice\PhpWord\Element\TextRun) {
                                            $newTextRun = $newCell->addTextRun($cellElement->getParagraphStyle());
                                            foreach ($cellElement->getElements() as $subElement) {
                                                if ($subElement instanceof \PhpOffice\PhpWord\Element\Text) {
                                                    $text = $subElement->getText();
                                                    $fontStyle = $subElement->getFontStyle();
                                                    $newText = str_replace('{{' . $primaryKey . '}}', $value, $text);
                                                    foreach ($mappings as $mapping) {
                                                        if ($mapping['doc_index'] == $docIndex) {
                                                            $newText = str_replace('{{' . $mapping['variable'] . '}}', $this->getFieldValue($excelFiles, $mapping, $row - 1, $worksheet), $newText);
                                                        }
                                                    }
                                                    $newTextRun->addText($newText, $fontStyle);
                                                } else {
                                                    $newCell->addElement($subElement);
                                                }
                                            }
                                        } else {
                                            $newCell->addElement($cellElement);
                                        }
                                    }
                                }
                            }
                        } else {
                            $newSection->addElement($element);
                        }
                    }
                }

                // Lưu file tạm vào session
                $filename = pathinfo($docFiles[$docIndex]['name'], PATHINFO_FILENAME) . '_' . ($index + 1) . '.docx';
                $tempPath = storage_path('app/temp/' . $filename);
                $writer = WordIOFactory::createWriter($newPhpWord, 'Word2007');
                $writer->save($tempPath);

                $generatedDocs[] = [
                    'filename' => $filename,
                    'path' => $tempPath,
                ];
            }

            // Lưu file ra thư mục đầu ra
            foreach ($generatedDocs as $doc) {
                $destinationPath = rtrim($outputFolder, DIRECTORY_SEPARATOR) . DIRECTORY_SEPARATOR . $doc['filename'];
                copy($doc['path'], $destinationPath);
            }

            // Lưu danh sách file đã tạo vào session
            $generatedDocFiles = session('generated_doc_files', []);
            $generatedDocFiles[$docIndex] = $generatedDocs;
            session(['generated_doc_files' => $generatedDocFiles]);

            return redirect()->route('file.index')->with('success', 'Đã tạo ' . count($generatedDocs) . ' file Doc mới từ "' . $docFiles[$docIndex]['name'] . '" và lưu vào "' . $outputFolder . '".');

        } catch (\Exception $e) {
            Log::error('Lỗi khi tạo file Doc: ' . $e->getMessage());
            return redirect()->route('file.index')->with('error', 'Không thể tạo file Doc: ' . $e->getMessage());
        }
    }

    private function getFieldValue($excelFiles, $mapping, $rowIndex, $worksheet)
    {
        $fileIndex = $mapping['field']['file_index'];
        $sheetIndex = $mapping['field']['sheet_index'];
        $field = $mapping['field']['field'];

        $highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($worksheet->getHighestColumn());

        for ($col = 1; $col <= $highestColumnIndex; $col++) {
            $cell = $worksheet->getCellByColumnAndRow($col, 1);
            if ($cell->getCalculatedValue() === $field) {
                $cellValue = $worksheet->getCellByColumnAndRow($col, $rowIndex + 2)->getCalculatedValue();
                return is_null($cellValue) ? '' : $cellValue;
            }
        }

        return '';
    }
}