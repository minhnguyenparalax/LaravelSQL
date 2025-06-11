<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel & Doc Reader</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        .status-column {
            font-weight: bold;
        }
        .table-smaller {
            font-size: 0.9rem;
        }
        .table-smaller th, .table-smaller td {
            padding: 0.5rem;
            line-height: 1.2;
        }
        .table-smaller th {
            background-color: #f1f1f1;
        }
        .sheet-list {
            margin-left: 1rem;
        }
        .doc-content {
            border: 1px solid #dee2e6;
            padding: 1rem;
            margin-top: 1rem;
            background-color: #f8f9fa;
            white-space: pre-wrap; /* Giữ khoảng trắng và cách dòng */
        }
        .doc-content p {
            margin: 0 0 0.5rem 0; /* Giữ khoảng cách đoạn văn */
        }
        .doc-content table {
            width: 100%;
            margin-bottom: 1rem;
        }
        .doc-content table th, .doc-content table td {
            padding: 0.5rem;
            vertical-align: top;
        }
    </style>
</head>
<body>
    <div class="container mt-5">
        <h2>Đọc File Excel</h2>
        
        <!-- Form thêm Excel -->
        <form action="{{ route('excel.addExcel') }}" method="POST" class="mb-3">
            @csrf
            <div class="mb-3">
                <label for="excel_file_path" class="form-label">Nhập Đường Dẫn File Excel (.xlsx hoặc .xls)</label>
                <input type="text" class="form-control" id="excel_file_path" name="file_path" 
                       placeholder="VD: F:\00 Template\2025 Project Management.xlsx" 
                       value="{{ old('file_path') }}" required>
                <small class="form-text text-muted">Đường dẫn có thể có hoặc không có dấu ngoặc kép</small>
            </div>
            <button type="submit" class="btn btn-primary">Thêm Excel</button>
        </form>

        <h2>Đọc File Doc</h2>
        
        <!-- Form thêm Doc -->
        <form action="{{ route('doc.addDoc') }}" method="POST" class="mb-3">
            @csrf
            <div class="mb-3">
                <label for="doc_file_path" class="form-label">Nhập Đường Dẫn File Doc (.doc hoặc .docx)</label>
                <input type="text" class="form-control" id="doc_file_path" name="file_path" 
                       placeholder="VD: F:\Documents\Report.docx" 
                       value="{{ old('file_path') }}" required>
                <small class="form-text text-muted">Đường dẫn có thể có hoặc không có dấu ngoặc kép</small>
            </div>
            <button type="submit" class="btn btn-primary">Thêm Doc</button>
        </form>

        <!-- Thông báo lỗi -->
        @if (session('error'))
            <div class="alert alert-danger mt-3">
                {{ session('error') }}
            </div>
        @endif

        <!-- Thông báo thành công -->
        @if (session('success'))
            <div class="alert alert-success mt-3">
                {{ session('success') }}
            </div>
        @endif

        <!-- Danh sách Excel -->
        @if (!empty($excelFiles))
            <h3 class="mt-5">Danh Sách Excel</h3>
            <ul class="list-group mb-3">
                @foreach ($excelFiles as $fileIndex => $file)
                    <li class="list-group-item d-flex justify-content-between align-items-start">
                        <div>
                            <strong>{{ $file['name'] }}</strong>
                            <ul class="sheet-list">
                                @foreach ($file['sheets'] as $sheetIndex => $sheetName)
                                    <li>
                                        <a href="{{ route('excel.readSheet', [$fileIndex, $sheetIndex]) }}"
                                           class="text-decoration-none {{ isset($currentFileIndex) && isset($currentSheetIndex) && $currentFileIndex == $fileIndex && $currentSheetIndex == $sheetIndex ? 'fw-bold' : '' }}">
                                            {{ $sheetName }}
                                        </a>
                                    </li>
                                @endforeach
                            </ul>
                        </div>
                        <form action="{{ route('excel.removeExcel') }}" method="POST">
                            @csrf
                            <input type="hidden" name="file_index" value="{{ $fileIndex }}">
                            <button type="submit" class="btn btn-danger btn-sm">Xóa Excel</button>
                        </form>
                    </li>
                @endforeach
            </ul>
        @endif

        <!-- Danh sách Doc -->
        @if (!empty($docFiles))
            <h3 class="mt-5">Danh Sách Doc</h3>
            <ul class="list-group mb-3">
                @foreach ($docFiles as $docIndex => $doc)
                    <li class="list-group-item d-flex justify-content-between align-items-center">
                        <a href="{{ route('doc.readDoc', $docIndex) }}"
                           class="text-decoration-none {{ isset($currentDocIndex) && $currentDocIndex == $docIndex ? 'fw-bold' : '' }}">
                            {{ $doc['name'] }}
                        </a>
                        <form action="{{ route('doc.removeDoc') }}" method="POST">
                            @csrf
                            <input type="hidden" name="doc_index" value="{{ $docIndex }}">
                            <button type="submit" class="btn btn-danger btn-sm">Xóa Doc</button>
                        </form>
                    </li>
                @endforeach
            </ul>
        @endif

        <!-- Dữ liệu Excel -->
        @if (isset($data) && !empty($data))
            <h3 class="mt-5">Dữ Liệu: {{ $excelFiles[$currentFileIndex]['name'] }} / {{ $excelFiles[$currentFileIndex]['sheets'][$currentSheetIndex] }}</h3>
            <table class="table table-bordered mt-3 table-smaller">
                <thead>
                    <tr>
                        @foreach ($data[0] as $index => $header)
                            <th class="{{ isset($statusColumnIndex) && $index == $statusColumnIndex ? 'status-column' : '' }}">
                                {{ $header ?? 'Cột ' . ($index + 1) }}
                            </th>
                        @endforeach
                    </tr>
                </thead>
                <tbody>
                    @foreach (array_slice($data, 1) as $row)
                        <tr>
                            @foreach ($row as $colIndex => $cell)
                                <td class="{{ isset($statusColumnIndex) && $colIndex == $statusColumnIndex ? 'status-column' : '' }}">
                                    {{ $cell ?? '' }}
                                </td>
                            @endforeach
                        </tr>
                    @endforeach
                </tbody>
            </table>
        @endif

        <!-- Nội dung Doc -->
        @if (isset($docContent))
            <h3 class="mt-5">Nội Dung: {{ $docFiles[$currentDocIndex]['name'] }}</h3>
            <div class="doc-content">
                {!! $docContent !!}
            </div>
        @endif
    </div>
</body>
</html>