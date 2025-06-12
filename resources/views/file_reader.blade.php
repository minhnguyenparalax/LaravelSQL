<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel & Doc Viewer</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.5/font/bootstrap-icons.css" rel="stylesheet">
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
        .sheet-list, .doc-list {
            list-style: none;
            margin-left: 0;
            padding-left: 0;
        }
        .doc-content {
            border: 1px solid #dee2e6;
            padding: 1rem;
            margin-top: 1rem;
            background-color: #f8f9fa;
            white-space: pre-wrap;
        }
        .doc-content p {
            margin: 0 0 0.5rem 0;
        }
        .doc-content table {
            width: 100%;
            margin-bottom: 1rem;
        }
        .doc-content table th, .doc-content table td {
            padding: 0.5rem;
            vertical-align: top;
        }
        .section-title {
            margin-bottom: 1rem;
        }
        .form-section, .list-section {
            margin-bottom: 2rem;
        }
        .sheet-item, .doc-item {
            display: flex;
            align-items: center;
            gap: 0.75rem;
            padding: 0.25rem 0;
        }
        .excel-item {
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 1.5rem;
        }
        .eye-btn, .map-btn, .toggle-btn {
            font-size: 1.3rem;
            padding: 0.4rem;
            border: none;
            background-color: #e9ecef;
            border-radius: 50%;
            cursor: pointer;
            transition: transform 0.2s, box-shadow 0.2s, background-color 0.2s;
            min-width: 2rem;
            text-align: center;
        }
        .eye-btn {
            color: #0d6efd;
        }
        .map-btn {
            color: #198754;
        }
        .toggle-btn {
            color: #6c757d;
        }
        .eye-btn:hover {
            transform: scale(1.1);
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.2);
            background-color: #d0e7ff;
            color: #0056b3;
        }
        .map-btn:hover {
            transform: scale(1.1);
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.2);
            background-color: #d4edda;
            color: #146c43;
        }
        .toggle-btn:hover {
            transform: scale(1.1);
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.2);
            background-color: #e2e6ea;
            color: #495057;
        }
        .sheet-actions, .doc-actions {
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }
        .doc-item .delete-form {
            margin-left: auto;
        }
        .sheet-name, .doc-name {
            flex: 1;
        }
        .fields-table {
            margin-top: 1rem;
            font-size: 0.9rem;
            width: 100%;
        }
        .fields-table th {
            background-color: #e9ecef;
            white-space: nowrap;
            padding: 0.5rem;
        }
        .fields-scroll {
            overflow-x: auto;
            overflow-y: auto;
            max-height: 200px;
            margin-bottom: 1rem;
        }
        .variables-list {
            margin-top: 1rem;
            padding-left: 1rem;
        }
        .variables-scroll {
            overflow-y: auto;
            max-height: 200px;
            margin-bottom: 1rem;
        }
        .field-section, .variable-section {
            margin-top: 1.5rem;
            padding: 1rem;
            border: 1px solid #dee2e6;
            border-radius: 0.25rem;
            background-color: #fff;
        }
        .fields-scroll::-webkit-scrollbar,
        .variables-scroll::-webkit-scrollbar {
            width: 8px;
            height: 8px;
        }
        .fields-scroll::-webkit-scrollbar-thumb,
        .variables-scroll::-webkit-scrollbar-thumb {
            background-color: #adb5bd;
            border-radius: 4px;
        }
        .fields-scroll::-webkit-scrollbar-track,
        .variables-scroll::-webkit-scrollbar-track {
            background-color: #f1f1f1;
        }
        .sheet-list.hidden,
        .doc-list.hidden {
            display: none;
        }
        .excel-header {
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }
    </style>
</head>
<body>
    <div class="container mt-5">
        <div class="row">
            <!-- Cột trái: Đọc File Excel và Danh sách Excel -->
            <div class="col-12 col-md-6">
                <div class="form-section">
                    <h2 class="section-title">Đọc File Excel</h2>
                    <form action="{{ route('excel.addExcel') }}" method="POST">
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
                </div>

                @if (!empty($excelFiles))
                    <div class="list-section">
                        <h3 class="section-title">Danh Sách Excel</h3>
                        <ul class="list-group">
                            @foreach ($excelFiles as $fileIndex => $file)
                                <li class="list-group-item">
                                    <div class="excel-item">
                                        <div>
                                            <div class="excel-header">
                                                <button type="button" class="toggle-btn" data-target="sheet-list-{{ $fileIndex }}">
                                                    <i class="bi bi-chevron-down"></i>
                                                </button>
                                                <strong>{{ $file['name'] }}</strong>
                                            </div>
                                            <ul class="sheet-list hidden" id="sheet-list-{{ $fileIndex }}">
                                                @foreach ($file['sheets'] as $sheetIndex => $sheetName)
                                                    <li class="sheet-item">
                                                        <div class="sheet-actions">
                                                            <form action="{{ route('excel.readSheet', [$fileIndex, $sheetIndex]) }}" method="GET" class="d-inline">
                                                                <button type="submit" class="eye-btn"><i class="bi bi-eye"></i></button>
                                                            </form>
                                                            <form action="{{ route('excel.fields', [$fileIndex, $sheetIndex]) }}" method="GET" class="d-inline">
                                                                <button type="submit" class="map-btn"><i class="bi bi-diagram-3"></i></button>
                                                            </form>
                                                        </div>
                                                        <span class="sheet-name {{ isset($currentFileIndex) && isset($currentSheetIndex) && $currentFileIndex == $fileIndex && $currentSheetIndex == $sheetIndex ? 'fw-bold' : '' }}">
                                                            {{ $sheetName }}
                                                        </span>
                                                    </li>
                                                @endforeach
                                            </ul>
                                        </div>
                                        <form action="{{ route('excel.removeExcel', $fileIndex) }}" method="POST">
                                            @csrf
                                            @method('DELETE')
                                            <button type="submit" class="btn btn-danger btn-sm">Xóa Excel</button>
                                        </form>
                                    </div>
                                </li>
                            @endforeach
                        </ul>

                        <!-- Hiển thị danh sách trường -->
                        @if (session('sheet_fields'))
                            @foreach (session('sheet_fields') as $fIndex => $sheets)
                                @foreach ($sheets as $sIndex => $sheetData)
                                    <div class="field-section">
                                        <div class="d-flex justify-content-between align-items-center mb-2">
                                            <h4 class="section-title">Trường {{ $sheetData['sheet_name'] }}</h4>
                                            <form action="{{ route('excel.removeFields', [$fIndex, $sIndex]) }}" method="GET">
                                                <button type="submit" class="btn btn-danger btn-sm">Xóa danh sách này</button>
                                            </form>
                                        </div>
                                        <div class="fields-scroll">
                                            <table class="table table-bordered fields-table">
                                                <thead>
                                                    <tr>
                                                        @foreach ($sheetData['fields'] as $field)
                                                            <th>{{ $field }}</th>
                                                        @endforeach
                                                    </tr>
                                                </thead>
                                            </table>
                                        </div>
                                    </div>
                                @endforeach
                            @endforeach
                        @endif
                    </div>
                @endif
            </div>

            <!-- Cột phải: Đọc File Doc và Danh sách Doc -->
            <div class="col-12 col-md-6">
                <div class="form-section">
                    <h2 class="section-title">Đọc File Doc</h2>
                    <form action="{{ route('doc.addDoc') }}" method="POST">
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
                </div>

                @if (!empty($docFiles))
                    <div class="list-section">
                        <div class="d-flex justify-content-between align-items-center mb-2">
                            <h3 class="section-title">Danh Sách Doc</h3>
                            <button type="button" class="toggle-btn" data-target="doc-list">
                                <i class="bi bi-chevron-down"></i>
                            </button>
                        </div>
                        <ul class="list-group doc-list hidden" id="doc-list">
                            @foreach ($docFiles as $docIndex => $doc)
                                <li class="list-group-item">
                                    <div class="doc-item">
                                        <div class="doc-actions">
                                            <form action="{{ route('doc.readDoc', $docIndex) }}" method="GET" class="d-inline">
                                                <button type="submit" class="eye-btn"><i class="bi bi-eye"></i></button>
                                            </form>
                                            <form action="{{ route('doc.variables', $docIndex) }}" method="GET" class="d-inline">
                                                <button type="submit" class="map-btn"><i class="bi bi-diagram-3"></i></button>
                                            </form>
                                            <span class="doc-name {{ isset($currentDocIndex) && $currentDocIndex == $docIndex ? 'fw-bold' : '' }}">
                                                {{ $doc['name'] }}
                                            </span>
                                        </div>
                                        <form action="{{ route('doc.removeDoc') }}" method="POST" class="delete-form">
                                            @csrf
                                            <input type="hidden" name="doc_index" value="{{ $docIndex }}">
                                            <button type="submit" class="btn btn-danger btn-sm">X</button>
                                        </form>
                                    </div>
                                </li>
                            @endforeach
                        </ul>

                        <!-- Hiển thị danh sách biến -->
                        @if (session('doc_variables'))
                            @foreach (session('doc_variables') as $dIndex => $docData)
                                <div class="variable-section">
                                    <div class="d-flex">
                                        <h4 class="section-title">Biến {{ $docData['doc_name'] }}</h4>
                                        <form action="{{ route('doc.removeVariables', $dIndex) }}" method="GET" class="ms-auto">
                                            <button type="submit" class="btn btn-danger btn-sm">Xóa danh sách này</button>
                                        </form>
                                    </div>
                                    <div class="variables-scroll">
                                        <ul class="variables-list">
                                            @foreach ($docData['variables'] as $variable)
                                                <li>{{ $variable }}</li>
                                            @endforeach
                                        </ul>
                                    </div>
                                </div>
                            @endforeach
                        @endif
                    </div>
                @endif
            </div>
        </div>

        <!-- Thông báo lỗi/thành công -->
        @if (session('error'))
            <div class="row mt-3">
                <div class="col-12">
                    <div class="alert alert-danger">
                        {{ session('error') }}
                    </div>
                </div>
            </div>
        @endif

        @if (session('success'))
            <div class="row mt-3">
                <div class="col-12">
                    <div class="alert alert-success">
                        {{ session('success') }}
                    </div>
                </div>
            </div>
        @endif

        <!-- Dữ liệu Excel -->
        @if (isset($data) && !empty($data))
            <div class="row mt-5">
                <div class="col-12">
                    <h3 class="section-title">Dữ Liệu: {{ $excelFiles[$currentFileIndex]['name'] }} / {{ $excelFiles[$currentFileIndex]['sheets'][$currentSheetIndex] }}</h3>
                    <table class="table table-bordered mt-3 table-smaller">
                        <thead>
                            <tr>
                                @foreach ($data[0] as $index => $cell)
                                    @if ($cell['colspan'] !== 0)
                                        <th class="{{ isset($statusColumnIndex) && $index == $statusColumnIndex ? 'status-column' : '' }}"
                                            @if ($cell['colspan'] > 1) colspan="{{ $cell['colspan'] }}" @endif>
                                            {{ $cell['value'] !== '' ? $cell['value'] : '' }}
                                        </th>
                                    @endif
                                @endforeach
                            </tr>
                        </thead>
                        <tbody>
                            @foreach (array_slice($data, 1) as $row)
                                <tr>
                                    @foreach ($row as $colIndex => $cell)
                                        @if ($cell['colspan'] !== 0)
                                            <td class="{{ isset($statusColumnIndex) && $colIndex == $statusColumnIndex ? 'status-column' : '' }}"
                                                @if ($cell['colspan'] > 1) colspan="{{ $cell['colspan'] }}" @endif>
                                                {{ $cell['value'] !== '' ? $cell['value'] : '' }}
                                            </td>
                                        @endif
                                    @endforeach
                                </tr>
                            @endforeach
                        </tbody>
                    </table>
                </div>
            </div>
        @endif

        <!-- Nội dung Doc -->
        @if (isset($docContent))
            <div class="row mt-5">
                <div class="col-12">
                    <h3 class="section-title">Nội Dung: {{ $docFiles[$currentDocIndex]['name'] }}</h3>
                    <div class="doc-content">
                        {!! $docContent !!}
                    </div>
                </div>
            </div>
        @endif
    </div>

    <script>
        // Xử lý toggle ẩn/hiện danh sách sheet và danh sách Doc
        document.querySelectorAll('.toggle-btn').forEach(button => {
            const targetId = button.getAttribute('data-target');
            const target = document.getElementById(targetId);
            const icon = button.querySelector('i');

            // Khôi phục trạng thái từ localStorage
            const isExpanded = localStorage.getItem(`toggle-${targetId}`) === 'true';
            if (isExpanded) {
                target.classList.remove('hidden');
                icon.classList.remove('bi-chevron-down');
                icon.classList.add('bi-chevron-up');
            } else {
                target.classList.add('hidden');
                icon.classList.remove('bi-chevron-up');
                icon.classList.add('bi-chevron-down');
            }

            // Xử lý toggle
            button.addEventListener('click', () => {
                const isHidden = target.classList.contains('hidden');
                if (isHidden) {
                    target.classList.remove('hidden');
                    icon.classList.remove('bi-chevron-down');
                    icon.classList.add('bi-chevron-up');
                    localStorage.setItem(`toggle-${targetId}`, 'true');
                } else {
                    target.classList.add('hidden');
                    icon.classList.remove('bi-chevron-up');
                    icon.classList.add('bi-chevron-down');
                    localStorage.setItem(`toggle-${targetId}`, 'false');
                }
            });
        });
    </script>
</body>
</html>