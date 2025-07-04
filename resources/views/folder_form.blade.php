<!DOCTYPE html>
<html>
<head>
    <title>Tạo thư mục</title>
</head>
<body>
    <h1>Tạo Thư Mục Mới</h1>

    @if(session('success'))
        <p style="color: green">{{ session('success') }}</p>
    @endif

    @if(session('error'))
        <p style="color: red">{{ session('error') }}</p>
    @endif

    <form method="POST" action="{{ route('folder.create') }}">
        @csrf

        <label>Đường dẫn (path):</label><br>
        <input type="text" name="path" value="{{ old('path') }}" style="width: 400px;"><br><br>

        <label>Tên thư mục mới:</label><br>
        <input type="text" name="folder_name" value="{{ old('folder_name') }}"><br><br>

        <button type="submit">Tạo folder</button>
    </form>
</body>
</html>
