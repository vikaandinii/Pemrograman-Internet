<?php
require 'vendor/autoload.php'; // Include PhpSpreadsheet via Composer

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Fungsi untuk mengkonversi nilai angka ke nilai huruf
function konversiNilaiHuruf($nilai) {
    if ($nilai >= 80) {
        return "A";
    } elseif ($nilai >= 71) {
        return "B+";
    } elseif ($nilai >= 65) {
        return "B";
    } elseif ($nilai >= 60) {
        return "C+";
    } elseif ($nilai >= 55) {
        return "C";
    } elseif ($nilai >= 50) {
        return "D+";
    } elseif ($nilai >= 40) {
        return "D";
    } else {
        return "E";
    }
}

if (isset($_POST['submit'])) {
    $nama = $_POST['nama'];
    $nim = $_POST['nim'];
    $nilai = $_POST['nilai'];

    // Validasi nama agar hanya mengandung huruf dan spasi
    if (!preg_match("/^[a-zA-Z\s]+$/", $nama)) {
        $error_message = "Nama hanya boleh berisi huruf dan spasi.";
    }

    // Validasi NIM agar hanya mengandung angka
    if (!preg_match("/^[0-9]+$/", $nim)) {
        $error_message = "NIM hanya boleh berisi angka.";
    }

    if (!isset($error_message)) {
        $nilai_huruf = konversiNilaiHuruf($nilai);

        // Cek apakah file Excel sudah ada
        $filePath = 'nilai_mahasiswa.xlsx';

        if (file_exists($filePath)) {
            $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($filePath);
        } else {
            $spreadsheet = new Spreadsheet();
            $sheet = $spreadsheet->getActiveSheet();
            $sheet->setCellValue('A1', 'Nama');
            $sheet->setCellValue('B1', 'NIM');
            $sheet->setCellValue('C1', 'Nilai Angka');
            $sheet->setCellValue('D1', 'Nilai Huruf');
        }

        $sheet = $spreadsheet->getActiveSheet();
        $highestRow = $sheet->getHighestRow();
        $sheet->setCellValue('A' . ($highestRow + 1), $nama);
        $sheet->setCellValue('B' . ($highestRow + 1), $nim);
        $sheet->setCellValue('C' . ($highestRow + 1), $nilai);
        $sheet->setCellValue('D' . ($highestRow + 1), $nilai_huruf);

        $writer = new Xlsx($spreadsheet);
        $writer->save($filePath);

        $success_message = "Data berhasil disimpan.";
    }
}
?>

<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Hasil Input Nilai</title>
    <link rel="stylesheet" href="style.css">
    <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
    <script>
        function validateForm() {
            var nama = document.forms["nilaiForm"]["nama"].value;
            var nim = document.forms["nilaiForm"]["nim"].value;

            var namaRegex = /^[a-zA-Z\s]+$/;
            var nimRegex = /^[0-9]+$/;

            if (!namaRegex.test(nama)) {
                alert("Nama hanya boleh berisi huruf dan spasi.");
                return false;
            }

            if (!nimRegex.test(nim)) {
                alert("NIM hanya boleh berisi angka.");
                return false;
            }

            return true;
        }
    </script>
</head>
<body>

<div class="result-container">
    <?php if (isset($error_message)): ?>
        <p><?php echo $error_message; ?></p>
    <?php elseif (isset($success_message)): ?>
        <div class='result'>
            <h3>Hasil Input Nilai</h3>
            <div class='result-content'>
                <p>Nama: <strong><?php echo htmlspecialchars($nama); ?></strong></p>
                <p>NIM: <strong><?php echo htmlspecialchars($nim); ?></strong></p>
                <p>Nilai Angka: <strong><?php echo htmlspecialchars($nilai); ?></strong></p>
                <p>Nilai Huruf: <strong><?php echo htmlspecialchars($nilai_huruf); ?></strong></p>
            </div>
            <a href='index.html' class='back-button'>Kembali</a>
        </div>
    <?php else: ?>
        <form name="nilaiForm" action="" method="POST" onsubmit="return validateForm()">
            <label for="nama">Nama:</label>
            <input type="text" id="nama" name="nama" required>
            
            <label for="nim">NIM:</label>
            <input type="text" id="nim" name="nim" required>
            
            <label for="nilai">Nilai:</label>
            <input type="number" id="nilai" name="nilai" required>

            <input type="submit" name="submit" value="Submit">
        </form>
    <?php endif; ?>
</div>

</body>
</html>
