const express = require('express');
const cors = require('cors');
const exceljs = require('exceljs');
const fs = require('fs');
const pdfkit = require('pdfkit');
const moment = require('moment');
require('moment/locale/id');
const { fakultasData, prodiData } = require('./data');

const app = express();

app.use(cors());

app.use(require('body-parser').json());
app.use(require('body-parser').urlencoded({ extended: true }));
app.use(express.json({ limit: '1mb' }));

const getCurrentTimestamp = () => {
    const now = new Date();

    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    const hours = String(now.getHours()).padStart(2, '0');
    const minutes = String(now.getMinutes()).padStart(2, '0');
    const seconds = String(now.getSeconds()).padStart(2, '0');

    const timestamp = `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
    return timestamp;
};

// Mengambil tanggal hari ini dan memformatnya menjadi "DD MMMM YYYY"
const today = moment().format('DD MMMM YYYY');

// Fungsi untuk memformat angka menjadi format rupiah
const formatRupiah = (number) => {
    return new Intl.NumberFormat('id-ID', {
        style: 'currency',
        currency: 'IDR',
        minimumFractionDigits: 0
    }).format(number);
};

// Fungsi untuk membuat huruf awal kapital
const capitalLetter = (str) => {
    return str
        .split(" ")
        .map(word => word.charAt(0).toUpperCase() + word.slice(1))
        .join(" ");
};

app.post('/search', (req, res) => {
    const programStudi = req.body.programStudi;

    const workbook = new exceljs.Workbook();
    workbook.xlsx.readFile('clearinghouse.xlsx')
        .then(() => {
            const worksheet = workbook.getWorksheet(1);
            let students = [];
            worksheet.eachRow((row, rowNumber) => {
                if (rowNumber !== 1) {
                    const programStudiFromSheet = row.getCell(5).value;

                    if (programStudiFromSheet === programStudi) {
                        students.push({
                            "No": students.length + 1,
                            "ID Pendaftar": row.getCell(2).value,
                            "Nama": row.getCell(3).value,
                            "Nilai": row.getCell(9).value,
                            "Pilihan 1": programStudiFromSheet,
                            "DP": row.getCell(17).value,
                            "Pilihan 2": row.getCell(18).value,
                        });
                    }
                }
            });

            if (students.length > 0) {
                res.json({ "found": true, "peserta": students, "jumlah": students.length });
            } else {
                res.json({ "message": `Tidak ada data untuk Program Studi: ${programStudi}` });
            }
        })
        .catch(err => {
            console.error(err);
            res.status(500).send('Error reading Excel file');
        });
});

function createTable(doc, data, startX, startY, columnWidths, rowHeight) {
    const numColumns = columnWidths.length;
    let currentY = startY;

    function drawHeader() {
        doc.font('Arial Bold Font').fontSize(8); // Mengatur font untuk header
        doc.rect(startX, currentY, columnWidths.reduce((a, b) => a + b, 0), rowHeight);
        data.header.forEach((header, i) => {
            const colStartX = startX + columnWidths.slice(0, i).reduce((a, b) => a + b, 0);
            doc.text(header, colStartX + 5, currentY + 5, {
                width: columnWidths[i] - 10,
                align: 'center' // Mengatur teks menjadi rata tengah
            });
            doc.lineWidth(0.5).moveTo(colStartX, currentY).lineTo(colStartX, currentY + rowHeight);
        });
        // Garis vertikal terakhir untuk header
        doc.lineWidth(0.5).moveTo(startX + columnWidths.reduce((a, b) => a + b, 0), currentY).lineTo(startX + columnWidths.reduce((a, b) => a + b, 0), currentY + rowHeight);
        currentY += rowHeight; // Pindah ke baris berikutnya
    }

    function checkPageBreak() {
        if (currentY + rowHeight > doc.page.height - doc.page.margins.bottom) {
            doc.addPage();
            currentY = doc.page.margins.top;
            drawHeader();
        }
    }

    // Menggambar header tabel untuk halaman pertama
    drawHeader();

    // Mengatur font untuk baris data
    doc.font('Arial Font').fontSize(8);

    // Menggambar baris tabel
    data.rows.forEach((row, rowIndex) => {
        checkPageBreak(); // Periksa apakah perlu halaman baru sebelum menggambar baris
        doc.font('Arial Font').fontSize(8);
        doc.rect(startX, currentY, columnWidths.reduce((a, b) => a + b, 0), rowHeight).stroke();
        row.forEach((cell, cellIndex) => {
            const colStartX = startX + columnWidths.slice(0, cellIndex).reduce((a, b) => a + b, 0);
            const align = cellIndex === 5 ? 'right' : 'left'; // Mengatur alignment untuk kolom DP menjadi right
            doc.text(cell, colStartX + 5, currentY + 5, {
                width: columnWidths[cellIndex] - 10,
                align: align
            });
            doc.lineWidth(0.5).moveTo(colStartX, currentY).lineTo(colStartX, currentY + rowHeight).stroke();
        });
        // Garis vertikal terakhir untuk setiap baris
        doc.lineWidth(0.5).moveTo(startX + columnWidths.reduce((a, b) => a + b, 0), currentY).lineTo(startX + columnWidths.reduce((a, b) => a + b, 0), currentY + rowHeight).stroke();
        currentY += rowHeight; // Pindah ke baris berikutnya
    });
}


app.post('/download', (req, res) => {
    const programStudi = req.body.programStudi;
    const idFakultas = req.body.fakultas;
    let namaFakultas = '';
    console.log(typeof (idFakultas));

    switch (idFakultas) {
        case "1":
            namaFakultas = 'Fakultas Hukum';
            break;
        case "2":
            namaFakultas = 'Fakultas Ilmu Sosial dan Ilmu Politik';
            break;
        case "3":
            namaFakultas = 'Fakultas Teknik';
            break;
        case "4":
            namaFakultas = 'Fakultas Ekonomi dan Bisnis';
            break;
        case "5":
            namaFakultas = 'Fakultas Keguruan dan Ilmu Pendidikan';
            break;
        case "6":
            namaFakultas = 'Fakultas Ilmu Seni dan Sastra';
            break;
    }

    const workbook = new exceljs.Workbook();
    workbook.xlsx.readFile('clearinghouse.xlsx')
        .then(() => {
            const worksheet = workbook.getWorksheet(1);
            let students = [];
            worksheet.eachRow((row, rowNumber) => {
                if (rowNumber !== 1) {
                    const programStudiFromSheet = row.getCell(5).value;

                    if (programStudiFromSheet === programStudi) {
                        students.push({
                            no: students.length + 1,
                            idPendaftar: row.getCell(2).value,
                            nama: row.getCell(3).value,
                            nilai: row.getCell(9).value,
                            pilihan1: programStudiFromSheet,
                            dp: formatRupiah(row.getCell(17).value),
                            pilihan2: row.getCell(18).value,
                        });
                    }
                }
            });

            if (students.length > 0) {
                const pdfPath = `BUKTI_WISUDA_${programStudi}.pdf`;

                const doc = new pdfkit({ size: 'A4', layout: 'landscape', margin: { right: 10 } });
                const buffers = [];
                doc.on('data', buffers.push.bind(buffers));
                doc.on('end', () => {
                    const pdfData = Buffer.concat(buffers);

                    // Simpan file PDF ke sistem file
                    const pdfPathOnServer = `pdf_output/${pdfPath}`;
                    fs.writeFile(pdfPathOnServer, pdfData, (err) => {
                        if (err) {
                            console.error('Error saving PDF file:', err);
                            res.status(500).send('Error saving PDF file');
                        } else {
                            const infoLog = `${getCurrentTimestamp()} - Download Success for Program Studi: ${programStudi}\n`;
                            fs.appendFileSync(`logDownloadSuccess.log`, infoLog);
                            // Kirim file PDF sebagai respons ke client
                            res.set({
                                'Content-Type': 'application/pdf',
                                'Content-Disposition': `attachment; filename=${pdfPath}`,
                                'Content-Length': pdfData.length
                            });
                            res.send(pdfData);
                        }
                    });
                });

                // Mendapatkan ukuran halaman PDF
                const pageWidth = doc.page.width;
                // Mendapatkan ukuran gambar
                const imageWidth = 539;

                // Menghitung koordinat untuk menempatkan gambar di tengah halaman
                const x = (pageWidth - imageWidth) / 2;

                // Menambahkan header dan footer
                // doc.image('header.PNG', x, 14, { width: imageWidth });
                // doc.image("footer.PNG", x, 771, { width: 539 });

                // Tambahkan konten PDF
                const text1 = "REKAPITULASI NILAI UJIAN SARINGAN MASUK CALON MAHASISWA";
                const text2 = "PMB 2024/2025 GELOMBANG 2 UNIVERSITAS PASUNDAN";

                const textFakultas = `${namaFakultas.toUpperCase()}`;
                const textProdi = `PROGRAM STUDI ${programStudi.toUpperCase()}`;

                const text3 = "Mengetahui,";
                const text4 = "Ketua PPMB,";
                const text5 = "Prof. Dr. Cartono, S.Pd., M.Pd., M.T.";

                const text6 = "Mengetahui,";
                const text7 = "Sekertaris PPMB,";
                const text8 = "Drs. H. Wawan Satriawan";

                const text9 = "Mengetahui,";
                const text10 = "Rektor,";
                const text11 = "Prof. Dr. H. Azhar Affandi, S.E., M.Sc.";

                doc.registerFont('Arial Font', 'fonts/arial.ttf');
                doc.registerFont('Arial Bold Font', 'fonts/arial-bold.ttf');
                doc.font('Arial Bold Font')
                    .fontSize(12).text(text1, 96, 50, { align: 'center' });

                doc.fontSize(12).text(text2, 96, 67, { align: 'center' });

                doc.fontSize(12).text(textFakultas, 96, 100, { align: 'center' });
                doc.fontSize(12).text(textProdi, 96, 117, { align: 'center' });

                // buat tabel disini
                const tableData = {
                    header: ["No", "ID Pendaftar", "Nama", "Nilai", "Pilihan 1", "DP", "Pilihan 2"],
                    rows: students.map(student => [
                        student.no,
                        student.idPendaftar,
                        student.nama,
                        student.nilai,
                        student.pilihan1,
                        student.dp,
                        student.pilihan2,
                    ])
                };

                //setingan font 9
                // const columnWidths = [26, 65, 212, 33, 209, 65, 209];
                // const rowHeight = 18;
                // const startX = 15;
                // const startY = 150;

                // setingan font 8
                const columnWidths = [26, 65, 212, 33, 200, 65, 200];
                const rowHeight = 18;
                const startX = 25;
                const startY = 150;

                createTable(doc, tableData, startX, startY, columnWidths, rowHeight);

                doc.font('Arial Bold Font')
                    .fontSize(11).text(text3, 40, 520, { align: 'left' });

                doc.fontSize(11).text(text4, 40, 85, { align: 'left' });
                doc.fontSize(11).text(text5, 40, 150, { align: 'left' });

                doc.fontSize(11).text(text6, 640, 68, { align: 'left' });
                doc.fontSize(11).text(text7, 640, 85, { align: 'left' });
                doc.fontSize(11).text(text8, 640, 150, { align: 'left', lineBreak: false });

                doc.fontSize(11).text(text9, 40, 170, { align: 'center' });
                doc.fontSize(11).text(text10, 40, 187, { align: 'center' });
                doc.fontSize(11).text(text11, 40, 252, { align: 'center' });

                doc.end();
            } else {
                res.status(404).send('NIM tidak terdaftar');
            }
        })
        .catch(err => {
            console.error(err);
            res.status(500).send('Error reading Excel file');
        });
});

// ==================================================DOWNLOAD DISTRIBUSI NILAI TES======================================

function createTableEstimasiDP(doc, data, startX, startY, columnWidths, rowHeight) {
    const numColumns = columnWidths.length;
    let currentY = startY;

    function drawHeader() {
        doc.font('Arial Bold Font').fontSize(8); // Mengatur font untuk header

        // Menggambar header
        let colStartX = startX;

        // Baris pertama header
        data.header[0].forEach((header, colIndex) => {
            const width = header.colSpan ? columnWidths.slice(colIndex, colIndex + header.colSpan).reduce((a, b) => a + b, 0) : columnWidths[colIndex];
            const align = 'center';
            const text = header.text;

            doc.rect(colStartX, currentY, width, rowHeight).stroke();
            doc.text(text, colStartX + 5, currentY + 5, {
                width: width - 10,
                align: align
            });

            colStartX += width;
        });

        currentY += rowHeight;

        // Baris kedua header
        colStartX = startX;
        data.header[1].forEach((header, colIndex) => {
            const width = columnWidths[colIndex];
            const align = 'center';
            const text = header;

            doc.rect(colStartX, currentY, width, rowHeight).stroke();
            doc.text(text, colStartX + 5, currentY + 5, {
                width: width - 10,
                align: align
            });

            colStartX += width;
        });

        currentY += rowHeight;

        // Garis vertikal terakhir untuk header
        doc.lineWidth(0.5).moveTo(startX, currentY).lineTo(startX + columnWidths.reduce((a, b) => a + b, 0), currentY).stroke();
    }

    function checkPageBreak() {
        if (currentY + rowHeight > doc.page.height - doc.page.margins.bottom) {
            doc.addPage();
            currentY = doc.page.margins.top;
            drawHeader();
        }
    }

    // Menggambar header tabel untuk halaman pertama
    drawHeader();

    // Mengatur font untuk baris data
    doc.font('Arial Font').fontSize(8);

    // Menggambar baris tabel
    data.rows.forEach((row, rowIndex) => {
        checkPageBreak(); // Periksa apakah perlu halaman baru sebelum menggambar baris
        doc.rect(startX, currentY, columnWidths.reduce((a, b) => a + b, 0), rowHeight).stroke();
        row.forEach((cell, cellIndex) => {
            const colStartX = startX + columnWidths.slice(0, cellIndex).reduce((a, b) => a + b, 0);
            // const align = cellIndex === 1 || cellIndex === 2 || cellIndex === 3  ? 'right' : 'left';
            const align = cellIndex !== 0 && cellIndex !== 8 ? 'right' : cellIndex === 0 ? 'left' : 'center';
            doc.text(cell.toString(), colStartX + 5, currentY + 5, {
                width: columnWidths[cellIndex] - 10,
                align: align
            });
            doc.lineWidth(0.5).moveTo(colStartX, currentY).lineTo(colStartX, currentY + rowHeight).stroke();
        });
        // Garis vertikal terakhir untuk setiap baris
        doc.lineWidth(0.5).moveTo(startX + columnWidths.reduce((a, b) => a + b, 0), currentY).lineTo(startX + columnWidths.reduce((a, b) => a + b, 0), currentY + rowHeight).stroke();
        currentY += rowHeight; // Pindah ke baris berikutnya
    });
}


app.post('/downloadEstimasiDP', (req, res) => {
    const idFakultas = req.body.fakultas;
    let namaFakultas = '';
    let namaDekan = '';

    switch (idFakultas) {
        case "1":
            namaFakultas = 'Fakultas Hukum';
            namaDekan = "Prof. Dr. Anthon Fredi Susanto, S.H., M.Hum.";
            break;
        case "2":
            namaFakultas = 'Fakultas Ilmu Sosial dan Ilmu Politik';
            namaDekan = "Dr. Kunkunrat, M.Si.";
            break;
        case "3":
            namaFakultas = 'Fakultas Teknik';
            namaDekan = "Prof. Dr. Ir. Yusman Taufik, MP.";
            break;
        case "4":
            namaFakultas = 'Fakultas Ekonomi dan Bisnis';
            namaDekan = "Dr. Juanim, S.E., M.Si.";
            break;
        case "5":
            namaFakultas = 'Keguruan dan Ilmu Pendidikan';
            namaDekan = "Dr. Hj. Dini Riani, S.E., M.M.";
            break;
        case "6":
            namaFakultas = 'Fakultas Ilmu Seni dan Sastra';
            namaDekan = "Dr. Hj. Senny Suzanna Alwasilah, S.S., M.Pd.";
            break;
    }

    const workbook = new exceljs.Workbook();
    workbook.xlsx.readFile('clearinghouse.xlsx')
        .then(() => {
            const worksheet = workbook.getWorksheet(1);
            let jumlahMahasiswaPerProdi = {};
            let totalMahasiswa = 0;
            let danaPembangunan = 0;

            worksheet.eachRow((row, rowNumber) => {
                if (rowNumber !== 1) {
                    const programStudiFromSheet = row.getCell(5).value;
                    const fakultasFromSheet = row.getCell(7).value.toUpperCase();

                    if (fakultasFromSheet === namaFakultas.toUpperCase()) {
                        if (!jumlahMahasiswaPerProdi[programStudiFromSheet]) {
                            jumlahMahasiswaPerProdi[programStudiFromSheet] = {
                                count: 0,
                                dp: row.getCell(17).value
                            };
                        }
                        danaPembangunan = row.getCell(17).value;
                        jumlahMahasiswaPerProdi[programStudiFromSheet].count++;
                        totalMahasiswa++;
                    }
                }
            });

            if (totalMahasiswa > 0) {
                if (namaFakultas === 'Keguruan dan Ilmu Pendidikan') {
                    namaFakultas = 'Fakultas Keguruan dan Ilmu Pendidikan';
                }

                const pdfPath = `ESTIMASI_DP_${namaFakultas}.pdf`;

                const doc = new pdfkit({ size: 'A4', layout: 'landscape', margin: { right: 10 } });
                const buffers = [];
                doc.on('data', buffers.push.bind(buffers));
                doc.on('end', () => {
                    const pdfData = Buffer.concat(buffers);

                    const pdfPathOnServer = `pdf_output/${pdfPath}`;
                    fs.writeFile(pdfPathOnServer, pdfData, (err) => {
                        if (err) {
                            console.error('Error saving PDF file:', err);
                            res.status(500).send('Error saving PDF file');
                        } else {
                            const infoLog = `${getCurrentTimestamp()} - Download Success for Fakultas: ${namaFakultas}\n`;
                            fs.appendFileSync(`logDownloadSuccess.log`, infoLog);
                            res.set({
                                'Content-Type': 'application/pdf',
                                'Content-Disposition': `attachment; filename=${pdfPath}`,
                                'Content-Length': pdfData.length
                            });
                            res.send(pdfData);
                        }
                    });
                });

            
                // Menambahkan image
                doc.image('logo_unpas.PNG', 35, 50, { width: 50 });

                const text1 = 'Estimasi Pendapatan dari Dana Pembangunan (DP)';
                const text2 = 'Asumsi Diterima di Pilihan Program Studi Kesatu';
                const text3 = 'USM Gelombang 2';
                const text4 = 'Tahun 2024/2025';

                const text5 = "Mengetahui,";
                const text6 = "Ketua PPMB,";
                const text7 = "Prof. Dr. Cartono, S.Pd., M.Pd., M.T.";

                const text8 = "Mengetahui,";
                const text9 = `Dekan ${capitalLetter(namaFakultas)},`;
                const text10 = namaDekan;

                const text11 = "Menyetujui,";
                const text12 = "Rektor,";
                const text13 = "Prof. Dr. H. Azhar Affandi, S.E., M.Sc.";

                const textFakultas = `${capitalLetter(namaFakultas)}`;

                doc.registerFont('Arial Font', 'fonts/arial.ttf');
                doc.registerFont('Arial Bold Font', 'fonts/arial-bold.ttf');
                doc.font('Arial Bold Font')
                    .fontSize(10).text(text1, 96, 50, { align: 'center' });
                doc.fontSize(10).text(text2, 96, 67, { align: 'center' });
                doc.fontSize(10).text(text3, 96, 84, { align: 'center' });
                doc.fontSize(10).text(text4, 96, 101, { align: 'center' });

                doc.fontSize(10).text(textFakultas, 30, 133, { align: 'left' });

                const tableData = {
                    header: [
                        [{ text: "Program Studi" }, { text: "Total tagihan Per Jenis", colSpan: 1 }, { text: "Total", colSpan: 5 }, { text: "Lunas", colSpan: 1 }, { text: "Jumlah", colSpan: 1 }],
                        ["", "DP", "INFAQ", "PKKMB", "Tagihan", "Denda", "Potongan", "", "Tagihan"]
                    ],
                    rows: []
                };

                const totalTagihanFakultas = totalMahasiswa * danaPembangunan;
                
                // Sisipkan baris untuk nama fakultas dan total mahasiswa
                const namaFakultasRow = [
                    `${capitalLetter(namaFakultas)}`,
                    formatRupiah(totalTagihanFakultas),
                    "0", 
                    "0",
                    formatRupiah(totalTagihanFakultas),
                    "0",
                    "0",
                    "0",
                    totalMahasiswa.toString(),
                ];
                
                tableData.rows.push(namaFakultasRow);
                
                // Tambahkan data program studi ke dalam rows
                Object.keys(jumlahMahasiswaPerProdi).forEach(programStudi => {
                    const dp = jumlahMahasiswaPerProdi[programStudi].dp || 0;
                    const jumlahMahasiswa = jumlahMahasiswaPerProdi[programStudi].count || 0;
                    const totalTagihanPerProdi = dp * jumlahMahasiswa;
                
                    const rowData = [
                        programStudi,
                        formatRupiah(totalTagihanPerProdi),
                        "0", 
                        "0", 
                        formatRupiah(totalTagihanPerProdi),
                        "0", 
                        "0", 
                        "0", 
                        jumlahMahasiswa.toString(),
                    ];
                
                    tableData.rows.push(rowData);
                });

                const columnWidths = [190, 100, 70, 70, 75, 70, 70, 70, 75]; // Penyesuaian lebar kolom
                const rowHeight = 18;
                const startX = 30;
                const startY = 150;

                createTableEstimasiDP(doc, tableData, startX, startY, columnWidths, rowHeight);

                doc.font('Arial Font')
                    .fontSize(10).text(text5, 40, 325, { align: 'left' });

                doc.fontSize(10).text(text6, 40, 342, { align: 'left' });
                doc.fontSize(10).text(text7, 40, 407, { align: 'left' });

                doc.fontSize(10).text(text8, 580, 325, { align: 'left' });
                doc.fontSize(10).text(text9, 580, 342, { align: 'left', lineBreak: false });
                doc.fontSize(10).text(text10, 580, 407, { align: 'left', lineBreak: false });

                doc.fontSize(10).text(text11, 40, 427, { align: 'center' });
                doc.fontSize(10).text(text12, 40, 444, { align: 'center' });
                doc.fontSize(10).text(text13, 40, 509, { align: 'center' });

                doc.end();
            } else {
                res.status(404).send('Tidak ada data untuk fakultas ini');
            }
        })
        .catch((error) => {
            console.error('Error reading Excel file:', error);
            res.status(500).send('Error reading Excel file');
        });
});


// Endpoint untuk mendapatkan semua data fakultas
app.get('/fakultas', (req, res) => {
    res.json(fakultasData);
});

// Endpoint untuk mendapatkan program studi (prodi) berdasarkan id fakultas
app.get('/prodi/:idFakultas', (req, res) => {
    const idFakultas = parseInt(req.params.idFakultas);
    const prodiByFakultas = prodiData.filter(prodi => prodi.idFakultas === idFakultas);
    res.json(prodiByFakultas);
});

const PORT = process.env.PORT || 8001;
app.listen(PORT, () => console.log(`Server started on port ${PORT}`));
