<?php
// index.php
?>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Advanced Image/CSV to Excel Converter</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet" />
<script src="https://cdn.jsdelivr.net/npm/tesseract.js@4.0.2/dist/tesseract.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/sweetalert2@11"></script>
<style>
td[contenteditable="true"]:focus { outline: 2px solid #0d6efd; }
</style>
</head>
<body class="bg-light">

<div class="container py-5">
    <div class="card shadow p-4">
        <h2 class="text-center text-primary mb-4">📄 Advanced Image/CSV to Excel Converter</h2>

        <div class="mb-3">
            <label>Upload Image:</label>
            <input type="file" class="form-control" id="image" accept="image/*">
        </div>
        <div class="mb-3">
            <label>Upload CSV:</label>
            <input type="file" class="form-control" id="csv" accept=".csv">
        </div>

        <button id="process-btn" class="btn btn-primary w-100 mb-3">Process</button>

        <div id="progress" class="d-none text-center">
            <div class="spinner-border text-primary"></div>
            <p class="mt-2">Processing... Please wait.</p>
        </div>

        <div id="result"></div>
    </div>
</div>

<script>
const processBtn = document.getElementById('process-btn');
const imageInput = document.getElementById('image');
const csvInput = document.getElementById('csv');
const progress = document.getElementById('progress');
const resultDiv = document.getElementById('result');
let parsedData = [];

processBtn.addEventListener('click', () => {
    if (imageInput.files.length) processImageOCR(imageInput.files[0]);
    else if (csvInput.files.length) processCSV(csvInput.files[0]);
    else Swal.fire('Error', 'Upload an image or CSV.', 'error');
});

function processImageOCR(file) {
    progress.classList.remove('d-none');
    resultDiv.innerHTML = '';
    const reader = new FileReader();
    reader.onload = e => {
        Tesseract.recognize(e.target.result, 'eng', { logger: m => console.log(m) })
        .then(({ data: { tsv } }) => {
            parsedData = parseTSV(tsv);
            displayTable(parsedData);
            progress.classList.add('d-none');
            addActionButtons();
        });
    };
    reader.readAsDataURL(file);
}

function parseTSV(tsv) {
    const lines = tsv.split('\n').slice(1);
    const data = [];
    let currentLine = [], currentTop = null, tolerance = 10;
    for (const line of lines) {
        const parts = line.split('\t');
        if (parts.length < 12) continue;
        const [level, page, block, par, lnum, wnum, left, top, width, height, conf, text] = parts;
        if (!text.trim()) continue;
        if (currentTop === null) currentTop = parseInt(top);
        if (Math.abs(parseInt(top) - currentTop) > tolerance) {
            data.push(currentLine);
            currentLine = [];
            currentTop = parseInt(top);
        }
        currentLine.push(text.trim());
    }
    if (currentLine.length) data.push(currentLine);
    return data;
}

function processCSV(file) {
    progress.classList.remove('d-none');
    resultDiv.innerHTML = '';
    const reader = new FileReader();
    reader.onload = e => {
        parsedData = e.target.result.split(/\r?\n/).filter(Boolean).map(l => l.split(','));
        displayTable(parsedData);
        progress.classList.add('d-none');
        addActionButtons();
    };
    reader.readAsText(file);
}

function displayTable(data) {
    let html = '<div class="table-responsive"><table class="table table-bordered table-sm table-hover"><tbody>';
    data.forEach((row, r) => {
        html += '<tr>';
        row.forEach((cell, c) => {
            html += `<td contenteditable="true" data-row="${r}" data-col="${c}">${escapeHtml(cell)}</td>`;
        });
        html += '</tr>';
    });
    html += '</tbody></table></div>';
    resultDiv.innerHTML = html;

    resultDiv.querySelectorAll('td[contenteditable="true"]').forEach(cell => {
        cell.addEventListener('input', e => {
            const row = parseInt(e.target.dataset.row);
            const col = parseInt(e.target.dataset.col);
            parsedData[row][col] = e.target.innerText;
        });
    });
}

function addActionButtons() {
    resultDiv.innerHTML += `
        <div class="text-center mt-3">
            <button class="btn btn-success me-2" id="download-btn">Download Excel</button>
            <button class="btn btn-secondary" id="copy-btn">Copy Data</button>
        </div>
    `;
    document.getElementById('download-btn').addEventListener('click', downloadExcel);
    document.getElementById('copy-btn').addEventListener('click', copyToClipboard);
}

function escapeHtml(text) {
    return text.replace(/[&<>"']/g, m => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#039;' })[m]);
}

function downloadExcel() {
    Swal.fire({
        title: 'Download?',
        text: 'Generate Excel from data?',
        icon: 'question',
        showCancelButton: true,
        confirmButtonText: 'Yes, download'
    }).then(result => {
        if (result.isConfirmed) {
            const text = parsedData.map(row => row.join("\t")).join("\n");
            const filename = imageInput.files.length ? imageInput.files[0].name : csvInput.files[0].name;
            fetch(imageInput.files.length ? 'upload_image.php' : 'upload_csv.php', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ filename, text })
            })
            .then(res => res.json())
            .then(data => {
                if (data.success) {
                    Swal.fire({
                        title: 'Download Ready',
                        html: `<a href="${data.file_url}" download class="btn btn-primary mt-2">Download Excel</a>`,
                        icon: 'success'
                    });
                } else {
                    Swal.fire('Error', data.message, 'error');
                }
            });
        }
    });
}

function copyToClipboard() {
    const text = parsedData.map(row => row.join("\t")).join("\n");
    navigator.clipboard.writeText(text)
    .then(() => Swal.fire('Copied!', 'Data copied to clipboard.', 'success'))
    .catch(() => Swal.fire('Error', 'Copy failed.', 'error'));
}
</script>
</body>
</html>
