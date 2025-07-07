<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Image to Excel Converter</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet" />
<script src="https://cdn.jsdelivr.net/npm/tesseract.js@4.0.2/dist/tesseract.min.js"></script>
</head>
<body>
<div class="container mt-5">
  <h1 class="mb-4">Image to Excel Converter</h1>

  <form id="upload-form" method="POST" enctype="multipart/form-data">
    <div class="mb-3">
      <label for="image" class="form-label">Choose Image File (PNG, JPG):</label>
      <input type="file" class="form-control" id="image" name="image" accept="image/*" required />
    </div>

    <button type="submit" class="btn btn-primary" id="convert-btn">Convert to Excel</button>
  </form>

  <div id="progress" class="mt-3" style="display:none;">
    <p>Processing... Please wait.</p>
  </div>

  <div id="result" class="mt-3"></div>
</div>

<script>
const form = document.getElementById('upload-form');
const progressDiv = document.getElementById('progress');
const resultDiv = document.getElementById('result');
const convertBtn = document.getElementById('convert-btn');

let parsedTextLines = [];

form.addEventListener('submit', function(e) {
  e.preventDefault();

  const fileInput = document.getElementById('image');
  if (!fileInput.files.length) {
    alert('Please select an image file!');
    return;
  }

  const file = fileInput.files[0];
  progressDiv.style.display = 'block';
  resultDiv.innerHTML = '';
  convertBtn.disabled = true;

  const reader = new FileReader();
  reader.onload = function(event) {
    const imgData = event.target.result;

    Tesseract.recognize(
      imgData,
      'eng',
      { logger: m => console.log(m) }
    ).then(({ data: { text } }) => {
      console.log('Extracted Text:', text);

      parsedTextLines = parseTextToRowsCols(text);
      showPreviewTable(parsedTextLines);

      resultDiv.innerHTML += `
        <div class="mt-3">
          <button id="confirm-btn" class="btn btn-success me-2">Confirm & Download Excel</button>
          <button id="copy-btn" class="btn btn-secondary">Copy Data</button>
        </div>
      `;

      document.getElementById('confirm-btn').addEventListener('click', sendDataToServer);
      document.getElementById('copy-btn').addEventListener('click', copyDataToClipboard);

      progressDiv.style.display = 'none';
      convertBtn.disabled = false;
    });
  };
  reader.readAsDataURL(file);
});

function parseTextToRowsCols(text) {
  const lines = text.split(/\r\n|\r|\n/);
  let data = [];
  for (let line of lines) {
    line = line.trim();
    if (!line) continue;

    let cols;
    if (line.includes('\t')) {
      cols = line.split('\t');
    } else if (line.includes(',')) {
      cols = line.split(',');
    } else {
      cols = line.split(/\s+/);
    }
    data.push(cols);
  }
  return data;
}

function showPreviewTable(data) {
  let html = '<div class="table-responsive"><table class="table table-bordered table-sm"><tbody>';
  for (let r = 0; r < data.length; r++) {
    html += '<tr>';
    for (let c = 0; c < data[r].length; c++) {
      html += `<td contenteditable="true" data-row="${r}" data-col="${c}">${escapeHtml(data[r][c])}</td>`;
    }
    html += '</tr>';
  }
  html += '</tbody></table></div>';
  resultDiv.innerHTML = html;

  const cells = resultDiv.querySelectorAll('td[contenteditable="true"]');
  cells.forEach(cell => {
    cell.addEventListener('input', (e) => {
      const row = parseInt(e.target.getAttribute('data-row'));
      const col = parseInt(e.target.getAttribute('data-col'));
      parsedTextLines[row][col] = e.target.innerText;
    });
  });
}

function escapeHtml(text) {
  return text.replace(/&/g, "&amp;")
             .replace(/</g, "&lt;")
             .replace(/>/g, "&gt;")
             .replace(/"/g, "&quot;")
             .replace(/'/g, "&#039;");
}

function sendDataToServer() {
  const textToSend = parsedTextLines.map(row => row.join("\t")).join("\n");
  const fileInput = document.getElementById('image');
  const filename = fileInput.files[0].name;

  fetch('upload.php', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({ filename: filename, text: textToSend })
  })
  .then(res => res.json())
  .then(data => {
    if(data.success){
      resultDiv.innerHTML = `<div class="alert alert-success">Excel file created successfully! <a href="${data.file_url}" target="_blank">Download Excel</a></div>`;
    } else {
      resultDiv.innerHTML = `<div class="alert alert-danger">Error: ${data.message}</div>`;
    }
  })
  .catch(err => {
    resultDiv.innerHTML = `<div class="alert alert-danger">Server error: ${err.message}</div>`;
  });
}

function copyDataToClipboard() {
  const textToCopy = parsedTextLines.map(row => row.join("\t")).join("\n");

  if (navigator.clipboard && window.isSecureContext) {
    navigator.clipboard.writeText(textToCopy).then(() => {
      alert('Table data copied to clipboard!');
    }, () => {
      alert('Failed to copy data.');
    });
  } else {
    const textArea = document.createElement('textarea');
    textArea.value = textToCopy;
    document.body.appendChild(textArea);
    textArea.focus();
    textArea.select();

    try {
      const successful = document.execCommand('copy');
      alert(successful ? 'Table data copied to clipboard!' : 'Failed to copy data.');
    } catch (err) {
      alert('Failed to copy data.');
    }

    document.body.removeChild(textArea);
  }
}
</script>
</body>
</html>
