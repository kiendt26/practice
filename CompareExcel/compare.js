function compareFiles() {
  const file1 = document.getElementById('file1').files[0];
  const file2 = document.getElementById('file2').files[0];
  if (!file1 || !file2) {
      document.getElementById('output').innerHTML = 'Vui lòng chọn cả hai tệp.';
      return;
  }
  const isFile1Excel = file1.name.endsWith('.xlsx');
  const isFile2Excel = file2.name.endsWith('.xlsx');
  if (isFile1Excel && isFile2Excel) {
      compareExcelFiles(file1, file2);
  } else if (!isFile1Excel && !isFile2Excel) {
      compareCSVFiles(file1, file2);
  } else {
      document.getElementById('output').innerHTML = 'Vui lòng chọn cả hai tệp cùng loại (Excel hoặc CSV).';
  }
}

function compareExcelFiles(file1, file2) {
  const reader1 = new FileReader();
  const reader2 = new FileReader();

  reader1.onload = function(e) {
      const data1 = new Uint8Array(e.target.result);
      const workbook1 = XLSX.read(data1, { type: 'array' });

      reader2.onload = function(e) {
          const data2 = new Uint8Array(e.target.result);
          const workbook2 = XLSX.read(data2, { type: 'array' });

          const output = document.getElementById('output');
          output.innerHTML = '';

          for (let i = 0; i < workbook1.SheetNames.length; i++) {
              const sheetName = workbook1.SheetNames[i];
              const sheet1 = workbook1.Sheets[sheetName];
              const sheet2 = workbook2.Sheets[sheetName];
              if (!sheet2) {
                  output.innerHTML += `Bảng tính "${sheetName}" không tồn tại trong tệp Excel thứ hai.<br>`;
                  continue;
              }

              const data1 = XLSX.utils.sheet_to_json(sheet1);
              const data2 = XLSX.utils.sheet_to_json(sheet2);
              const missing = [];

              for (let j = 0; j < data1.length; j++) {
                  const row1 = data1[j];
                  let found = false;
                  for (let k = 0; k < data2.length; k++) {
                      const row2 = data2[k];
                      if (JSON.stringify(row1) === JSON.stringify(row2)) {
                          found = true;
                          break;
                      }
                  }
                  if (!found) {
                      missing.push(row1);
                  }
              }

              if (missing.length > 0) {
                  output.innerHTML += `<h3>Bảng tính "${sheetName}" thiếu ${missing.length} hàng dữ liệu:</h3>`;
                  const table = createTable(missing);
                  output.appendChild(table);
                  const newWorkbook = XLSX.utils.book_new();
                  const newSheet = XLSX.utils.json_to_sheet(missing);
                  XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Missing Data');
                  const newFile = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'binary' });
                  const blob = new Blob([s2ab(newFile)], { type: 'application/octet-stream' });
                  const url = URL.createObjectURL(blob);
                  output.innerHTML += `<br><a href="${url}" download="missing_data.xlsx" class="btn btn-success mt-3">Tải xuống file Excel với dữ liệu thiếu</a><br>`;
              } else {
                  output.innerHTML += `<h3>Bảng tính "${sheetName}" không có dữ liệu thiếu.</h3>`;
              }
          }
      };
      reader2.readAsArrayBuffer(file2);
  };
  reader1.readAsArrayBuffer(file1);
}

function compareCSVFiles(file1, file2) {
  const reader1 = new FileReader();
  const reader2 = new FileReader();

  reader1.onload = function(e) {
      const data1 = Papa.parse(e.target.result, { header: true, skipEmptyLines: true }).data;

      reader2.onload = function(e) {
          const data2 = Papa.parse(e.target.result, { header: true, skipEmptyLines: true }).data;
          compareCSVData(data1, data2);
      };
      reader2.readAsText(file2);
  };
  reader1.readAsText(file1);
}

function compareCSVData(data1, data2) {
  const output = document.getElementById('output');
  output.innerHTML = '';

  const missing = [];

  for (let i = 0; i < data1.length; i++) {
      const row1 = data1[i];
      let found = false;
      for (let j = 0; j < data2.length; j++) {
          const row2 = data2[j];
          if (JSON.stringify(row1) === JSON.stringify(row2)) {
              found = true;
              break;
          }
      }
      if (!found) {
          missing.push(row1);
      }
  }

  if (missing.length > 0) {
      output.innerHTML += '<h3>Thiếu dữ liệu:</h3>';
      const table = createTable(missing);
      output.appendChild(table);
      const newWorkbook = XLSX.utils.book_new();
      const newSheet = XLSX.utils.json_to_sheet(missing);
      XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Missing Data');
      const newFile = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'binary' });
      const blob = new Blob([s2ab(newFile)], { type: 'application/octet-stream' });
      const url = URL.createObjectURL(blob);
      output.innerHTML += `<br><a href="${url}" download="missing_data.xlsx" class="btn btn-success mt-3">Tải xuống file Excel với dữ liệu thiếu</a><br>`;
  } else {
      output.innerHTML += '<h3>Không có dữ liệu thiếu.</h3>';
  }
}

function createTable(data) {
  const table = document.createElement('table');
  table.className = 'table table-bordered mt-3';
  const thead = document.createElement('thead');
  const tbody = document.createElement('tbody');

  // Create table headers
  const headers = Object.keys(data[0]);
  const tr = document.createElement('tr');
  headers.forEach(header => {
      const th = document.createElement('th');
      th.innerText = header;
      tr.appendChild(th);
  });
  thead.appendChild(tr);


  data.forEach(row => {
      const tr = document.createElement('tr');
      headers.forEach(header => {
          const td = document.createElement('td');
          td.innerText = row[header];
          tr.appendChild(td);
      });
      tbody.appendChild(tr);
  });

  table.appendChild(thead);
  table.appendChild(tbody);
  return table;
}

function s2ab(s) {
  const buf = new ArrayBuffer(s.length);
  const view = new Uint8Array(buf);
  for (let i = 0; i < s.length; i++) {
      view[i] = s.charCodeAt(i) & 0xFF;
  }
  return buf;
}

const script1 = document.createElement('script');
script1.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.0/xlsx.full.min.js';
document.head.appendChild(script1);

const script2 = document.createElement('script');
script2.src = 'https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.3.2/papaparse.min.js';
document.head.appendChild(script2);