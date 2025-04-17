// Update file name once uploaded
function addFileNameEventListener(inputId, spanId) {
  const fileInput = document.getElementById(inputId);
  const fileNameSpan = document.getElementById(spanId);

  fileInput.addEventListener('change', function() {
    if (fileInput.files.length > 0) {
      fileNameSpan.textContent = fileInput.files[0].name;
    } else {
      fileNameSpan.textContent = '';
    }
  });
}

addFileNameEventListener('projectFile', 'projectFileName');
addFileNameEventListener('inventoryFile', 'inventoryFileName');

// Function to read an Excel file and return parsed data
function readExcelFile(file, callback) {
  const reader = new FileReader();
  reader.onload = function(event) {
    const data = event.target.result;
    const workbook = XLSX.read(data, { type: 'array' });
    callback(workbook);
  };
  reader.readAsArrayBuffer(file);
}

// Function to handle export
function exportExcel() {
  const engineerName = document.getElementById('engineer').value;
  const projectName = document.getElementById('projectName').value;
  const projectFile = document.getElementById('projectFile').files[0];
  const inventoryFile = document.getElementById('inventoryFile').files[0];

  if (!projectFile || !inventoryFile) {
    alert("Please select both project and inventory files.");
    return;
  }

  // Read files
  readExcelFile(projectFile, function(projectWorkbook) {
    readExcelFile(inventoryFile, function(inventoryWorkbook) {
      // Parse the data from both files
      const projectSheet = projectWorkbook.Sheets[projectWorkbook.SheetNames[0]];
      const inventorySheet = inventoryWorkbook.Sheets[inventoryWorkbook.SheetNames[0]];

      const projectData = XLSX.utils.sheet_to_json(projectSheet, { header: 1 });
      const inventoryData = XLSX.utils.sheet_to_json(inventorySheet, { header: 1 });

      // Create a new workbook for the export
      const wb = XLSX.utils.book_new();

      // Prepare the header information (Engineer and Project Name)
      const wsData = [
        ['Engineer', engineerName],
        ['Project', projectName],
        [] // Empty row for separation
      ];

      // Prepare the matched part data
      const matchedData = [];

      // Column names
      matchedData.push(['Item Number', 'Part Name', 'Quantity', 'Part Cost']);

      // Build a map of Part Numbers to Inventory data for quick lookup
      // WARNING: assumes specific column order
      const inventoryMap = {};
      for (let i = 1; i < inventoryData.length; i++) {
        const partNumber = inventoryData[i][0];
        const description = inventoryData[i][2];
        const partCost = inventoryData[i][11];
        inventoryMap[partNumber] = { description, partCost };
      }

      // Loop through project data and match with inventory
      // WARNING: assumes specific column order
      for (let i = 1; i < projectData.length; i++) {
        const hydroNumber = projectData[i][2];
        const quantity = projectData[i][0];

        if (inventoryMap[hydroNumber]) {
          const partInfo = inventoryMap[hydroNumber];
          matchedData.push([hydroNumber, partInfo.description, quantity, partInfo.partCost]);
        }
      }

      // Combine into a single worksheet
      const finalData = wsData.concat(matchedData);
      const ws = XLSX.utils.aoa_to_sheet(finalData);
      XLSX.utils.book_append_sheet(wb, ws, 'Exported Project');

      // Download
      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([wbout], { type: 'application/octet-stream' });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = 'parts_list.xlsx';
      link.click();
    });
  });
}

document.getElementById('exportBtn').addEventListener('click', exportExcel);