window.addEventListener('DOMContentLoaded', () => {
    loadDefaultWorkbook();
});

document.getElementById('uploadExcel').addEventListener('change', handleFileUpload);
document.getElementById('companyLogoInput').addEventListener('change', handleLogoUpload);

function loadDefaultWorkbook() {
    fetch('VT_pricelist1.xlsx')
        .then(response => response.arrayBuffer())
        .then(data => {
            let workbook = XLSX.read(data, { type: 'array' });
            processWorkbook(workbook);
        })
        .catch(error => console.error('Error loading default workbook:', error));
}

function handleFileUpload(event) {
    let file = event.target.files[0];

    if (!file) {
        return; // No file selected, do nothing
    }

    // Clear existing table data
    clearTable('excelDataTable');

    let reader = new FileReader();
    reader.onload = function(event) {
        let data = new Uint8Array(event.target.result);
        let workbook = XLSX.read(data, { type: 'array' });
        processWorkbook(workbook);
    };
    reader.readAsArrayBuffer(file);
}

function clearTable(tableId) {
    let table = document.getElementById(tableId).getElementsByTagName('tbody')[0];
    table.innerHTML = ''; // Clear all rows from the table body
}
function handleLogoUpload(event) {
    let file = event.target.files[0];
    let reader = new FileReader();
    reader.onload = function(event) {
        let logo = document.getElementById('companyLogo');
        logo.src = event.target.result;
        logo.style.display = 'block';
    };
    reader.readAsDataURL(file);
}

function processWorkbook(workbook) {
    let worksheet = workbook.SheetNames;
    let table = document.getElementById('excelDataTable').getElementsByTagName('tbody')[0];
    table.innerHTML = '';  // Clear existing table rows

    worksheet.forEach(name => {
        let sheet = XLSX.utils.sheet_to_json(workbook.Sheets[name], {header: 1});
        sheet.forEach((row, rowIndex) => {
            if (rowIndex > 0) { // Skip header row
                let newRow = table.insertRow();
                row.forEach((cellData, cellIndex) => {
                    // if (cellIndex !== 2) { 
                        let newCell = newRow.insertCell();
                        newCell.textContent = cellData;
                    // }
                });
                let addButtonCell = newRow.insertCell();
                let addButton = document.createElement('button');
                addButton.className = 'add-btn';
                addButton.textContent = 'Add';
                addButton.onclick = () => {
                    addItemToTable(row, newRow);
                    addButton.remove();
                };
                addButtonCell.appendChild(addButton);
            }
        });
    });

    document.getElementById('searchInput').addEventListener('input', function() {
        let searchText = this.value.toLowerCase();
        filterContent(searchText);
    });
}

function addItemToTable(row, originalRow) {
    let table = document.getElementById('addedItemsTable').getElementsByTagName('tbody')[0];

    // Add product row
    let newRow = table.insertRow();
    row.forEach((cellData, cellIndex) => {
        if (cellIndex !== 2) { // Exclude column index 2
            let newCell = newRow.insertCell();
            newCell.textContent = cellData;
        }
    });
    let gstRate = 0.18; // GST rate of 18%
    let rateCell = row[4];
    let rateExcludingGst = (parseFloat(rateCell) - (parseFloat(rateCell) * gstRate)).toFixed(2);
    let rateExcludingGstCell = newRow.insertCell();
    rateExcludingGstCell.textContent = rateExcludingGst;

    // Add remove button
    let removeButtonCell = newRow.insertCell();
    let removeButton = document.createElement('button');
    removeButton.className = 'remove-btn';
    removeButton.textContent = 'Remove';
    removeButton.onclick = () => {
        newRow.remove();
        let addButtonCell = originalRow.lastChild;
        let addButton = document.createElement('button');
        addButton.className = 'add-btn';
        addButton.textContent = 'Add';
        addButton.onclick = () => {
            addItemToTable(row, originalRow);
            addButton.remove();
        };
        addButtonCell.appendChild(addButton);
    };
    removeButtonCell.appendChild(removeButton);
}

function filterContent(searchText) {
    let table = document.getElementById('excelDataTable');
    let rows = table.getElementsByTagName('tr');
    for (let i = 1; i < rows.length; i++) { // Skip header row
        let cells = rows[i].getElementsByTagName('td');
        let matches = Array.from(cells).some(cell => cell.textContent.toLowerCase().includes(searchText));
        rows[i].style.display = matches ? '' : 'none';
    }
}

document.getElementById('downloadExcel').addEventListener('click', function() {
    let addedItemsTable = document.getElementById('addedItemsTable');
    let rows = Array.from(addedItemsTable.rows).slice(1); // Exclude header row

    // Sort rows by category name (assuming category name is in the third cell of each row)
    rows.sort((a, b) => {
        let categoryNameA = a.cells[2].textContent.toLowerCase(); // Adjust index based on actual column position
        let categoryNameB = b.cells[2].textContent.toLowerCase(); // Adjust index based on actual column position
        if (categoryNameA < categoryNameB) return -1;
        if (categoryNameA > categoryNameB) return 1;
        return 0;
    });

    let receiverName = document.getElementById('receiverName').value;
    let receiverAddress = document.getElementById('receiverAddress').value;
    let receiverPhone = document.getElementById('receiverPhone').value;
    let companyLogoSrc = document.getElementById('companyLogo').src;

    let workbook = new ExcelJS.Workbook();
    let sheet = workbook.addWorksheet('Invoice');

    // Add company details
    sheet.addRow([]);
    sheet.addRow(['','', 'Vimalnath Traders']).getCell(3).font = { size:13, italic: true , bold:true };
    sheet.addRow(['','', 'L-3, Madhulika Apartment, Opp. Milk Palace']).getCell(3).font = {size:13, italic: true,bold:true };
    sheet.addRow(['','', 'Bhatar Road, Surat-395007']).getCell(3).font = {size:13, italic: true,bold:true };
    sheet.addRow(['','', '8799606997']).getCell(3).font = {size:13, italic: true,bold:true };
    sheet.addRow([]);
    sheet.addRow([]);

    let underlineRow = sheet.addRow([]);
    underlineRow.getCell(1).border = { bottom: { style: 'thick' } };
    underlineRow.getCell(2).border = { bottom: { style: 'thick' } };
    underlineRow.getCell(3).border = { bottom: { style: 'thick' } };
    underlineRow.getCell(4).border = { bottom: { style: 'thick' } };
    underlineRow.getCell(5).border = { bottom: { style: 'thick' } };
    underlineRow.getCell(6).border = { bottom: { style: 'thick' } };
    
    // Add receiver details
    sheet.addRow([]);
    sheet.addRow(['',receiverName]).getCell(2).font={bold:true};
    sheet.addRow(['',receiverAddress]).getCell(2).font={bold:true};
    sheet.addRow(['',receiverPhone]).getCell(2).font={bold:true};
    sheet.addRow([]);

    // Set header row bold
    let headerRow = sheet.addRow(['Sr', 'Product', 'Size', 'Category', 'Rate', 'Rate Excluding GST']);
    headerRow.eachCell((cell) => {
        cell.font = { bold: true, size: 13 }; // Increase font size and make it bold
        cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
    });

    // Add sorted items table rows with serial numbers
    rows.forEach((row, index) => {
        let rowData = [index + 1]; // Add serial number starting from 1
        Array.from(row.cells).slice(0, -1).forEach(cell => { // Exclude the last "Remove" button cell
            rowData.push(cell.textContent);
        });
        let newRow = sheet.addRow(rowData);
        newRow.eachCell((cell) => {
            cell.font= {size:13};
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });
    });

    // Auto size columns for better readability
    sheet.columns.forEach((column, index) => {
        if(index===0){
            column.width = 5;
        }
        else if (index === 1) {
            column.width = 25; // Set dynamic width for the first column
        }else if(index==2){
            column.width = 10;
        }
        else if(index == 5){
            column.width = 18;
        }
        else
        {
            column.width = 15; // Set fixed width for other columns
        }
    });

    sheet.pageSetup = {
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0, // This ensures all columns fit on one page
        paperSize: 9, // A4 size
    };

    // Add company logo
    if (companyLogoSrc) {
        let img = new Image();
        img.src = companyLogoSrc;
        img.onload = function() {
            let canvas = document.createElement('canvas');
            canvas.width = img.width;
            canvas.height = img.height;
            let ctx = canvas.getContext('2d');
            ctx.drawImage(img, 0, 0);
            canvas.toBlob(function(blob) {
                const reader = new FileReader();
                reader.onload = function() {
                    const base64Image = reader.result.split(',')[1];
                    const imageId = workbook.addImage({
                        base64: base64Image,
                        extension: 'png',
                    });
                    sheet.addImage(imageId, {
                        tl: { col: 0.1, row: 0.1 },
                        ext: { width: 173.33, height: 120 }
                    });
                    generateAndDownloadExcel(workbook);
                };
                reader.readAsDataURL(blob);
            });
        };
    } else {
        generateAndDownloadExcel(workbook);
    }
});

function generateAndDownloadExcel(workbook) {
    // Sort by category name and generate Excel file
    workbook.xlsx.writeBuffer().then(buffer => {
        let blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        let url = URL.createObjectURL(blob);
        let a = document.createElement('a');
        a.href = url;
        a.download = 'invoice.xlsx';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    });
}

