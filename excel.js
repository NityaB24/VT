
        document.getElementById('uploadExcel').addEventListener('change', handleFileUpload);
        document.getElementById('companyLogoInput').addEventListener('change', handleLogoUpload);

        function handleFileUpload(event) {
            let file = event.target.files[0];
            let reader = new FileReader();
            reader.onload = function(event) {
                let data = new Uint8Array(event.target.result);
                let workbook = XLSX.read(data, {type: 'array'});
                processWorkbook(workbook);
            };
            reader.readAsArrayBuffer(file);
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
            let rows = addedItemsTable.rows;

            let receiverName = document.getElementById('receiverName').value;
            let receiverAddress = document.getElementById('receiverAddress').value;
            let receiverPhone = document.getElementById('receiverPhone').value;
            let companyLogoSrc = document.getElementById('companyLogo').src;

            let workbook = new ExcelJS.Workbook();
            let sheet = workbook.addWorksheet('Invoice');

            // Add company details
            sheet.addRow([]);
            sheet.addRow(['', 'Vimalnath Traders']);
            sheet.addRow(['', 'L-3,Madhulika Apartment,Opp. Milk Palace,Bhatar Road,Surat']);
            sheet.addRow(['', '9825456405']);
            sheet.addRow([]);
            sheet.addRow([]);
            sheet.addRow([]);
            // Add receiver details
            sheet.addRow(['Company Details']);
            sheet.addRow([receiverName]);
            sheet.addRow([receiverAddress]);
            sheet.addRow([receiverPhone]);
            sheet.addRow([]);
            sheet.eachRow((row) => {
                row.eachCell((cell) => {
                    cell.font = { bold: true };
                });
            });
            
            // Add items table headers
            let headers = Array.from(rows[0].cells).map((cell, index) => index < rows[0].cells.length - 1 ? cell.textContent : null).filter(cell => cell !== null);
            // let headers = Array.from(rows[0].cells).map((cell, index) => index == 2 && index < rows[0].cells.length - 1 ? cell.textContent : null).filter(cell => cell !== null);
            sheet.addRow(headers);
            sheet.addRow([]);
            
            // Add items table rows
            for (let i = 1; i < rows.length; i++) {
                let row = Array.from(rows[i].cells).map((cell, index) => index < rows[i].cells.length - 1 ? cell.textContent : null).filter(cell => cell !== null);
                sheet.addRow(row);
            }

            // Auto size columns for better readability
            sheet.columns.forEach((column, index) => {
                if (index < 1) {
                    column.width = 20; // Set dynamic width for first two columns
                }
                else if(index == 1 || index==4){
                    column.width = 15;
                }
                 else {
                    column.width = 10; // Set fixed width for other columns
                }
            });

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
                                ext: { width: 120, height: 120 }
                            });
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
                        };
                        reader.readAsDataURL(blob);
                    });
                };
            } else {
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
        });
