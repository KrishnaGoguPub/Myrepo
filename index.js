(function () {
    // Initialize the extension when Tableau is ready
    tableau.extensions.initializeAsync().then(() => {
      console.log("Extension initialized");
      renderViz();
    });
  
    // Function to render the visualization
    function renderViz() {
      const worksheet = tableau.extensions.dashboardContent.dashboard.worksheets[0];
      worksheet.getSummaryDataAsync().then((data) => {
        const columns = data.columns; // Array of column metadata
        const rows = data.data; // Array of row data
  
        // Build table header
        const header = document.getElementById("tableHeader");
        let headerRow = "<tr>";
        columns.forEach((col) => {
          headerRow += `<th>${col.fieldName}</th>`;
        });
        headerRow += "</tr>";
        header.innerHTML = headerRow;
  
        // Build table body
        const body = document.getElementById("tableBody");
        let bodyContent = "";
        rows.forEach((row) => {
          bodyContent += "<tr>";
          row.forEach((cell) => {
            bodyContent += `<td>${cell.formattedValue}</td>`;
          });
          bodyContent += "</tr>";
        });
        body.innerHTML = bodyContent;
  
        // Add export button functionality, passing worksheet name
        document.getElementById("exportButton").onclick = () => exportToXLSX(columns, rows, worksheet.name);
      });
    }
  
    // Function to export data as XLSX with formatting
    function exportToXLSX(columns, rows, worksheetName) {
      // Prepare data for SheetJS
      const headers = columns.map((col) => col.fieldName);
      const dataArray = [headers]; // First row is headers
      rows.forEach((row) => {
        const rowData = row.map((cell, index) => {
          const col = columns[index];
          return col.dataType === "float" || col.dataType === "int" ? cell.value : cell.formattedValue;
        });
        dataArray.push(rowData);
      });
  
      // Create a worksheet
      const ws = XLSX.utils.aoa_to_sheet(dataArray);
      const range = XLSX.utils.decode_range(ws["!ref"]); // Get the range of cells
  
      // Style header row (dark blue background, white text)
      for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
        if (!ws[cellAddress]) continue;
        ws[cellAddress].s = {
          fill: { fgColor: { rgb: "003087" } }, // Dark blue
          font: { bold: true, color: { rgb: "FFFFFF" } }, // White text
          alignment: { horizontal: "center" },
        };
      }
  
      // Style data rows (plain white background)
      for (let row = range.s.r + 1; row <= range.e.r; row++) {
        for (let col = range.s.c; col <= range.e.c; col++) {
          const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
          if (!ws[cellAddress]) continue;
          const colType = columns[col].dataType;
          ws[cellAddress].s = {
            fill: { fgColor: { rgb: "FFFFFF" } }, // White background
          };
          if (colType === "float" || colType === "int") {
            ws[cellAddress].z = "#,##0"; // Thousands separator format
          }
        }
      }
  
      // Set column widths (optional, for readability)
      ws["!cols"] = columns.map(() => ({ wch: 15 }));
  
      // Create a workbook and append the styled worksheet with worksheetName
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, worksheetName); // Use worksheet name as sheet name
  
      // Write the file and trigger download with worksheetName
      XLSX.writeFile(wb, `${worksheetName}.xlsx`);
    }
  })();