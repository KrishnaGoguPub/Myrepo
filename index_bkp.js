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
        const columns = data.columns; // Array of column metadata (dimensions and measures)
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
  
        // Add export button functionality
        document.getElementById("exportButton").onclick = () => exportToXLSX(columns, rows);
      });
    }
  
    // Function to export data as XLSX with formatting
    function exportToXLSX(columns, rows) {
      // Prepare data for SheetJS
      const headers = columns.map((col) => col.fieldName);
      const dataArray = [headers]; // First row is headers
      rows.forEach((row) => {
        const rowData = row.map((cell, index) => {
          const col = columns[index];
          // Use raw value for measures to keep as numbers, formattedValue for dimensions
          return col.dataType === "float" || col.dataType === "int" ? cell.value : cell.formattedValue;
        });
        dataArray.push(rowData);
      });
  
      // Create a worksheet
      const ws = XLSX.utils.aoa_to_sheet(dataArray);
  
      // Apply formatting
      const range = XLSX.utils.decode_range(ws["!ref"]); // Get the range of cells
  
      // Style header row (blue background)
      for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
        if (!ws[cellAddress]) continue;
        ws[cellAddress].s = {
          fill: { fgColor: { rgb: "FF4F81D3" } }, // Blue color
          font: { bold: true, color: { rgb: "FFFFFFFF" } }, // White text
          alignment: { horizontal: "center" },
        };
      }
  
      // Style alternating rows and ensure measures are numbers with formatting
      for (let row = range.s.r + 1; row <= range.e.r; row++) {
        for (let col = range.s.c; col <= range.e.c; col++) {
          const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
          if (!ws[cellAddress]) continue;
          const colType = columns[col].dataType;
          ws[cellAddress].s = {
            fill: {
              fgColor: { rgb: row % 2 === 0 ? "FFF5F6F5" : "FFFFFFFF" }, // Alternating colors
            },
          };
          // Apply number format to measures (e.g., 123456 -> 123,456)
          if (colType === "float" || colType === "int") {
            ws[cellAddress].z = "#,##0"; // Thousands separator format
          }
        }
      }
  
      // Set column widths (optional, for readability)
      ws["!cols"] = columns.map(() => ({ wch: 15 }));
  
      // Create a workbook and append the styled worksheet
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  
      // Write the file and trigger download
      XLSX.writeFile(wb, "exported_data.xlsx");
    }
  })();
