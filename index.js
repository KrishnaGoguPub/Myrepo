(function () {
  tableau.extensions.initializeAsync().then(() => {
    console.log("Extension initialized");
    renderViz();
  });

  function renderViz() {
    const worksheet = tableau.extensions.dashboardContent.dashboard.worksheets[0];
    worksheet.getSummaryDataAsync().then((data) => {
      const columns = data.columns;
      const rows = data.data;

      const header = document.getElementById("tableHeader");
      let headerRow = "<tr>";
      columns.forEach((col) => {
        headerRow += `<th>${col.fieldName}</th>`;
      });
      headerRow += "</tr>";
      header.innerHTML = headerRow;

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

      document.getElementById("exportButton").onclick = () => exportToXLSX(columns, rows, worksheet.name);
    });
  }

  function exportToXLSX(columns, rows, worksheetName) {
    // Prepare headers with "Row Index" as the first column
    const headers = ["Row Index", ...columns.map((col) => col.fieldName)];
    const dataArray = [headers];

    // Prepare data with row index
    rows.forEach((row, index) => {
      const rowData = [(index + 1).toString(), ...row.map((cell, i) => {
        const col = columns[i];
        return col.dataType === "float" || col.dataType === "int" ? cell.value : cell.formattedValue;
      })];
      dataArray.push(rowData);
    });

    // Create worksheet
    const ws = XLSX.utils.aoa_to_sheet(dataArray);
    console.log("Worksheet created:", ws);

    const range = XLSX.utils.decode_range(ws["!ref"]);
    console.log("Range:", range);

    // Style header row (dark blue background, white text)
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
      if (!ws[cellAddress]) {
        console.warn(`Cell ${cellAddress} not found`);
        continue;
      }
      ws[cellAddress].s = {
        fill: { patternType: "solid", fgColor: { rgb: "003087" } }, // Dark blue
        font: { bold: true, color: { rgb: "FFFFFF" } }, // White text
        alignment: { horizontal: "center" },
      };
      console.log(`Styled ${cellAddress}:`, ws[cellAddress].s);
    }

    // Apply number formatting to data rows (no color, default white)
    for (let row = range.s.r + 1; row <= range.e.r; row++) {
      for (let col = range.s.c + 1; col <= range.e.c; col++) { // Start at col 1 to skip Row Index
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
        if (!ws[cellAddress]) continue;
        const colType = columns[col - 1].dataType; // Adjust for Row Index offset
        if (colType === "float" || colType === "int") {
          ws[cellAddress].z = "#,##0";
        }
      }
    }

    // Set column widths and hide the Row Index column (column A)
    ws["!cols"] = [{ wch: 10, hidden: true }, ...columns.map(() => ({ wch: 15 }))];

    // Create workbook and export
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, worksheetName);
    console.log("Workbook prepared:", wb);
    XLSX.writeFile(wb, `${worksheetName}.xlsx`);
  }
})();
