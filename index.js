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

      // Populate table header
      const header = document.getElementById("tableHeader");
      let headerRow = "<tr>";
      columns.forEach((col) => {
        headerRow += `<th>${col.fieldName}</th>`;
      });
      headerRow += "</tr>";
      header.innerHTML = headerRow;

      // Populate table body
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

      // Attach export functionality to the button (now outside the table)
      document.getElementById("exportButton").onclick = () => exportToXLSX(columns, rows, worksheet.name);
    });
  }

  function exportToXLSX(columns, rows, worksheetName) {
    const headers = ["Row Index", ...columns.map((col) => col.fieldName)];
    const dataArray = [headers];
    rows.forEach((row, index) => {
      const rowData = [(index + 1).toString(), ...row.map((cell, i) => {
        const col = columns[i];
        return col.dataType === "float" || col.dataType === "int" ? cell.value : cell.formattedValue;
      })];
      dataArray.push(rowData);
    });

    const ws = XLSX.utils.aoa_to_sheet(dataArray);
    const range = XLSX.utils.decode_range(ws["!ref"]);
    
    // Style header row
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
      if (!ws[cellAddress]) continue;
      ws[cellAddress].s = {
        fill: { patternType: "solid", fgColor: { rgb: "003087" } },
        font: { bold: true, color: { rgb: "FFFFFF" } },
        alignment: { horizontal: "center" },
      };
    }

    // Number formatting for data rows
    for (let row = range.s.r + 1; row <= range.e.r; row++) {
      for (let col = range.s.c + 1; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
        if (!ws[cellAddress]) continue;
        const colType = columns[col - 1].dataType;
        if (colType === "float" || colType === "int") {
          ws[cellAddress].z = "###0";
        }
      }
    }

    ws["!cols"] = [{ wch: 10, hidden: true }, ...columns.map(() => ({ wch: 15 }))];
    
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, worksheetName);
    XLSX.writeFile(wb, `${worksheetName}.xlsx`);
  }
})();
