(function () {
  let renamedColumns = {};
  let worksheet;

  tableau.extensions.initializeAsync().then(() => {
    console.log("Extension initialized");
    worksheet = tableau.extensions.dashboardContent.dashboard.worksheets[0];
    renderViz();
    setupEventListeners(); // Set up listeners for filter and parameter changes
  });

  function setupEventListeners() {
    // Listen for filter changes on the worksheet
    worksheet.addEventListener(tableau.TableauEventType.FilterChanged, (event) => {
      console.log("Filter changed:", event);
      renderViz(); // Re-render table when filter changes
    });

    // Listen for data source refresh
    worksheet.addEventListener(tableau.TableauEventType.DataSourceChanged, (event) => {
      console.log("Data source changed:", event);
      renderViz();
    });

    // Listen for parameter changes on the dashboard
    tableau.extensions.dashboardContent.dashboard.getParametersAsync().then(parameters => {
      parameters.forEach(parameter => {
        parameter.addEventListener(tableau.TableauEventType.ParameterChanged, (event) => {
          console.log("Parameter changed:", event);
          renderViz(); // Re-render table when any parameter changes
        });
      });
    }).catch(error => {
      console.error("Error fetching parameters:", error);
    });
  }

  function renderViz() {
    worksheet.getSummaryDataAsync().then((data) => {
      const columns = data.columns;
      const rows = data.data;

      // Populate table header with editable column names
      const header = document.getElementById("tableHeader");
      let headerRow = "<tr>";
      columns.forEach((col, index) => {
        const currentName = renamedColumns[index] || col.fieldName;
        headerRow += `<th contenteditable="true" data-index="${index}" onblur="updateColumnName(this, '${col.fieldName}')">${currentName}</th>`;
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

      // Attach export functionality
      document.getElementById("exportButton").onclick = () => exportToXLSX(columns, rows, worksheet.name);

      // Adjust column widths after rendering
      adjustColumnWidths();
    }).catch(error => {
      console.error("Error fetching summary data:", error);
    });
  }

  // Function to update column names
  window.updateColumnName = function(element, originalName) {
    const newName = element.textContent.trim() || originalName;
    const index = element.getAttribute("data-index");
    renamedColumns[index] = newName;
    element.textContent = newName;
  };

  // Function to adjust column widths dynamically
  function adjustColumnWidths() {
    const thElements = document.querySelectorAll("#tableHeader th");
    thElements.forEach((th, index) => {
      th.addEventListener("resize", () => {
        const width = th.offsetWidth;
        document.querySelectorAll(`#dataTable td:nth-child(${index + 1})`).forEach(td => {
          td.style.width = `${width}px`;
        });
      });
    });
  }

  function exportToXLSX(columns, rows, worksheetName) {
    const headers = ["Row Index", ...columns.map((col, i) => renamedColumns[i] || col.fieldName)];
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
    
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
      if (!ws[cellAddress]) continue;
      ws[cellAddress].s = {
        fill: { patternType: "solid", fgColor: { rgb: "003087" } },
        font: { bold: true, color: { rgb: "FFFFFF" } },
        alignment: { horizontal: "center" },
      };
    }

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
