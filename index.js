(function () {
  let renamedColumns = {};
  let worksheet;
  let lastRowCount = 0;

  if (!tableau.extensions) {
    console.error("Tableau Extensions API not loaded!");
    return;
  }

  tableau.extensions.initializeAsync().then(() => {
    console.log("Extension initialized");
    worksheet = tableau.extensions.dashboardContent.dashboard.worksheets[0];
    renderViz();
    setupParameterListeners();

    const refreshButton = document.getElementById("refreshButton");
    if (refreshButton) {
      refreshButton.addEventListener("click", () => {
        console.log("Manual refresh triggered");
        renderViz();
      });
    } else {
      console.warn("Refresh button not found in DOM");
    }
  }).catch(error => {
    console.error("Initialization failed:", error);
  });

  function setupParameterListeners() {
    const dashboard = tableau.extensions.dashboardContent.dashboard;
    dashboard.getParametersAsync().then(parameters => {
      if (parameters.length === 0) {
        console.warn("No parameters found on dashboard");
        return;
      }
      console.log("Parameters found:", parameters.map(p => ({ name: p.name, currentValue: p.currentValue.value })));
      parameters.forEach(parameter => {
        console.log(`Subscribing to ParameterChanged for: ${parameter.name}`);
        parameter.addEventListener(tableau.TableauEventType.ParameterChanged, (event) => {
          console.log(`ParameterChanged event fired - ${event.parameterName} changed to:`, event.field.value);
          setTimeout(() => {
            console.log("Starting refresh process after parameter change...");
            pollForDataChange();
          }, 2000);
        });
      });
    }).catch(error => {
      console.error("Error fetching parameters:", error);
    });
  }

  function pollForDataChange() {
    let attempts = 0;
    const maxAttempts = 10;
    const interval = setInterval(() => {
      attempts++;
      console.log(`Polling attempt ${attempts}/${maxAttempts}...`);
      worksheet.getSummaryDataAsync().then((data) => {
        const rows = data.data;
        console.log("Polling row count:", rows.length);
        if (rows.length !== lastRowCount) {
          clearInterval(interval);
          console.log("Data changed, rendering...");
          renderViz();
        } else if (attempts >= maxAttempts) {
          clearInterval(interval);
          console.log("Max attempts reached, no data change detected. Use manual refresh if needed.");
        }
      }).catch(error => {
        console.error("Error during polling:", error);
        clearInterval(interval);
      });
    }, 1000);
  }

  function renderViz() {
    console.log("Rendering table with latest data...");
    worksheet.getSummaryDataAsync().then((data) => {
      const columns = data.columns;
      const rows = data.data;

      console.log("Columns:", columns.map(c => c.fieldName));
      console.log("Row count:", rows.length);
      lastRowCount = rows.length;

      const header = document.getElementById("tableHeader");
      let headerRow = "<tr>";
      columns.forEach((col, index) => {
        const currentName = renamedColumns[index] || col.fieldName;
        headerRow += `<th contenteditable="true" data-index="${index}" onblur="updateColumnName(this, '${col.fieldName}')">${currentName}</th>`;
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

      const exportButton = document.getElementById("exportButton");
      if (exportButton) {
        exportButton.onclick = () => exportToXLSX(columns, rows, worksheet.name);
      } else {
        console.warn("Export button not found in DOM");
      }

      adjustColumnWidths();
      autoAdjustColumnWidths();
    }).catch(error => {
      console.error("Error fetching summary data:", error);
    });
  }

  window.updateColumnName = function(element, originalName) {
    const newName = element.textContent.trim() || originalName;
    const index = element.getAttribute("data-index");
    renamedColumns[index] = newName;
    element.textContent = newName;
    adjustColumnWidths();
  };

  function adjustColumnWidths() {
    const thElements = document.querySelectorAll("#tableHeader th");
    thElements.forEach((th, index) => {
      // Remove old listeners to prevent duplicates
      th.removeEventListener("resize", resizeHandler);
      th.addEventListener("resize", resizeHandler);
    });
  }

  function resizeHandler(event) {
    const th = event.target;
    const index = parseInt(th.getAttribute("data-index"));
    const newWidth = th.offsetWidth;
    console.log(`Column ${index + 1} resized to: ${newWidth}px`);

    // Sync all td elements in this column
    document.querySelectorAll(`#dataTable td:nth-child(${index + 1})`).forEach(td => {
      td.style.width = `${newWidth}px`;
      td.style.minWidth = `${newWidth}px`; // Ensure consistency
    });

    // Update table width if needed
    updateTableWidth();
  }

  function autoAdjustColumnWidths() {
    const vizContainer = document.getElementById("vizContainer");
    const panelWidth = vizContainer.offsetWidth;
    const thElements = document.querySelectorAll("#tableHeader th");
    const numColumns = thElements.length;
    const baseWidth = Math.max(100, Math.floor(panelWidth / numColumns));

    console.log("Panel width:", panelWidth, "Base column width:", baseWidth);

    let totalWidth = 0;
    thElements.forEach((th, index) => {
      const contentWidth = th.scrollWidth;
      const width = Math.max(baseWidth, contentWidth);
      th.style.width = `${width}px`;
      th.style.minWidth = `${width}px`; // Set initial min-width
      document.querySelectorAll(`#dataTable td:nth-child(${index + 1})`).forEach(td => {
        td.style.width = `${width}px`;
        td.style.minWidth = `${width}px`;
      });
      totalWidth += width;
    });

    updateTableWidth();
  }

  function updateTableWidth() {
    const thElements = document.querySelectorAll("#tableHeader th");
    let totalWidth = 0;
    thElements.forEach(th => {
      totalWidth += th.offsetWidth;
    });
    const dataTable = document.getElementById("dataTable");
    const vizContainer = document.getElementById("vizContainer");
    dataTable.style.width = totalWidth > vizContainer.offsetWidth ? `${totalWidth}px` : "100%";
    console.log("Updated table width:", totalWidth, "Container width:", vizContainer.offsetWidth);
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
          ws[cellAddress].z = "#,##0";
        }
      }
    }

    ws["!cols"] = [{ wch: 10, hidden: true }, ...columns.map(() => ({ wch: 15 }))];
    
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, worksheetName);
    XLSX.writeFile(wb, `${worksheetName}.xlsx`);
  }
})();
