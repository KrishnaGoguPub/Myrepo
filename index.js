console.log("index.js loaded");

(function () {
  let renamedColumns = {};
  let worksheet;

  if (!tableau.extensions) {
    console.error("Tableau Extensions API not loaded!");
    return;
  }

  tableau.extensions.initializeAsync().then(() => {
    console.log("Extension initialized");
    worksheet = tableau.extensions.dashboardContent.dashboard.worksheets[0];
    console.log("Worksheet:", worksheet.name);
    renderViz();
    setupParameterListeners();

    document.getElementById("refreshButton").addEventListener("click", () => {
      console.log("Manual refresh triggered");
      renderViz(); // Simple refresh like yours, no filter reapplication
    });

    document.getElementById("exportButton").addEventListener("click", () => {
      console.log("Export button clicked");
      worksheet.getSummaryDataAsync().then(data => {
        console.log("Data fetched for export");
        exportToXLSX(data.columns, data.data, worksheet.name);
      }).catch(err => console.error("Error fetching data:", err));
    });
  }).catch(error => {
    console.error("Initialization failed:", error);
  });

  function setupParameterListeners() {
    tableau.extensions.dashboardContent.dashboard.getParametersAsync().then(parameters => {
      parameters.forEach(parameter => {
        parameter.addEventListener(tableau.TableauEventType.ParameterChanged, (event) => {
          console.log(`Parameter ${event.parameterName} changed to:`, event.field.value);
          setTimeout(renderViz, 2000); // Your delay logic
        });
      });
    }).catch(error => console.error("Error fetching parameters:", error));
  }

  function renderViz() {
    worksheet.getSummaryDataAsync().then(data => {
      const columns = data.columns;
      const rows = data.data;

      const header = document.getElementById("tableHeader");
      let headerRow = "<tr>";
      columns.forEach((col, index) => {
        const name = renamedColumns[index] || col.fieldName;
        headerRow += `<th data-index="${index}" contenteditable="true" onblur="updateColumnName(this, '${col.fieldName}')">${name}</th>`;
      });
      headerRow += "</tr>";
      header.innerHTML = headerRow;

      const body = document.getElementById("tableBody");
      let bodyContent = "";
      rows.forEach(row => {
        bodyContent += "<tr>";
        row.forEach(cell => {
          const value = parseFloat(cell.value);
          let style = "";
          if (!isNaN(value)) {
            if (value > 1000) style = 'style="background-color: #ffcccc;"'; // Pink
            else if (value > 500) style = 'style="background-color: #ffffcc;"'; // Yellow
          }
          bodyContent += `<td ${style}>${cell.formattedValue}</td>`;
        });
        bodyContent += "</tr>";
      });
      body.innerHTML = bodyContent;

      adjustColumnWidths();
      autoAdjustColumnWidths();
    }).catch(error => console.error("Error fetching data:", error));
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
    thElements.forEach(th => {
      th.removeEventListener("resize", resizeHandler);
      th.addEventListener("resize", resizeHandler);
    });
  }

  function resizeHandler(event) {
    const th = event.target;
    const index = parseInt(th.getAttribute("data-index"));
    const width = th.offsetWidth;
    document.querySelectorAll(`#dataTable td:nth-child(${index + 1})`).forEach(td => {
      td.style.width = `${width}px`;
      td.style.minWidth = `${width}px`;
    });
    updateTableWidth();
  }

  function autoAdjustColumnWidths() {
    const vizContainer = document.getElementById("vizContainer");
    const panelWidth = vizContainer.offsetWidth;
    const thElements = document.querySelectorAll("#tableHeader th");
    const baseWidth = Math.max(100, Math.floor(panelWidth / thElements.length));

    let totalWidth = 0;
    thElements.forEach((th, index) => {
      const width = Math.max(baseWidth, th.scrollWidth);
      th.style.width = `${width}px`;
      th.style.minWidth = `${width}px`;
      document.querySelectorAll(`#dataTable td:nth-child(${index + 1})`).forEach(td => {
        td.style.width = `${width}px`;
        td.style.minWidth = `${width}px`;
      });
      totalWidth += width;
    });

    document.getElementById("dataTable").style.width = totalWidth > panelWidth ? `${totalWidth}px` : "100%";
  }

  function updateTableWidth() {
    const thElements = document.querySelectorAll("#tableHeader th");
    const totalWidth = Array.from(thElements).reduce((sum, th) => sum + th.offsetWidth, 0);
    const vizContainer = document.getElementById("vizContainer");
    document.getElementById("dataTable").style.width = totalWidth > vizContainer.offsetWidth ? `${totalWidth}px` : "100%";
  }

  function exportToXLSX(columns, rows, worksheetName) {
    const wsData = [];
    const headers = ["Row Index", ...columns.map((col, i) => renamedColumns[i] || col.fieldName)];
    wsData.push(headers);

    rows.forEach((row, index) => {
      const rowData = [(index + 1).toString(), ...row.map((cell, i) => {
        const col = columns[i];
        return col.dataType === "float" || col.dataType === "int" ? cell.value : cell.formattedValue;
      })];
      wsData.push(rowData);
    });

    const ws = XLSX.utils.aoa_to_sheet(wsData);

    // Apply data cell colors
    const range = XLSX.utils.decode_range(ws["!ref"]);
    for (let R = 1; R <= range.e.r; ++R) {
      for (let C = 1; C <= range.e.c; ++C) { // Start at C=1 to skip "Row Index"
        const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
        const cell = ws[cellAddress];
        if (cell && cell.v !== undefined) {
          const cellValue = parseFloat(cell.v.toString().replace(/[^0-9.-]+/g, ""));
          if (!isNaN(cellValue)) {
            cell.s = cell.s || {};
            if (cellValue > 1000) {
              cell.s.fill = { patternType: "solid", fgColor: { rgb: "FFCCCC" } };
              console.log(`Applied pink to ${cellAddress}: ${cellValue}`);
            } else if (cellValue > 500) {
              cell.s.fill = { patternType: "solid", fgColor: { rgb: "FFFFCC" } };
              console.log(`Applied yellow to ${cellAddress}: ${cellValue}`);
            }
          }
        }
      }
    }

    // Apply header colors
    for (let C = 0; C <= range.e.c; ++C) {
      const headerCell = XLSX.utils.encode_cell({ r: 0, c: C });
      const columnName = headers[C].toLowerCase();
      let headerColor;

      switch (columnName) { // Customize to match your Tableau view
        case "sales": headerColor = "D3D3D3"; break; // Darker gray
        case "profit": headerColor = "CCFFCC"; break; // Light green
        case "quantity": headerColor = "CCE5FF"; break; // Light blue
        default: headerColor = "F2F2F2"; // Default light gray
      }

      ws[headerCell].s = {
        font: { bold: true },
        fill: { patternType: "solid", fgColor: { rgb: headerColor } },
        border: { top: { style: "thin" }, bottom: { style: "thin" }, left: { style: "thin" }, right: { style: "thin" } }
      };
      console.log(`Applied header color ${headerColor} to ${headerCell}`);
    }

    ws["!cols"] = headers.map(() => ({ wpx: 100 }));

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, worksheetName);
    const fileData = XLSX.write(wb, { bookType: "xlsx", type: "binary" });
    const blob = new Blob([s2ab(fileData)], { type: "application/octet-stream" });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${worksheetName}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
    console.log("Export complete");
  }

  function s2ab(s) {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
  }
})();
