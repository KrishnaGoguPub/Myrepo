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
    renderViz();

    document.getElementById("refreshButton").addEventListener("click", () => {
      console.log("Manual refresh triggered");
      renderViz();
    });
    setupParameterListeners();
  }).catch(error => {
    console.error("Initialization failed:", error);
  });

  function setupParameterListeners() {
    tableau.extensions.dashboardContent.dashboard.getParametersAsync().then(parameters => {
      parameters.forEach(parameter => {
        parameter.addEventListener(tableau.TableauEventType.ParameterChanged, (event) => {
          console.log(`Parameter ${event.parameterName} changed to:`, event.field.value);
          setTimeout(renderViz, 2000);
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
        row.forEach(cell => bodyContent += `<td>${cell.formattedValue}</td>`);
        bodyContent += "</tr>";
      });
      body.innerHTML = bodyContent;

      document.getElementById("exportButton").onclick = () => exportToXLSX(columns, rows, worksheet.name);
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
    XLSX.writeFile({ Sheets: { [worksheetName]: ws }, SheetNames: [worksheetName] }, `${worksheetName}.xlsx`);
  }
})();
