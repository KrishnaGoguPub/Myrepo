(function () {
  let renamedColumns = {};
  let worksheet;

  tableau.extensions.initializeAsync().then(() => {
    console.log("Extension initialized");
    worksheet = tableau.extensions.dashboardContent.dashboard.worksheets[0];
    renderViz();
    setupEventListeners();
  });

  function setupEventListeners() {
    // Listen for filter changes
    worksheet.addEventListener(tableau.TableauEventType.FilterChanged, (event) => {
      console.log("FilterChanged event:", event);
      renderViz();
    });

    // Listen for data source refresh
    worksheet.addEventListener(tableau.TableauEventType.DataSourceChanged, (event) => {
      console.log("DataSourceChanged event:", event);
      renderViz();
    });

    // Listen for summary data changes
    worksheet.addEventListener(tableau.TableauEventType.SummaryDataChanged, (event) => {
      console.log("SummaryDataChanged event:", event);
      renderViz();
    });

    // Listen for parameter changes
    const dashboard = tableau.extensions.dashboardContent.dashboard;
    dashboard.getParametersAsync().then(parameters => {
      console.log("Parameters found:", parameters.map(p => p.name));
      parameters.forEach(parameter => {
        console.log(`Subscribing to ParameterChanged for: ${parameter.name}`);
        parameter.addEventListener(tableau.TableauEventType.ParameterChanged, (event) => {
          console.log(`ParameterChanged event - ${event.parameterName} changed to:`, event.field.value);
          // Trigger a dummy filter change after parameter change
          triggerDummyFilterChange();
        });
      });
    }).catch(error => {
      console.error("Error fetching parameters:", error);
    });
  }

  // Function to trigger a dummy filter change
  function triggerDummyFilterChange() {
    console.log("Triggering dummy filter change...");
    // Get the filters on the worksheet to find a suitable one
    worksheet.getFiltersAsync().then(filters => {
      if (filters.length === 0) {
        console.warn("No filters found on worksheet to toggle. Extension may not refresh.");
        // Fallback: Force renderViz directly with delay
        setTimeout(() => renderViz(), 1000);
        return;
      }

      // Use the first filter (assuming it’s a categorical or range filter)
      const filter = filters[0];
      console.log("Using filter:", filter.fieldName);

      // Depending on filter type, apply a dummy change
      if (filter.filterType === tableau.FilterType.Categorical) {
        worksheet.getSummaryDataAsync().then(data => {
          const column = data.columns.find(col => col.fieldName === filter.fieldName);
          if (column && column.dataType === "string") {
            // Get current values and toggle one
            const currentValues = filter.appliedValues.map(v => v.value);
            const dummyValue = currentValues[0] || "dummy"; // Use first value or a placeholder
            // Clear and reapply to trigger refresh
            worksheet.clearFilterAsync(filter.fieldName).then(() => {
              worksheet.applyFilterAsync(filter.fieldName, [dummyValue], tableau.FilterUpdateType.Replace).then(() => {
                console.log("Dummy filter applied, should trigger refresh...");
                // Revert immediately to minimize impact
                worksheet.applyFilterAsync(filter.fieldName, currentValues, tableau.FilterUpdateType.Replace);
              });
            });
          }
        });
      } else if (filter.filterType === tableau.FilterType.Range) {
        // For range filters (e.g., date), tweak min/max slightly
        const currentMin = filter.minValue ? filter.minValue.value : null;
        const currentMax = filter.maxValue ? filter.maxValue.value : null;
        if (currentMin && currentMax) {
          worksheet.applyFilterAsync(filter.fieldName, { min: currentMin, max: currentMax + 1 }, tableau.FilterUpdateType.Replace).then(() => {
            worksheet.applyFilterAsync(filter.fieldName, { min: currentMin, max: currentMax }, tableau.FilterUpdateType.Replace);
            console.log("Range filter toggled and reverted, should trigger refresh...");
          });
        }
      } else {
        console.warn("Filter type not handled for dummy change:", filter.filterType);
        setTimeout(() => renderViz(), 1000); // Fallback
      }
    }).catch(error => {
      console.error("Error fetching filters for dummy change:", error);
      setTimeout(() => renderViz(), 1000); // Fallback if no filters
    });
  }

  function renderViz() {
    console.log("Rendering table with latest data...");
    worksheet.getSummaryDataAsync().then((data) => {
      const columns = data.columns;
      const rows = data.data;

      console.log("Columns:", columns.map(c => c.fieldName));
      console.log("Row count:", rows.length);

      // Populate table header
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
