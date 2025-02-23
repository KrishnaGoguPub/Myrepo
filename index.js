(function () {
  let renamedColumns = {};
  let worksheet;
  let lastRowCount = 0; // Track the previous row count to detect data changes

  // Initialize the extension
  tableau.extensions.initializeAsync().then(() => {
    console.log("Extension initialized");
    worksheet = tableau.extensions.dashboardContent.dashboard.worksheets[0];
    renderViz(); // Initial render
    setupEventListeners();
  });

  // Set up event listeners for filters, worksheet updates, and parameters
  function setupEventListeners() {
    // Listen for filter changes
    worksheet.addEventListener(tableau.TableauEventType.FilterChanged, (event) => {
      console.log("FilterChanged event:", event);
      renderViz();
    });

    // Listen for worksheet updates (triggered when data refreshes)
    worksheet.addEventListener(tableau.TableauEventType.WorksheetUpdated, (event) => {
      console.log("WorksheetUpdated event:", event);
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
          // Delay fetching data to allow Tableau to update
          setTimeout(() => {
            console.log("Fetching updated data after parameter change...");
            renderViz();
          }, 2000); // 2-second delay
        });
      });
    }).catch(error => {
      console.error("Error fetching parameters:", error);
    });
  }

  // Fetch and render the worksheet data
  function renderViz() {
    console.log("Rendering table with latest data...");
    worksheet.getSummaryDataAsync().then((data) => {
      const columns = data.columns;
      const rows = data.data;

      console.log("Columns:", columns.map(c => c.fieldName));
      console.log("Row count:", rows.length);

      // Only update the table if the row count has changed
      if (rows.length !== lastRowCount) {
        lastRowCount = rows.length;
        updateTable(columns, rows);
      } else {
        console.log("No change in row count, likely same data.");
      }
    }).catch(error => {
      console.error("Error fetching summary data:", error);
    });
  }

  // Update the HTML table with the latest data
  function updateTable(columns, rows) {
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

    // Attach export button functionality (assuming an export function exists)
    document.getElementById("exportButton").onclick = () => exportToXLSX(columns, rows, worksheet.name);
  }

  // Function to update column names (for editable headers)
  window.updateColumnName = function(element, originalName) {
    const newName = element.textContent.trim() || originalName;
    const index = element.getAttribute("data-index");
    renamedColumns[index] = newName;
    element.textContent = newName;
  };

  // Placeholder for export function (implement as needed)
  function exportToXLSX(columns, rows, worksheetName) {
    console.log("Exporting to XLSX:", worksheetName);
    // Add your export logic here, e.g., using SheetJS
  }
})();
