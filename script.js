let workbook; // Global variable to store the workbook

// Initialize event listeners on page load
document.addEventListener("DOMContentLoaded", () => {
    setupLineSelectionListener();
    loadExcelFromURL();
});

// Function to fetch and parse Excel file from a URL
function loadExcelFromURL() {
    const excelURL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ0f0gLQZ2jTCv8BBBnRXAEAXo1C3vEYDL9qDTh0hdrjgyzScUsidr0Um-NuBXJXda8FM_FRcCbfZaa/pub?output=xlsx";

    fetch(excelURL)
        .then(response => response.arrayBuffer()) // Fetch as binary data
        .then(data => {
            workbook = XLSX.read(data, { type: "array" }); // Parse binary data into workbook

            // Validate required sheets
            const sheetNames = workbook.SheetNames;
            if (sheetNames.includes("Summary_Line1") && sheetNames.includes("Summary_Line2")) {
                document.getElementById("lineSelection").style.display = "block";

                // Default display of Line 1 data
                displayLineData("Summary_Line1");
            } else {
                alert("Sheets 'Summary_Line1' and 'Summary_Line2' not found in the Excel file.");
            }
        })
        .catch(error => {
            console.error("Error loading Excel file:", error);
            alert("Failed to load the Excel file.");
        });
}

// Function to display data for a selected line
function displayLineData(sheetName) {
    const container = document.getElementById("tableContainer");
    container.innerHTML = ""; // Clear previous data

    const sheet = workbook.Sheets[sheetName];
    const sheetData = XLSX.utils.sheet_to_json(sheet, { header: 1 }); // Convert sheet to JSON

    if (!sheetData || sheetData.length === 0) {
        alert(`No data found in the sheet "${sheetName}".`);
        return;
    }

    const headers = sheetData[0];
    const rows = sheetData.slice(1);

    // Render table
    const table = createTable(headers, rows);
    container.appendChild(table);

    // Display details for the first row (if available)
    if (rows.length > 0) {
        displayNumberDetails(sheetName, rows[0][0]);
    }
}

// Function to create and render a table
function createTable(headers, rows) {
    const table = document.createElement("table");
    const caption = document.createElement("caption");
    caption.textContent = "Sheet Data";
    table.appendChild(caption);

    // Create table header
    const headerRow = document.createElement("tr");
    headers.forEach(header => {
        const th = document.createElement("th");
        th.textContent = header;
        headerRow.appendChild(th);
    });
    table.appendChild(headerRow);

    // Create table rows
    rows.forEach(row => {
        const tr = document.createElement("tr");
        row.forEach(cell => {
            const td = document.createElement("td");
            td.textContent = cell || ""; // Handle empty cells
            tr.appendChild(td);
        });
        table.appendChild(tr);
    });

    return table;
}

// Function to display details for a specific number
function displayNumberDetails(sheetName, number) {
    const sheet = workbook.Sheets[sheetName];
    const sheetData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    const headers = sheetData[0];
    const rows = sheetData.slice(1);

    const matchingRow = rows.find(row => row[0] === number);

    if (matchingRow) {
        const detailsContainer = document.getElementById("cellDetails");
        detailsContainer.innerHTML = `<h3>Data for cell: ${number}</h3>`;
        const ul = document.createElement("ul");
        headers.forEach((header, i) => {
            const li = document.createElement("li");
            li.textContent = `${header}: ${matchingRow[i] || "N/A"}`;
            ul.appendChild(li);
        });
        detailsContainer.appendChild(ul);
    } else {
        alert(`No data found for Number: ${number}`);
    }
}

// Event listener for line selection dropdown
function setupLineSelectionListener() {
    document.getElementById("lineSelect").addEventListener("change", () => {
        const selectedLine = getSelectedLine();
        displayLineData(selectedLine);
    });
}

// Helper function to get the selected line
function getSelectedLine() {
    return document.getElementById("lineSelect").value === "line1"
        ? "Summary_Line1"
        : "Summary_Line2";
}
