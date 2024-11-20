let workbook; // Mock workbook to store processed sheet data

// Initialize event listeners on page load
document.addEventListener("DOMContentLoaded", () => {
    setupLineSelectionListener();
    loadGoogleSheetData();
});

// Function to load data from a Google Sheet published as CSV
function loadGoogleSheetData() {
    const googleSheetURL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ0f0gLQZ2jTCv8BBBnRXAEAXo1C3vEYDL9qDTh0hdrjgyzScUsidr0Um-NuBXJXda8FM_FRcCbfZaa/pub?output=csv"; 
    fetch(googleSheetURL)
        .then(response => response.text())
        .then(data => {
            const parsedData = parseCSV(data);

            // Mocking workbook structure with parsed data
            workbook = {
                Summary_Line1: convertToSheet(parsedData.line1),
                Summary_Line2: convertToSheet(parsedData.line2),
            };

            // Validate sheets
            if (workbook.Summary_Line1 && workbook.Summary_Line2) {
                document.getElementById("lineSelection").style.display = "block";

                // Handle URL parameters
                const urlParams = new URLSearchParams(window.location.search);
                const line = urlParams.get("line") || "Summary_Line1";
                const number = urlParams.get("number");

                if (number) {
                    displayNumberDetails(line, parseInt(number));
                } else {
                    displayLineData(line);
                }
            } else {
                alert("Required data not found in Google Sheet.");
            }
        })
        .catch(error => {
            console.error("Error fetching Google Sheet data:", error);
            alert("Failed to load data from the Google Sheet.");
        });
}

// Parse CSV into JSON
function parseCSV(data) {
    const rows = data.split("\n").map(row => row.split(","));
    const headers = rows[0];
    const line1Data = [];
    const line2Data = [];

    rows.slice(1).forEach(row => {
        if (row[headers.indexOf("Line")] === "1") {
            line1Data.push(row);
        } else if (row[headers.indexOf("Line")] === "2") {
            line2Data.push(row);
        }
    });

    return {
        line1: [headers, ...line1Data],
        line2: [headers, ...line2Data],
    };
}

// Convert parsed data into a sheet-like structure
function convertToSheet(data) {
    return data.reduce((acc, row, index) => {
        acc[index] = row;
        return acc;
    }, {});
}

// Function to display data for the selected line
function displayLineData(line) {
    const container = document.getElementById("tableContainer");
    container.innerHTML = ""; // Clear previous data

    const sheet = workbook[line];
    const sheetData = Object.values(sheet);

    if (!sheetData || sheetData.length === 0) {
        alert(`No data found in the sheet "${line}".`);
        return;
    }

    const headers = sheetData[0];
    const rows = sheetData.slice(1);

    // Render table
    const table = createTable(headers, rows);
    container.appendChild(table);

    // Display details for the first row if available
    if (rows.length > 0) {
        displayNumberDetails(line, rows[0][0]);
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
function displayNumberDetails(line, number) {
    const sheet = workbook[line];
    const sheetData = Object.values(sheet);
    const headers = sheetData[0];
    const rows = sheetData.slice(1);

    const matchingRow = rows.find(row => row[0] == number);

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
