document.addEventListener("DOMContentLoaded", () => {
    const tableContainer = document.getElementById("tableContainer");

    // URL of the Google Sheet published as XLSX
    const excelURL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ0f0gLQZ2jTCv8BBBnRXAEAXo1C3vEYDL9qDTh0hdrjgyzScUsidr0Um-NuBXJXda8FM_FRcCbfZaa/pub?output=xlsx";

    // Extract parameters from the URL
    const urlParams = new URLSearchParams(window.location.search);
    const line = urlParams.get("line") || "Summary_Line1"; // Default to Summary_Line1
    const number = parseInt(urlParams.get("number"), 10); // Parse the number parameter

    // Fetch the Excel file and display its data
    fetch(excelURL)
        .then(response => response.arrayBuffer()) // Fetch as binary data
        .then(data => {
            const workbook = XLSX.read(data, { type: "array" }); // Parse Excel data

            // Check if the sheet exists
            if (!workbook.SheetNames.includes(line)) {
                tableContainer.innerHTML = `<p>Sheet "${line}" not found in the Excel file.</p>`;
                return;
            }

            const sheet = workbook.Sheets[line]; // Get the specified sheet
            const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 }); // Convert to JSON (array of arrays)

            // If a number is specified, locate the row; otherwise, display the full table
            if (number && !isNaN(number)) {
                displayRowByNumber(jsonData, number);
            } else {
                displayTable(jsonData);
            }
        })
        .catch(error => {
            console.error("Error fetching Excel file:", error);
            tableContainer.innerHTML = "<p>Failed to load data. Please try again later.</p>";
        });

    // Function to display a specific row by number
    function displayRowByNumber(data, number) {
        if (!data || data.length === 0) {
            tableContainer.innerHTML = "<p>No data found in the Excel file.</p>";
            return;
        }

        const headers = data[0]; // First row as headers
        const rows = data.slice(1); // All rows excluding the headers

        // Locate the row with the specified number in the first column
        const matchingRow = rows.find(row => parseInt(row[0], 10) === number);

        if (matchingRow) {
            const table = document.createElement("table");
            const caption = document.createElement("caption");
            caption.textContent = `Data for Number: ${number}`;
            table.appendChild(caption);

            // Create table header
            const headerRow = document.createElement("tr");
            headers.forEach(header => {
                const th = document.createElement("th");
                th.textContent = header || ""; // Handle empty headers
                headerRow.appendChild(th);
            });
            table.appendChild(headerRow);

            // Create table row for the matching data
            const dataRow = document.createElement("tr");
            matchingRow.forEach(cell => {
                const td = document.createElement("td");
                td.textContent = cell || ""; // Handle empty cells
                dataRow.appendChild(td);
            });
            table.appendChild(dataRow);

            tableContainer.innerHTML = ""; // Clear previous content
            tableContainer.appendChild(table);
        } else {
            tableContainer.innerHTML = `<p>No data found for Number: ${number} in sheet "${line}".</p>`;
        }
    }

    // Function to display the entire sheet in an HTML table
    function displayTable(data) {
        if (!data || data.length === 0) {
            tableContainer.innerHTML = "<p>No data found in the Excel file.</p>";
            return;
        }

        const table = document.createElement("table");
        const caption = document.createElement("caption");
        caption.textContent = `Data from ${line}`;
        table.appendChild(caption);

        // Create table header
        const headers = data[0]; // First row as header
        const headerRow = document.createElement("tr");
        headers.forEach(header => {
            const th = document.createElement("th");
            th.textContent = header || ""; // Handle empty headers
            headerRow.appendChild(th);
        });
        table.appendChild(headerRow);

        // Create table rows
        data.slice(1).forEach(row => {
            const tr = document.createElement("tr");
            row.forEach(cell => {
                const td = document.createElement("td");
                td.textContent = cell || ""; // Handle empty cells
                tr.appendChild(td);
            });
            table.appendChild(tr);
        });

        tableContainer.innerHTML = ""; // Clear previous content
        tableContainer.appendChild(table);
    }
});
