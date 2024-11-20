document.addEventListener("DOMContentLoaded", () => {
    const tableContainer = document.getElementById("tableContainer");

    // URL of the Google Sheet published as XLSX
    const excelURL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ0f0gLQZ2jTCv8BBBnRXAEAXo1C3vEYDL9qDTh0hdrjgyzScUsidr0Um-NuBXJXda8FM_FRcCbfZaa/pub?output=xlsx";

    // Extract the desired sheet name from the URL query parameters
    const urlParams = new URLSearchParams(window.location.search);
    const line = urlParams.get("line") || "Summary_Line1"; // Default to Summary_Line1 if not specified

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
            displayTable(jsonData); // Display the data in a table
        })
        .catch(error => {
            console.error("Error fetching Excel file:", error);
            tableContainer.innerHTML = "<p>Failed to load data. Please try again later.</p>";
        });

    // Function to display data in an HTML table
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

        tableContainer.appendChild(table);
    }
});
