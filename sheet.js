// filepath: d:\REFINAL\sheet.js
document.addEventListener("DOMContentLoaded", async function () {
    let excelData = [];

    async function loadExcelData() {
        try {
            const response = await fetch("data.xlsx"); // Fetch the Excel file directly
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            const arrayBuffer = await response.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer, { type: "array" });

            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            excelData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

            console.log("Excel Data Loaded", excelData);
        } catch (error) {
            console.error("Error loading Excel file:", error);
        }
    }

    await loadExcelData();

    document.getElementById("searchButton").addEventListener("click", function () {
        const inputText = document.getElementById("inputField").value.trim();
        if (!inputText || excelData.length === 0) {
            document.getElementById("outputArea").innerHTML = "<p>⚠️ Enter a valid input!</p>";
            return;
        }

        let outputHTML = "";
        let found = false;

        console.log("Searching for:", inputText); // Debugging statement

        for (let i = 1; i < excelData.length; i++) {
            console.log("Checking row:", excelData[i]); // Debugging statement
            if (excelData[i][0]?.toString().trim().toLowerCase() === inputText.toLowerCase()) {
                found = true;
                outputHTML += `
                    <table>
                        <tr>
                            <th>Grade</th>
                            <th>Drill Type</th>
                            <th>Drill Bit</th>
                        </tr>
                        <tr>
                            <td>${excelData[i][1] || "N/A"}</td>
                            <td>${excelData[i][2] || "N/A"}</td>
                            <td>${excelData[i][3] || "N/A"}</td>
                        </tr>
                    </table>
                `;
                break;
            }
        }

        document.getElementById("outputArea").innerHTML = found ? outputHTML : "<p>❌ No matching record found!</p>";
    });

    document.getElementById("resetButton").addEventListener("click", function () {
        document.getElementById("inputField").value = "";
        document.getElementById("outputArea").innerText = "Enter a material to see the output.";
    });
});