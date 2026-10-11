import express from "express";
import multer from "multer";
import xlsx from "xlsx";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();

const upload = multer({ dest: "uploads/" });

app.set("view engine", "ejs");

app.get("/", (req, res) => {
    res.render("index");
});

app.post("/exceltojson", upload.fields([
    { name: "excelFile", maxCount: 1 },
    { name: "templateJson", maxCount: 1 }
]), (req, res) => {
    let excelPath = null;
    let jsonPath = null;
    let outputPath = null;

    try {
        if (!req.files || !req.files.excelFile || !req.files.templateJson) {
            return res.status(400).send("Both Excel file and JSON template file are required.");
        }

        excelPath = req.files.excelFile[0].path;
        jsonPath = req.files.templateJson[0].path;

        // 1. Read and parse the template JSON dictionary
        const templateContent = fs.readFileSync(jsonPath, "utf-8");
        const templateData = JSON.parse(templateContent);

        // 2. Read the Excel workbook
        const workbook = xlsx.readFile(excelPath);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // Convert sheet to a 2D array
        const rows = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

        // 3. Extract Metadata from template or Excel top rows
        let instCode = templateData.InstCode;
        if (!instCode && rows[0] && rows[0][2] !== undefined) {
            instCode = String(rows[0][2]).trim();
        }

        // 4. Dynamically locate header row indices and data start row
        let headerRowIndices = [];
        let dataStartIndex = -1;

        for (let r = 0; r < rows.length; r++) {
            const rowStr = JSON.stringify(rows[r] || []);
            // Identify table header rows by scanning for structural layout labels
            if (rowStr.includes("Region") || rowStr.includes("By Gender") || rowStr.includes("Occupation") || rowStr.includes("Delivery") || rowStr.includes("Description") || rowStr.includes("Items")) {
                headerRowIndices.push(r);
            }
            if (headerRowIndices.length > 0 && r > Math.max(...headerRowIndices)) {
                if (rows[r] && rows[r][0] !== undefined && rows[r][0] !== null && String(rows[r][0]).trim() !== "") {
                    const label = String(rows[r][0]).trim();
                    if (!label.toLowerCase().includes("prepared by") && !label.toLowerCase().includes("approved by")) {
                        if (dataStartIndex === -1) dataStartIndex = r;
                    }
                }
            }
        }

        const startRow = dataStartIndex !== -1 ? dataStartIndex : 7;
        let endRow = rows.length - 1;
        for (let r = startRow; r < rows.length; r++) {
            if (rows[r] && rows[r][0]) {
                const label = String(rows[r][0]).trim().toLowerCase();
                if (label.includes("prepared by") || label.includes("approved by") || label.includes("signiture")) {
                    endRow = r - 1;
                    break;
                }
            }
        }

        // 5. Build dynamic multi-level hierarchical headers for each column
        const maxCols = Math.max(...rows.map(row => row ? row.length : 0));
        const columnHierarchies = [];

        // If multiple header rows are identified, forward fill merged category headers horizontally
        const filledHeaderRows = headerRowIndices.map(hRow => {
            const originalRow = rows[hRow] || [];
            const filledRow = [];
            let lastVal = "";
            for (let c = 0; c < maxCols; c++) {
                const val = originalRow[c];
                if (val !== undefined && val !== null && String(val).trim() !== "") {
                    lastVal = String(val).trim();
                    filledRow[c] = lastVal;
                } else {
                    // For multi-level headers, propagate group headers forward if subheaders exist below
                    filledRow[c] = lastVal;
                }
            }
            return filledRow;
        });

        for (let c = 1; c < maxCols; c++) {
            let pathParts = [];
            
            filledHeaderRows.forEach((fRow) => {
                const cellVal = fRow[c];
                if (cellVal !== undefined && cellVal !== null && String(cellVal).trim() !== "") {
                    const text = String(cellVal).trim();
                    // Avoid duplicating identical consecutive parent/child headers
                    if (pathParts.length === 0 || pathParts[pathParts.length - 1] !== text) {
                        pathParts.push(text);
                    }
                }
            });

            columnHierarchies.push({
                colIndex: c,
                pathParts: pathParts
            });
        }

        // 6. Extract values into a lookup map while skipping section rows starting with numbers (e.g., "1.0 ...", "2.1 ...")
        const extractedDataMap = new Map();

        for (let r = startRow; r <= endRow; r++) {
            const rowData = rows[r];
            if (!rowData) continue;

            const rowName = rowData[0] ? String(rowData[0]).trim() : "";
            if (!rowName) continue;

            // STRICT SKIP RULE: If the row label starts with a number followed by a dot or space (e.g., "1.", "1.0", "2.1"), 
            // treat it as a section title and skip the entire row from being parsed as data.
            if (/^\d+(\.\d+)?\s/.test(rowName) || /^\d+\.\s/.test(rowName)) {
                continue;
            }

            // Process data columns for valid rows only
            columnHierarchies.forEach(colInfo => {
                const rawVal = rowData[colInfo.colIndex];
                
                // If empty/null/whitespace, explicitly set to "0", otherwise use correct value
                const valueStr = (rawVal !== undefined && rawVal !== null && String(rawVal).trim() !== "") 
                    ? String(rawVal).trim() 
                    : "0";

                let descriptionParts = [rowName, ...colInfo.pathParts];
                descriptionParts = descriptionParts.filter(p => p && p.trim() !== "");
                
                const description = descriptionParts.join("_");

                extractedDataMap.set(description.trim().toLowerCase(), valueStr);
            });
        }

        // 7. Map extracted items accurately into the template's ReturnItemsList matching descriptions
        if (templateData.ReturnItemsList && Array.isArray(templateData.ReturnItemsList)) {
            templateData.ReturnItemsList = templateData.ReturnItemsList.map((item) => {
                const templateDesc = item._description ? item._description.trim().toLowerCase() : "";
                
                let matchedVal = "0";
                if (extractedDataMap.has(templateDesc)) {
                    matchedVal = extractedDataMap.get(templateDesc);
                } else {
                    // Fallback search with whitespace normalization
                    for (const [desc, val] of extractedDataMap.entries()) {
                        if (desc === templateDesc || desc.replace(/\s+/g, '') === templateDesc.replace(/\s+/g, '')) {
                            matchedVal = val;
                            break;
                        }
                    }
                }

                return {
                    ...item,
                    Value: matchedVal
                };
            });
        }

        // Ensure metadata from template is retained
        templateData.InstCode = instCode;

        // 8. Write and send output JSON file
        outputPath = path.join(__dirname, "uploads", "output.json");
        fs.writeFileSync(outputPath, JSON.stringify(templateData, null, 2), "utf-8");

        res.download(outputPath, "output.json", (err) => {
            if (err) {
                console.error("Error downloading file:", err);
            }
            cleanupFiles([excelPath, jsonPath, outputPath]);
        });

    } catch (error) {
        console.error("Conversion error:", error);
        cleanupFiles([excelPath, jsonPath, outputPath]);
        res.status(500).send("Error processing files: " + error.message);
    }
});

function cleanupFiles(filePaths) {
    filePaths.forEach((filePath) => {
        if (filePath && fs.existsSync(filePath)) {
            try {
                fs.unlinkSync(filePath);
            } catch (e) {
                console.error("Failed to delete file:", filePath, e);
            }
        }
    });
}

app.listen(3000, () => {
    console.log("App listen at port 3000");
});