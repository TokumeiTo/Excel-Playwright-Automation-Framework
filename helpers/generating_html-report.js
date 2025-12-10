import { runLog } from "./runLog.js";
import fs from "fs";
import path from "path";
import chalk from "chalk";

export async function generateHTMLReport(file, testResults, htmlFilePath = null, totalSeconds = "N/A", browserName) {
    const htmlFile = htmlFilePath || `results/${path.basename(file, ".xlsx")}_report.html`;
    fs.mkdirSync(path.dirname(htmlFile), { recursive: true });
    const minutes = Math.floor(totalSeconds / 60);
    const seconds = totalSeconds % 60;

    let htmlContent = `
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Test Report - ${path.basename(file)}</title>
<style>
    body {
        font-family: "Segoe UI", Arial, sans-serif;
        background: #f4f6f8;
        margin: 20px;
        color: #222;
    }

    h1 {
        text-align: center;
        color: #333;
    }

    table {
        width: 100%;
        border-collapse: collapse;
        margin-bottom: 20px;
        background: white;
        border-radius: 8px;
        overflow: hidden;
        box-shadow: 0 1px 4px rgba(0,0,0,0.1);
    }

    th, td {
        border: 1px solid #ddd;
        padding: 8px 10px;
        text-align: left;
        vertical-align: top;
        word-break: break-word;
        white-space: normal;
        max-width: 600px;
    }

    th {
        background-color: #1976d2;
        color: white;
    }

    tr.pass {
        background-color: #e8f5e9;
    }

    tr.fail {
        background-color: #ffebee;
    }

    details {
        margin: 5px 0;
    }

    summary {
        cursor: pointer;
        font-weight: bold;
        color: #1976d2;
    }

    summary:hover {
        text-decoration: underline;
    }

    /* --- Screenshot thumbnails --- */
    .thumbnail-wrapper {
        display: inline-block;
        margin: 4px;
    }

    .thumbnail {
        width: 100px;
        cursor: pointer;
        border-radius: 4px;
        box-shadow: 0 0 3px rgba(0,0,0,0.3);
        transition: transform 0.2s ease;
    }

    .thumbnail:hover {
        transform: scale(1.05);
    }

    /* --- Modal for enlarged screenshot --- */
    .modal-overlay {
        display: none;
        position: fixed;
        inset: 0;
        background: rgba(0,0,0,0.8);
        z-index: 1000;
        justify-content: center;
        align-items: center;
    }

    .modal-overlay img {
        max-width: 90%;
        max-height: 90%;
        border-radius: 8px;
        box-shadow: 0 0 40px rgba(0,0,0,0.6);
    }

    .modal-close {
        position: fixed;
        top: 20px;
        right: 30px;
        font-size: 30px;
        color: white;
        cursor: pointer;
        z-index: 1001;
        font-weight: bold;
    }

    /* --- Network Table --- */
    .network-table {
        width: 100%;
        border-collapse: collapse;
        margin-top: 10px;
        background: #fafafa;
        border-radius: 6px;
        overflow: hidden;
    }

    .network-table th {
        background-color: #f0f0f0;
        color: #444;
        text-align: left;
        padding: 6px 8px;
        border: 1px solid #ddd;
    }

    .network-table td {
        padding: 6px 8px;
        border: 1px solid #ddd;
        word-break: break-all;
        white-space: normal;
        max-width: 300px;
    }

    /* Scrollable table for long network logs */
    details > table {
        max-height: 300px;
        overflow-y: auto;
        display: block;
    }
</style>

</head>
<body>
<h1>📊 Test Report: ${path.basename(file)}</h1>
<div style="display: flex; justify-content: space-between; font-size:15px; padding: 20px;">
<div>
    <div><strong>Total Run Time:</strong> ${minutes}min ${seconds}s</div>
    <div><strong>Browser:</strong> ${browserName}</div>
</div>
<div><strong>Generated on</strong>: ${new Date()}</div>
</div>
<table>
<thead>
<tr>
<th>Test Case</th>
<th>Result</th>
<th>Outcome</th>
<th>Duration (s)</th>
<th>Screenshots</th>
<th>Device Type</th>
</tr>
</thead>
<tbody>
`;
    let text_color;
    for (const t of testResults) {
        htmlContent += `
        <tr class="${t.result.toLowerCase()}">
            <td><strong> ${t.testCase} </strong></td>
        `

        if (t.result.toLowerCase() === "pass") {
            text_color="default";
            htmlContent += `
            <td><strong style="color:green;">${t.result}</strong></td>
        `;
        } else if (t.result.toLowerCase() === "fail") {
            text_color="red"
            htmlContent += `
            <td><strong style="color: red;">${t.result}</strong></td>
        `;
        } else {
            text_color="yellow"
            htmlContent += `
            <td><strong style="color: gray;">${t.result}</strong></td>
        `;
        }

        htmlContent += `
            <td>${t.outcome}</td>
            <td>${t.duration.toFixed(2)}</td>
            <td>${t.screenshots
                .map(s => {
                    const shotPerColor = s.includes("ERROR") ? "red" : "default"
                    return `
                    <div class="thumbnail-wrapper">
                        <div style="display:flex; position: relative; align-items:center; font-size: 13px;">
                            <div style="transform: rotate(270deg); margin-top: 20px; color: #1976d2; font-weight: bold;">
                                ${t.imgHeight} px
                            </div>
                            <div>
                                <div style="text-align:center; font-weight: bold; color: #1976d2; height: 23px;">
                                    ${t.imgWidth} px
                                </div>
                                <img class="thumbnail" src="${s}" data-full="${s}" alt="screenshot">
                            </div>
                        </div>
                    </div>
                    <div style="color:${shotPerColor}">${s}</div>
                    `;
                })
                .join("")}
            </td>
            <td><i>${t.deviceNames}</i></td>
        </tr>

        <tr><td colspan="6">
        <details>
        <summary>📡 Network Log</summary>
        <table class="network-table">
            <tr>
                <th>ActionType</th>
                <th>Selector/Value</th>
                <th>Method</th>
                <th>URL</th>
                <th>Status</th>
            </tr>
        `;

        for (const net of t.networkLog || []) {
            htmlContent += `
    <tr>
        <td>${net.ActionType || ""}</td>
        <td>${net.SelectorOrValue || ""}</td>
        <td>${net.Method || ""}</td>
        <td>${net.URL || ""}</td>
        <td>${net.NetworkStatus || ""}</td>
    </tr>`;
        }

        htmlContent += `
</table>
</details>
</td></tr>`;
    }

    htmlContent += `
</tbody>
</table>

<!-- Modal container at the end of <body> -->
<div class="modal-overlay" id="modal">
    <span class="modal-close" id="modalClose">&times;</span>
    <img id="modalImg" src="">
</div>

<script>
    const modal = document.getElementById("modal");
    const modalImg = document.getElementById("modalImg");
    const modalClose = document.getElementById("modalClose");

    document.querySelectorAll(".thumbnail").forEach(img => {
        img.addEventListener("click", () => {
            modalImg.src = img.dataset.full; 
            modal.style.display = "flex";
        });
    });

    modalClose.addEventListener("click", () => {
        modal.style.display = "none";
    });

    window.addEventListener("keydown", (e) => {
        if (e.key === "Escape") {
            modal.style.display = "none";
        }
    });
</script>

</body>
</html>`;

    fs.writeFileSync(htmlFile, htmlContent, "utf8");
    console.log(`📄 ${chalk.bold("HTML report saved:")} ${chalk.underline(htmlFile)}`);
    runLog(`📄 HTML report saved: ${htmlFile}`);
}