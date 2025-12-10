import fs from "fs";
import path from "path";

let currentLogFile = null;
let preInitBuffer = [];   // store logs BEFORE the file is known

// Called FIRST time inside runExcelTest
export function initRunLog(logFilePath) {
    const timestamp = new Date().toString().replace(/\sGMT.*$/, "");
    currentLogFile = logFilePath;

    fs.mkdirSync(path.dirname(logFilePath), { recursive: true });
    fs.writeFileSync(currentLogFile, `=== Test Run Log Started at 🌸[${timestamp}]🌸 ===\n`);

    // Flush buffered logs
    for (const msg of preInitBuffer) {
        fs.appendFileSync(currentLogFile, msg);
    }
    preInitBuffer = []; // clear buffer
}

export function runLog(message) {
    const fullMessage = `${message}\n`;

    // Not initialized yet → push to buffer
    if (!currentLogFile) {
        preInitBuffer.push(fullMessage);
        return;
    }

    // Already initialized → write normally
    fs.appendFileSync(currentLogFile, fullMessage);
}
