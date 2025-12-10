import { chromium, firefox, webkit } from "playwright";
import ExcelJS from "exceljs";
import XLSX from "xlsx";
import fs from "fs";
import path, { resolve } from "path";
import { fileURLToPath } from "url";
import { spawn } from "child_process";
import chalk, { Chalk } from "chalk";
import sharp from "sharp";

// Helpers
import { highlightAndScreenshot } from "../helpers/highlight.js";
import { generateHTMLReport } from "../helpers/generating_html-report.js";
import { getDeviceProfile } from "../helpers/playwright_deviceProfile.js";
import { waitForPaintToSettle } from "../helpers/wait.js";
import { initRunLog, runLog } from "../helpers/runLog.js";
import { getTimestamp } from "../helpers/getTimestamp_YYYYMMDD_HHmmss.js";
import { spinnerFrames, spinnerTick } from "../helpers/consoleLoading.js";
import { rewriteStepErrorForUser } from "../helpers/readableError.js";
import { getImageDimensions } from "../helpers/getImageSize.js";

try {
    // #region   --- Global Scope ---
    const startTime = new Date();
    const timestamp = getTimestamp();
    let totalTests = 0;

    function normalizeName(name) {
        return name
            .replace(/[!./+\-]/g, '') // remove special chars
            .replace(/\s+/g, '_')     // spaces -> underscores
            .replace(/_+$/g, '')      // trim trailing _
            .toLowerCase();
    }
    // #endregion ---

    // #region   --- Arguments Parsing Section ---
    const args = process.argv.slice(2);
    if (args.length < 3) {
        console.error("Usage: node runFromExcel.js <excelFiles> <browser> <headless>");
        process.exit(1);
    }
    const excelFilesRaw = args[0];
    const browserName = args[1];
    const headless = args[2] === "false" ? false : true;
    const excelFiles = excelFilesRaw.split(",").map(f => f.trim().replace(/^"|"$/g, ''));

    // Headless Log
    console.log(
        chalk.bold("🙈 Headless run: ") +
        (headless
            ? chalk.green.bold("TRUE")
            : chalk.red.bold("FALSE")
        )
    );
    runLog(`🙈 Headless run: ${headless ? "TRUE" : "FALSE"}`);

    // Selected Browser Log
    console.log(
        chalk.bold("🌐 Browser: ") + chalk.yellowBright(browserName)
    );
    runLog(`🌐 Browser: ${browserName}`);


    const browserMap = {
        Chrome: (headless) => chromium.launch({ headless }),
        Firefox: (headless) => firefox.launch({ headless }),
        Webkit: (headless) => webkit.launch({ headless }),
        Miscrosoft_Edge: (headless) => chromium.launch({ channel: 'msedge', headless })
    };
    const progressFile = args[3] || path.join(process.env.TEMP || ".", "test_progress.txt");

    // #endregion ---

    // #region   --- Normalize Excel row's keys helper ---
    function normalizeRow(row) {
        const newRow = {};
        for (const key of Object.keys(row)) {
            const trimmedKey = key.trim();
            if (trimmedKey.toLowerCase() === "test cases") {
                newRow.testCase = row[key];
            } else {
                const normalized = trimmedKey.replace(/\s+/g, "").toLowerCase();
                newRow[normalized] = row[key];
            }
        }
        return newRow;
    }
    // #endregion ---

    // #region --- Excel result Sheet Name helper ---
    function sanitizeSheetName(name) {
        // remove invalid chars and trim to 31 chars
        const invalid = /[:\\\/\?\*\[\]]/g;
        let s = (name || "Testcase").replace(invalid, "").trim();
        if (!s) s = "Testcase";
        return s.length > 31 ? s.slice(0, 28).trim() + "..." : s;
    }
    function uniqueSheetName(wb, base) {
        let name = sanitizeSheetName(base);
        let i = 1;
        while (wb.getWorksheet(name)) {
            const suffix = ` (${i++})`;
            const maxBase = 31 - suffix.length;
            name = (sanitizeSheetName(base).slice(0, maxBase) + suffix);
        }
        return name;
    }
    // #endregion ---

    // #region --- Normalize Selector helper ---
    function getSelector(sel) {
        sel = sel.trim().replace(/^"|"$/g, '');
        if (sel.startsWith("text=")) return sel;
        if (sel.startsWith("id=")) return `#${sel.replace("id=", "")}`;
        if (sel.startsWith("name=")) return `[name="${sel.replace("name=", "")}"]`;
        if (sel.startsWith("type=")) return `[type="${sel.replace("type=", "")}"]`;
        if (sel.startsWith("class=")) return `.${sel.replace("class=", "")}`;
        if (sel.startsWith("placeholder=")) return `[placeholder="${sel.replace("placeholder=", "")}"]`;
        if (sel.startsWith("xpath=")) return sel.replace("xpath=", "");
        if (sel.startsWith("fullxpath=")) return sel.replace("fullxpath=", "xpath=");
        return sel;
    }
    // #endregion ---

    // #region   --- Total Steps Counting Section ---
    for (const file of excelFiles) {
        if (!fs.existsSync(file)) continue;

        const workbook = XLSX.readFile(file);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet);

        for (const rowRaw of rows) {
            const row = normalizeRow(rowRaw);

            // Count all actions
            ["goto", "click", "write", "select", "selectbutton", "highlight", "wait", "keyboard", "download", "waitfordocumentloaded"]
                .forEach(key => { if (row[key] && row[key].toString().trim() !== "") totalTests++; });

            // Count expected outcomes exactly like incrementProgress
            if (row.expectedoutcome) {
                const expectedVal = row.expectedoutcome ?? "";  // fallback for undefined/null
                const expectedParts = String(expectedVal)       // convert numbers to string
                    .split("&")
                    .map(e => e.trim())
                    .filter(e => e);

                const loggedExpected = new Set();
                let loggedStay = false;

                for (const trimmedExp of expectedParts) {
                    const key = trimmedExp.toLowerCase();
                    if (!loggedExpected.has(key)) {
                        totalTests++;  // count every unique expected outcome
                        if (key === "stay") loggedStay = true;
                        loggedExpected.add(key);
                    }
                }
            }
        }
    }

    // #endregion ---

    // #region   --- Progress Steps and Setup Section ---
    if (!fs.existsSync(progressFile)) fs.writeFileSync(progressFile, "0", "utf8");

    const __filename = fileURLToPath(import.meta.url);
    const __dirname = path.dirname(__filename);
    const psScriptPath = path.join(__dirname, "../ProgressWindow.ps1");

    let completedTests = 0;
    const useProgress = headless;

    async function incrementProgress(stepDescription, action = 'N/A', error = null, serious = false) {
        completedTests++;
        if (completedTests < totalTests) {
            console.log(`⏳ ${chalk.bold("Step")} ${chalk.yellowBright(completedTests)}/${chalk.green.bold(totalTests)}(${chalk.hex("#A59ACA)").bold(action)}):: ${stepDescription}`);
            runLog(`⏳ Step ${completedTests}/${totalTests}(${action}):: ${stepDescription}`);
        } else if (completedTests === totalTests) {
            console.log(`⌛ ${chalk.bold("Step")} ${chalk.greenBright(completedTests)}/${chalk.green.bold(totalTests)}(${chalk.hex("#A59ACA").bold(action)}):: ${stepDescription}`);
            runLog(`⌛ Step ${completedTests}/${totalTests}(${action}):: ${stepDescription}`);
        }

        if (error && !serious) {
            const userFriendlyError = rewriteStepErrorForUser(stepDescription, error);
            console.warn(chalk.yellow(`⚠️ ${chalk.bold(`Step ${completedTests} Issue:`)} ${userFriendlyError}`));
            runLog(`⚠️ Step ${completedTests} (ISSUE): ${userFriendlyError}`);
        } else if (error && serious) {
            const userFriendlyError = rewriteStepErrorForUser(stepDescription, error);
            console.warn(chalk.redBright(`❌ ${chalk.bold(`Step ${completedTests} Error:`)} ${userFriendlyError}`));
            runLog(`❌ Step ${completedTests} (ERROR): ${userFriendlyError}`);
        }

        if (useProgress && progressFile) {
            fs.writeFileSync(progressFile, completedTests.toString(), "utf8");
        }
    }
    // --- Kill progress bar function ---
    async function killProgressBar() {
        if (global.psProcess && !global.psProcess.killed) {
            global.psProcess.kill();
            await new Promise(r => setTimeout(r, 300)); // ensure it closes
            global.progressWindowKilledBySystem = true; // flag for system kill
        }
    }
    // #endregion ---

    // #region   --- Main execution ---
    (async () => {
        fs.mkdirSync(path.dirname(progressFile), { recursive: true });


        // --- Launch PowerShell progress window ---
        if (headless) {
            console.log("💻 Launching PowerShell progress window...");

            global.progressWindowClosed = false;
            global.progressWindowKilledBySystem = false; // NEW: track system kills

            global.psProcess = spawn("powershell.exe", [
                "-NoProfile",
                "-ExecutionPolicy", "Bypass",
                "-File", `"${psScriptPath}"`,
                "-TotalTests", totalTests,
                "-TempFile", `"${progressFile}"`
            ], { shell: true, stdio: "ignore" });

            global.psProcess.on("close", () => {
                global.progressWindowClosed = true;

                let status = "UNKNOWN";
                if (fs.existsSync(progressFile)) {
                    status = fs.readFileSync(progressFile, "utf8").trim();
                }

                if (status === "DONE") {
                    console.log(chalk.bold(`🎉 All tests ${chalk.greenBright.bold("completed!")} Progress marked ${chalk.greenBright.bold('100%')}`));
                    runLog(`🎉 All tests completed! Progress marked at 100%`);
                } else {
                    console.log(chalk.bold(`🛑 Progress bar was closed! Procress will be stopped executing...`));
                    runLog(`🛑 Progress bar was closed! Procres will be stopped executing...`);
                }
            });
        } else {
            console.log("❕ Running in headed mode. Progress window skipped.");
            runLog("❕ Progress window got skipped or closed because of headed mode.");
        }

        // --- Run Excel files ---
        await Promise.all(
            excelFiles.map(async (file) => {
                console.log(`📂 ${chalk.bold('Test file:')} ${chalk.underline(file)}`);
                runLog(`📂 Test file: ${file}`);
                await runExcelTest(file, browserName, headless);
                console.log(`✅ ${chalk.bold('Finished test file:')} ${chalk.underline(file)}`);
                runLog(`✅ Finished test file: ${file}`);
            })
        );

        // --- Finalize Progress Bar ---
        if (!headless) {
            console.log(chalk.bold(`🎉 All tests ${chalk.greenBright.bold("completed!")} Progress marked ${chalk.greenBright.bold('100%')}`));
            runLog(`🎉 All tests completed! Progress marked 100% \n`);
        }
    })();
    // #endregion ---

    // #region   --- Run From Excel Main Section ---
    async function runExcelTest(file, browserName, headless) {
        console.log(`📄 ${chalk.bold("Progress file initialized at")}: ${chalk.underline(progressFile)}`);
        runLog(`📄 Progress file initialized at: ${progressFile}`);
        console.log(`ℹ️ ${chalk.bold("Total steps to run:")}`, totalTests);
        runLog(`ℹ️ Total steps to run: ${totalTests}`);
        const networkLogs = [];
        const testResults = [];

        if (!fs.existsSync(file)) {
            console.error(`File not found: ${file}`);
            return;
        }

        const workbook = XLSX.readFile(file);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet);

        const baseName = path.basename(file, path.extname(file));
        const testResultDir = path.join("results", `${baseName}_${timestamp}_results`);

        // Subfolders
        const screenshotsDir = path.join(testResultDir, "screenshots");
        const downloadsDir = path.join(testResultDir, "downloads");

        const xslxResultFileName = `${baseName}_${timestamp}_result.xlsx`;
        const htmlResultFileName = `${baseName}_${timestamp}_result.html`;
        const logFileName = `${baseName}_${timestamp}_runLog.txt`;
        const excelResultPath = path.join(testResultDir, xslxResultFileName);
        const htmlReportPath = path.join(testResultDir, htmlResultFileName);
        const logFilePath = path.join(testResultDir, logFileName);
        // Initialize the log file
        initRunLog(logFilePath);

        // First log
        runLog(`💨 Test started: ${baseName}\n`);

        fs.mkdirSync(screenshotsDir, { recursive: true });
        fs.mkdirSync(downloadsDir, { recursive: true });


        const browser = await browserMap[browserName](headless);
        let context = await browser.newContext({ acceptDownloads: true });
        const page = await context.newPage();

        let lastNetwork = {
            status: "N/A",
            method: "N/A",
            url: "N/A",
            time: 0
        };

        page.on("response", async (response) => {
            try {
                const request = response.request();
                lastNetwork = {
                    status: response.status(),
                    method: request.method(),
                    url: response.url(),
                    time: Date.now()
                };
            } catch {
                lastNetwork = { status: "N/A", method: "N/A", url: "N/A", time: Date.now() };
            }
        });

        page.on("requestfailed", (request) => {
            lastNetwork = {
                status: "N/A",
                method: request.method(),
                url: request.url(),
                time: Date.now()
            };
        });

        // --- Capture console messages for this test case ---
        const consoleMessages = [];
        page.on("console", msg => {
            consoleMessages.push({ type: msg.type(), text: msg.text() });
        });


        // --- ExcelJS results workbook ---
        const resultsWorkbook = new ExcelJS.Workbook();

        const wsNetwork = resultsWorkbook.addWorksheet("NetworkLog");
        wsNetwork.columns = [
            { header: "TestCase", key: "TestCase", width: 30 },
            { header: "ActionType", key: "ActionType", width: 15 },
            { header: "Selector/Value", key: "SelectorOrValue", width: 40 },
            { header: "Method", key: "Method", width: 10 },
            { header: "URL", key: "URL", width: 60 },
            { header: "NetworkStatus", key: "NetworkStatus", width: 15 },
        ];

        const wsConsole = resultsWorkbook.addWorksheet("Console Captures");
        wsConsole.columns = [
            { header: "TestCase", key: "TestCase", width: 30 },
            { header: "ConsoleType", key: "ConsoleType", width: 15 },
            { header: "Message", key: "Message", width: 80 },
        ];

        // --- Run each row ---
        for (const rowRaw of rows) {
            const row = normalizeRow(rowRaw);
            const testCaseName = (row.testCase || "").toString().trim();
            const executedActions = [];


            // Create sheet per test case
            const sheetName = testCaseName.substring(0, 31); // Excel limit
            const ws = resultsWorkbook.addWorksheet(sheetName);

            let rowIndex = 10;

            wsNetwork.addRow({
                TestCase: testCaseName,
                ActionType: "Click",
                SelectorOrValue: "#submit",
                Method: "POST",
                URL: "https://example.com/api",
                NetworkStatus: 200
            });

            wsConsole.addRow({
                TestCase: testCaseName,
                ConsoleType: "log",
                Message: "Button clicked successfully"
            });

            // --- Array to store all screenshots ---
            const screenshots = [];

            if (global.progressWindowClosed) {
                const pct = totalTests > 0
                    ? Math.round((completedTests / totalTests) * 100)
                    : 0;
                console.log(chalk.bold(`🛑 Testing was stopped at ${pct}% (${completedTests}/${totalTests}) Before reaching: TestCase(${testCaseName})`));
                runLog(`🛑 Testing was stopped at ${pct}% (${completedTests}/${totalTests}) completed!`);
                break;
            }

            console.log(`---------- 🎌 ${chalk.bold("TestCase")}: (${chalk.cyan(testCaseName)}) ----------`);
            runLog(`---------- 🎌 TestCase to run: (${testCaseName}) ----------`);
            // --- Multi-device handling ---
            const deviceType = row.devicetype?.trim().toLowerCase();  // Get the device type from Excel (ensure it's lowercased)
            console.log(`📱 ${chalk.bold("Device Type specified")}: ${chalk.blueBright(deviceType)}`);
            runLog(`📱 Device Type specified "${testCaseName}": ${deviceType}`);
            // Map it to the corresponding device in the deviceMap
            const deviceProfile = getDeviceProfile(deviceType);

            context = await browser.newContext({
                userAgent: deviceProfile.userAgent,
                viewport: deviceProfile.viewport,  // Explicitly setting the viewport size
                deviceScaleFactor: deviceProfile.deviceScaleFactor,
                // isMobile: deviceProfile.isMobile,
                hasTouch: deviceProfile.hasTouch,
                acceptDownloads: true
            });
            await page.setViewportSize(deviceProfile.viewport); // Update viewport for the page

            let logEntry
            // --- Network logging helper ---
            function logNetwork(actionType, selectorOrValue) {
                const now = Date.now();
                const age = now - lastNetwork.time;
                const isFresh = age < 1500;

                logEntry = {
                    TestCase: testCaseName,
                    ActionType: actionType,
                    SelectorOrValue: selectorOrValue,
                    Method: isFresh ? lastNetwork.method : "N/A",
                    URL: isFresh ? lastNetwork.url : "N/A",
                    NetworkStatus: isFresh ? lastNetwork.status : "N/A"
                };

                // Store for HTML Export
                networkLogs.push(logEntry);
            }

            const startTime = Date.now();
            const record = {
                TestCase: testCaseName,
                Result: "Pass",
                Network: "",
                Method: "",
                URL: "",
                Outcome: "",
                Duration: 0,
            };

            try {
                // --- ScreenshotOnly helper ---
                async function screenshotOnly(page, path) {
                    try { await page.screenshot({ path, fullPage: true }); } catch { }
                }
                async function takeScreenshot(actionName) {
                    // --- Ensure screenshotsDir exists and is absolute
                    const screenshotsDirAbs = path.resolve(screenshotsDir);
                    if (!fs.existsSync(screenshotsDirAbs)) {
                        fs.mkdirSync(screenshotsDirAbs, { recursive: true });
                    }

                    // --- Create file name safe for Excel/FS
                    const fileName = `${testCaseName.replace(/\s+/g, "_")}_${actionName.replace(/\s+/g, "_")}.png`;

                    // --- Full absolute path
                    const fullShotPath = path.resolve(screenshotsDirAbs, fileName);

                    // --- Take screenshot using Playwright
                    await screenshotOnly(page, fullShotPath);

                    // --- Store in array for later embedding
                    screenshots.push(fullShotPath);

                    // --- Store in record.Screenshot (multi-line)
                    if (!record.Screenshot) record.Screenshot = "";
                    record.Screenshot += (record.Screenshot ? "\n" : "") + fullShotPath;

                    // --- Optional: sanity check file exists
                    if (!fs.existsSync(fullShotPath)) {
                        console.warn("⚠️ Screenshot file not found after capture:", fullShotPath);
                    }

                    return fullShotPath;
                }

                // --- Action helpers ---
                async function safeGoto(url) {
                    try {
                        const resp = await page.goto(url, { waitUntil: "domcontentloaded" });
                        record.Method = "GET";
                        record.URL = url;
                        record.Network = resp?.status() || "N/A";
                        await page.waitForLoadState("networkidle");
                        await waitForPaintToSettle(page);
                        logNetwork("GoTo", url);
                        executedActions.push("GoTo");
                        await incrementProgress(`Navigated to: ${url}`, "Go To");
                    } catch (err) {
                        await incrementProgress(`Navigated to: ${url}`, "Go To", err, true);
                        await killProgressBar();
                    }
                }

                async function safeWrite(writeStr) {
                    const [sel, value] = writeStr.split(":");
                    const selector = getSelector(sel);
                    await page.waitForSelector(selector, { timeout: 10000, state: "visible" });
                    await page.fill(selector, value.trim());
                    logNetwork("Write", writeStr);
                    executedActions.push("Write");
                    await incrementProgress(`Filled input: ${writeStr}`, "Write");
                }

                async function safeKeyboard(keyStr) {
                    // Split by '+' for combos like "Ctrl+Shift+I"
                    const keys = keyStr.split("+").map(k => k.trim());

                    // Map friendly names to Playwright keys
                    const keyMap = {
                        "enter": "Enter",
                        "tab": "Tab",
                        "del": "Delete",
                        "delete": "Delete",
                        "esc": "Escape",
                        "escape": "Escape",
                        "backspace": "Backspace",
                        "arrowup": "ArrowUp",
                        "arrowdown": "ArrowDown",
                        "arrowleft": "ArrowLeft",
                        "arrowright": "ArrowRight",
                        "ctrl": "Control",
                        "control": "Control",
                        "shift": "Shift",
                        "alt": "Alt",
                        "meta": "Meta" // Cmd on Mac
                    };

                    const mappedKeys = keys.map(k => keyMap[k.toLowerCase()] || k);

                    if (mappedKeys.length === 1) {
                        // Single key press
                        await page.keyboard.press(mappedKeys[0]);
                        record.Outcome += `✅ Pressed '${mappedKeys[0]}' successfully. `;
                    } else {
                        // Key combination: hold all except last, then press last, then release all
                        const combo = mappedKeys.slice(0, -1);
                        const lastKey = mappedKeys[mappedKeys.length - 1];

                        // Hold modifier keys
                        for (const k of combo) await page.keyboard.down(k);

                        // Press last key
                        await page.keyboard.press(lastKey);

                        // Release modifier keys
                        for (const k of combo.reverse()) await page.keyboard.up(k);

                        record.Outcome += `✅ Pressed key combination '${mappedKeys.join("+")}' successfully. `;
                    }

                    await waitForPaintToSettle(page);
                    logNetwork("Keyboard", keyStr);
                    executedActions.push("Keyboard");
                    await incrementProgress(`Pressed key(s): ${mappedKeys.join("+")}`, "Keyboard");
                }


                async function safeDownload(selStr) {
                    try {
                        const selector = getSelector(selStr);
                        await page.waitForSelector(selector, { timeout: 15000 });

                        // Start listening for the download event
                        const [download] = await Promise.all([
                            page.waitForEvent("download"),
                            page.click(selector) // Click the download button
                        ]);

                        // Define the save path
                        const suggestedName = await download.suggestedFilename();
                        const downloadPath = path.join(downloadsDir, suggestedName);

                        // Ensure folder exists
                        fs.mkdirSync(path.dirname(downloadPath), { recursive: true });

                        // Save file
                        await download.saveAs(downloadPath);

                        // Verify file actually exists
                        if (fs.existsSync(downloadPath)) {
                            logNetwork("Download", selStr);
                            executedActions.push("Download");
                            record.Outcome += `✅ File downloaded: ${suggestedName}. `;
                            await incrementProgress(`Download file: ${suggestedName}`, "Download");
                        } else {
                            logNetwork("Download", selStr);
                            record.Result = "Fail";
                            record.Outcome += `❌ Download failed or missing file (${suggestedName}). `;
                            await incrementProgress(`Download file: ${suggestedName} wasn't completed`, "Download");
                        }
                    } catch (err) {
                        record.Result = "Fail";
                        record.Outcome += `❌ Download error: ${err.message}. `;
                        console.warn(`⚠️ ${chalk.yellowBright(`Download failed: ${err.message}`)}`);
                        runLog(`⚠️ Download failed: ${err.message}`);
                    }
                }

                async function safeClick(selStr) {
                    const selector = getSelector(selStr);
                    await page.waitForSelector(selector, { timeout: 10000, state: "visible" });
                    await waitForPaintToSettle(page);
                    await page.click(selector);
                    logNetwork("Click", selStr);
                    executedActions.push("Click");
                    await incrementProgress(`Clicked: ${selStr}`, "Click");
                }

                async function safeSelect(selectStr) {
                    const [sel, option] = selectStr.split(":");
                    const selector = getSelector(sel);
                    await page.waitForSelector(selector, { timeout: 15000 });
                    await page.selectOption(selector, { label: option.trim() });
                    logNetwork("Select", selectStr);
                    executedActions.push("Select");
                    await incrementProgress(`Selected: ${option} in ${selectStr}`, "Select");
                }

                async function safeSelectButton(selStr) {
                    executedActions.push("SelectButton");

                    const outlineColor = (row["outlinecolor"] || "red").trim() || "red";
                    const [selectorRaw, optionText] = selStr.split(":").map(s => s.trim());
                    const selector = getSelector(selectorRaw);

                    await page.waitForSelector(selector, { timeout: 15000 });
                    await page.click(selector);

                    logNetwork("SelectButton (Open)", selStr);
                    if (optionText) {
                        const optionSelector = `text=${optionText}`;
                        await page.waitForSelector(optionSelector, { timeout: 10000 });
                        await page.click(optionSelector);
                        logNetwork("SelectButton (Choose)", optionText);

                        // --- After choosing
                        if (row.highlight) {
                            // --- Prepare absolute screenshot path
                            const fileName = `${testCaseName.replace(/\s+/g, "_")}_selectbutton_choose.png`;
                            const fullShotPath = path.resolve(screenshotsDir, fileName);

                            // --- Take highlight screenshot(s)
                            const { screenshots: chooseShots, outcomeMessage } = await highlightAndScreenshot(
                                page,
                                optionSelector,
                                fullShotPath,
                                outlineColor
                            );

                            // --- Store absolute paths in global screenshots array
                            screenshots.push(...chooseShots);

                            // --- Store in record.Screenshot (multi-line)
                            if (!record.Screenshot) record.Screenshot = "";
                            record.Screenshot += (record.Screenshot ? "\n" : "") + chooseShots.join("\n");

                            // --- Optional: sanity check files exist
                            chooseShots.forEach(shot => {
                                if (!fs.existsSync(shot)) {
                                    console.warn("⚠️ Screenshot missing:", shot);
                                }
                            });
                        }
                    }

                    await incrementProgress(`Selected Button: ${selStr}`, "SelectButton");
                }

                async function safeWait(ms) {
                    const waitTime = parseInt(ms || "0", 10) || 3000;
                    if (waitTime <= 0) return;

                    executedActions.push("Wait");

                    let spinnerIndex = 0;

                    // update one line
                    const interval = setInterval(() => {
                        process.stdout.write(`\r${spinnerFrames[spinnerIndex]} Waiting for ${waitTime}ms...`);
                        spinnerIndex = (spinnerIndex + 1) % spinnerFrames.length;
                    }, 120);

                    // actual wait
                    await page.waitForTimeout(waitTime);

                    // stop spinner
                    clearInterval(interval);
                    process.stdout.write(chalk.greenBright('\r\x1b[K✔ Waiting complete!\n'));

                    // increment progress for your framework
                    await incrementProgress(`Waited: ${waitTime}ms`, "Wait");
                }


                // --- Run actions ---
                if (row.goto) await safeGoto(row.goto);
                const currentUrlBefore = page.url();
                if (row.waitfordocumentloaded) {
                    const shouldWait = row.waitfordocumentloaded.toString().trim().toLowerCase() === "true";
                    if (shouldWait) {
                        await page.waitForLoadState("load");
                        logNetwork("DocumentLoaded", "true");
                        executedActions.push("DocumentLoaded");

                        await incrementProgress(`Waited for document to fully load!`, "Wait For Document Loaded");
                    }
                }
                if (row.write) await safeWrite(row.write);
                if (row.highlight) {
                    // --- Get outline color from Excel, fallback to red
                    const outlineColor = (row["outlinecolor"] || "red").trim() || "red";

                    // --- Prepare full absolute screenshot path
                    const fileName = `${testCaseName.replace(/\s+/g, "_")}_highlight.png`;
                    const fullShotPath = path.resolve(screenshotsDir, fileName);

                    // --- Get the selector to highlight
                    const selector = getSelector(row.highlight);

                    // --- Take highlight screenshots
                    const { screenshots: shots, outcomeMessage } = await highlightAndScreenshot(
                        page,
                        selector,
                        fullShotPath,
                        outlineColor
                    );

                    // --- Store absolute paths in global screenshots array
                    screenshots.push(...shots);

                    // --- Store in record.Screenshot (multi-line)
                    if (!record.Screenshot) record.Screenshot = "";
                    record.Screenshot += (record.Screenshot ? "\n" : "") + shots.join("\n");

                    // --- Log network / executed actions
                    logNetwork("Highlight", row.highlight);
                    executedActions.push("Highlight");

                    // --- Update outcome
                    record.Outcome += (record.Outcome ? "\n" : "") + outcomeMessage;

                    // --- Update progress
                    await incrementProgress(`Highlighted: ${row.highlight}`, "Highlight");

                    // --- Optional: sanity check files exist
                    shots.forEach(shot => {
                        if (!fs.existsSync(shot)) {
                            console.warn("⚠️ Highlight screenshot missing:", shot);
                        }
                    });
                }
                else if (row.screenshot == true) await takeScreenshot("screenshot");
                if (row.keyboard) {
                    // pre-keyboard screenshot
                    await takeScreenshot(`Before_${row.keyboard}`);

                    // send keyboard action
                    await Promise.all([
                        page.waitForNavigation({ waitUntil: "domcontentloaded" }),
                        safeKeyboard(row.keyboard)
                    ]);

                    // post-action screenshot
                    await takeScreenshot(`After_${row.keyboard}`);
                }
                if (row.download) await safeDownload(row.download);
                if (row.select) await safeSelect(row.select);
                if (row.selectbutton) await safeSelectButton(row.selectbutton);
                if (row.click) await safeClick(row.click);
                if (row.expectedoutcome) {
                    const expectedParts = row.expectedoutcome
                        .split("&")
                        .map(e => e.trim())
                        .filter(e => e);

                    let loggedStay = false;
                    const loggedExpected = new Set();

                    for (const trimmedExp of expectedParts) {
                        const key = trimmedExp.toLowerCase();
                        if (loggedExpected.has(key)) continue;

                        try {
                            // --- Stay check ---
                            if (key === "stay" && !loggedStay) {
                                const currentUrlAfter = page.url();
                                const stayed = currentUrlAfter === currentUrlBefore;

                                if (stayed) {
                                    await takeScreenshot("page_url_stayed_same");
                                    record.Outcome += "✅ Stayed on same page. ";
                                } else {
                                    await takeScreenshot("ERROR_STAY_FAILED");
                                    record.Result = "Fail";
                                    record.Outcome += `❌ Expected to stay on same page, but URL changed from ${currentUrlBefore} to ${currentUrlAfter}. `;
                                    await incrementProgress(`Checking expected outcome: stay`, "Expected Outcome", new Error(`URL did not stay the same. Current: ${currentUrlAfter}`));
                                }

                                await incrementProgress(`Checking expected outcome: stay`, "Expected Outcome");
                                loggedStay = true;

                                // --- URL contains check ---
                            } else if (!trimmedExp.includes("=")) {
                                const cur = page.url();
                                const matched = cur.includes(trimmedExp);

                                if (matched) {
                                    await takeScreenshot("page_url_matched");
                                    record.Outcome += `✅ Expected URL matched: ${trimmedExp}. `;
                                } else {
                                    await takeScreenshot("ERROR_URL_MISMATCH");
                                    record.Result = "Fail";
                                    record.Outcome += `❌ Expected URL ${trimmedExp}, got ${cur}. `;
                                    await incrementProgress(`Checking expected outcome: URL contains '${trimmedExp}'`, "Expected Outcome", new Error(`URL mismatch. Expected: ${trimmedExp}, Found: ${cur}`));
                                }

                                await incrementProgress(`Checking expected outcome: URL contains '${trimmedExp}'`, "Expected Outcome");

                                // --- AppearText check ---
                            } else {
                                const [k, val] = trimmedExp.split("=").map(s => s.trim());
                                if (k.toLowerCase() === "appeartext") {
                                    const expectedText = val;
                                    const optionSelector = `text=${expectedText}`;
                                    const outlineColor = (row["outlinecolor"] || "red").trim() || "red";

                                    try {
                                        const element = await page.locator(optionSelector);
                                        await element.waitFor({ timeout: 5000, state: "visible" });
                                        const visibleText = await element.textContent();

                                        let fileName;

                                        if (visibleText === expectedText) {
                                            record.Outcome += `✅ AppearText found exactly: (${expectedText})`;
                                            fileName = `${testCaseName}_appeartext_${normalizeName(expectedText)}.png`;
                                        } else if (visibleText.includes(expectedText)) {
                                            record.Outcome += `✅ AppearText partially matched: (${expectedText}). Found: (${visibleText})`;
                                            fileName = `${testCaseName}_appeartext_${normalizeName(expectedText)}.png`;
                                        } else {
                                            record.Result = "Fail";
                                            record.Outcome += `❌ AppearText not met: ${expectedText}. Found: (${visibleText}). `;
                                            fileName = `ERROR_APPEARTEXT_NOTFOUND.png`;
                                            await incrementProgress(
                                                `Checking expected outcome: AppearText='${expectedText}'`,
                                                "Expected Outcome",
                                                new Error(`AppearText='${expectedText}' not found. Actual='${visibleText}'`)
                                            );
                                        }

                                        // --- Absolute path for screenshot
                                        const fullShotPath = path.resolve(screenshotsDir, fileName);

                                        // --- Take highlight screenshot(s)
                                        const { screenshots: highlightShots, outcomeMessage } = await highlightAndScreenshot(
                                            page,
                                            optionSelector,
                                            fullShotPath,
                                            outlineColor
                                        );

                                        // --- Store in global screenshots array
                                        screenshots.push(...highlightShots);

                                        // --- Store in record.Screenshot (multi-line)
                                        record.Screenshot = record.Screenshot
                                            ? `${record.Screenshot}\n${highlightShots.join("\n")}`
                                            : highlightShots.join("\n");

                                        // --- Append outcome messages
                                        record.Outcome += outcomeMessage;

                                        await incrementProgress(
                                            `Checking expected outcome: AppearText='${expectedText}'`,
                                            "Expected Outcome"
                                        );

                                    } catch (err) {
                                        record.Result = "Fail";
                                        record.Outcome += `❌ AppearText=${expectedText} not met. `;

                                        // --- Absolute path for error screenshot
                                        const fullShotPath = path.resolve(screenshotsDir, `ERROR_APPEARTEXT_CONFLICT.png`);

                                        const { screenshots: highlightShots, outcomeMessage } = await highlightAndScreenshot(
                                            page,
                                            optionSelector,
                                            fullShotPath,
                                            outlineColor
                                        );

                                        screenshots.push(...highlightShots);

                                        record.Screenshot = record.Screenshot
                                            ? `${record.Screenshot}\n${highlightShots.join("\n")}`
                                            : highlightShots.join("\n");

                                        record.Outcome += outcomeMessage;

                                        // --- Pass error to incrementProgress for user-friendly formatting
                                        await incrementProgress(
                                            `Checking expected outcome: AppearText='${expectedText}'`,
                                            "Expected Outcome",
                                            err
                                        );
                                    }

                                }
                            }
                        } catch (err) {
                            record.Result = "Fail";
                            record.Outcome += `❌ ${trimmedExp} not met. `;
                            await incrementProgress(`Checking expected outcome: ${trimmedExp}`, "Expected Outcome", err);
                        }

                        loggedExpected.add(key);
                    }

                    if (executedActions.length > 0) {
                        record.Outcome = `Run Action: ${executedActions.join(",")}${record.Outcome ? "\n" + record.Outcome : ""}`;
                    }
                }
                if (row.wait) await safeWait(row.wait);

            } catch (err) {
                record.Result = "Fail";
                record.Outcome += `❌ Error: ${err.message}`;
                console.warn(`❌ ${chalk.redBright(`TestCase ${testCaseName} failed: ${err.message}`)}`);
                runLog(`❌ TestCase ${testCaseName} failed: ${err.message}`);
                await killProgressBar();
            }

            const endTime = Date.now();
            record.Duration = (endTime - startTime) / 1000;

            // --- Add row in Excel ---
            ws.getColumn(1).width = 40;
            ws.getColumn(2).width = 75;
            ws.getColumn(3).width = 43;
            ws.getColumn(4).width = 54;
            ws.getColumn(5).width = 43;
            ws.getColumn(5).width = 20;
            ws.mergeCells("A1:B1");
            ws.getCell("A1").border = {
                top: { style: "thin", color: { argb: "FF000000" } },
                left: { style: "thin", color: { argb: "FF000000" } },
                bottom: { style: "thin", color: { argb: "FF000000" } },
                right: { style: "thin", color: { argb: "FF000000" } }
            };
            ws.getCell("A1").value = "Results";
            ws.getCell("A1").fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "FFBDD7EE" } // pale blue
            };

            ws.getCell("A1").alignment = { vertical: "middle", horizontal: "center" };
            ws.getCell("A1").font = { bold: true };

            ws.getCell("A2").value = "Test Case";
            ws.getCell("A2").font = { bold: true };
            ws.getCell("B2").value = testCaseName;

            ws.getCell("A3").value = "Outcome";
            ws.getCell("A3").font = { bold: true };
            ws.getCell("B3").value = record.Outcome;

            ws.getCell("A4").value = "Duration(s)";
            ws.getCell("A4").font = { bold: true };
            ws.getCell("B4").value = record.Duration.toFixed(3) + "s";

            ws.getCell("A5").value = "Ran Actions";
            ws.getCell("A5").font = { bold: true };
            ws.getCell("B5").value = executedActions.join(",");

            ws.getCell("A6").value = "DeviceType";
            ws.getCell("A6").font = { bold: true };
            ws.getCell("B6").value = deviceType || "Default(Window)";

            ws.getCell("A7").value = "Overall result";
            ws.getCell("A7").font = { bold: true };
            ws.getCell("B7").value = (`✔ +${record.Result}`) || "❌Fail";
            if (record.Result == "Pass") {
                ws.getCell("B7").value = (`✔ PASS`);
                ws.getCell("B7").font = {
                    bold: true,
                    size: 13,
                    color: { argb: "FF00FF00" } // green
                };
            } else if (record.Result == "Fail") {
                ws.getCell("B7").value = (`❌ FAIL`);
                ws.getCell("B7").font = {
                    bold: true,
                    size: 13,
                    color: { argb: "FFFF0000" } // red
                };
            }

            // Clear console messages for next test case
            consoleMessages.length = 0;

            let topOffset = 0;
            const startRow = rowIndex;

            ws.mergeCells(`A${startRow - 1}:C${startRow - 1}`);
            ws.getCell(`A${startRow - 1}`).value = "Screenshots";
            ws.getCell(`A${startRow - 1}`).alignment = { horizontal: "center" };
            ws.getCell(`A${startRow - 1}`).font = { bold: true };
            ws.getCell(`A${startRow - 1}`).fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "FFBDD7EE" }// pale blue
            };

            ws.getCell(`A${startRow}`).value = "Step";
            ws.getCell(`A${startRow}`).alignment = { horizontal: "center" };
            ws.getCell(`A${startRow}`).font = { bold: true };
            ws.getCell(`A${startRow}`).fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "FFBDD7EE" }// pale blue
            };

            ws.getCell(`B${startRow}`).value = "Image location";
            ws.getCell(`B${startRow}`).alignment = { horizontal: "center" };
            ws.getCell(`B${startRow}`).font = { bold: true };
            ws.getCell(`B${startRow}`).fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "FFBDD7EE" }//pale blue
            };

            ws.getCell(`C${startRow}`).value = "Note";
            ws.getCell(`C${startRow}`).alignment = { horizontal: "center" };
            ws.getCell(`C${startRow}`).font = { bold: true };
            ws.getCell(`C${startRow}`).fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "FFBDD7EE" }//pale blue
            };
            rowIndex++;

            // Xlsx Screenshot section
            // --- Set default column widths
            [1, 2, 3].forEach(i => {
                if (!ws.getColumn(i).width || ws.getColumn(i).width < 10) ws.getColumn(i).width = 15;
            });

            // --- Fixed container row height in Excel units
            const containerRowHeight = 200; // Excel row height (approx.)
            const rowHeightPixels = containerRowHeight * 1.33; // approximate visual pixels

            // Ensure columns have a minimum width
            [1, 2, 3].forEach(i => {
                if (!ws.getColumn(i).width || ws.getColumn(i).width < 10) ws.getColumn(i).width = 15;
            });

            // Helper: convert pixels to EMU for absolute positioning
            const pxToEMU = px => px * 9525;

            // Precompute total width of merged columns in pixels
            const colWidthToPixels = w => Math.floor(w * 7.1 * 8);
            const totalWidthPx =
                colWidthToPixels(ws.getColumn(1).width) +
                colWidthToPixels(ws.getColumn(2).width) +
                colWidthToPixels(ws.getColumn(3).width);

            for (const fullShotPath of screenshots) {
                if (!fs.existsSync(fullShotPath)) continue;

                const shotName = path.basename(fullShotPath, path.extname(fullShotPath));

                // --- Get image dimensions using sharp helper
                let dimensions;
                try {
                    dimensions = await getImageDimensions(fullShotPath);
                } catch (err) {
                    console.error("⚠️ Failed to read image size:", fullShotPath, err);
                    continue;
                }

                const imgW = dimensions.width;
                const imgH = dimensions.height;

                // --- Fill label row
                ws.getCell(`A${rowIndex}`).value = shotName;
                // --- Fill background color
                if (record.Result == "Pass") {
                    ws.getCell(`A${rowIndex}`).fill = {
                        type: "pattern",
                        pattern: "solid",
                        fgColor: { argb: "FFDFFFD6" } // pale green
                    };
                } else if (record.Result == "Fail") {
                    ws.getCell(`A${rowIndex}`).fill = {
                        type: "pattern",
                        pattern: "solid",
                        fgColor: { argb: "FFFFD6D6" } // pale red
                    };
                }
                // --- Image location as clickable hyperlink
                ws.getCell(`B${rowIndex}`).value = {
                    text: fullShotPath,       // the visible text
                    hyperlink: fullShotPath   // the actual clickable link
                };
                ws.getCell(`B${rowIndex}`).font = { color: { argb: "FF0000FF" }, underline: true };

                ws.getCell(`C${rowIndex}`).value = record.Outcome;

                // --- Prepare image row
                const imageId = ws.workbook.addImage({ filename: fullShotPath, extension: "png" });
                const mergedRow = rowIndex + 1;
                const tlRow = mergedRow - 1;

                ws.mergeCells(`A${mergedRow}:C${mergedRow}`);
                ws.getRow(mergedRow).height = containerRowHeight;

                // --- Scale image to fit row height visually while keeping aspect ratio
                const scale = rowHeightPixels / imgH;
                const excelImgHeight = Math.round(imgH * scale);
                const excelImgWidth = Math.round(imgW * scale);

                // --- Compute horizontal center offset in EMU
                const offsetX = pxToEMU(Math.floor((totalWidthPx - excelImgWidth) / 2));

                // --- Add image to Excel centered in merged cells
                ws.addImage(imageId, {
                    tl: { col: 0, row: tlRow, offsetX, offsetY: 0 },
                    ext: { width: excelImgWidth, height: excelImgHeight },
                    editAs: "absolute"
                });

                // --- Move to next screenshot
                rowIndex += 3;
            }

            // --- Network captures ---
            // Section header
            ws.mergeCells(`A${rowIndex}:E${rowIndex}`);
            ws.getCell(`A${rowIndex}`).value = "Network Log";
            ws.getCell(`A${rowIndex}`).fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "FFBDD7EE" }//pale blue
            };
            ws.getCell(`A${rowIndex}`).alignment = { horizontal: "center" };
            ws.getCell(`A${rowIndex}`).font = { bold: true };
            rowIndex++;

            // Column headers
            ws.getCell(`A${rowIndex}`).value = "ActionTypes";
            ws.getCell(`B${rowIndex}`).value = "Selector/Value";
            ws.getCell(`C${rowIndex}`).value = "Method";
            ws.getCell(`D${rowIndex}`).value = "URL";
            ws.getCell(`E${rowIndex}`).value = "NetworkStatus";
            rowIndex++;

            // Write current log entry only
            ws.getCell(`A${rowIndex}`).value = logEntry.ActionType;
            ws.getCell(`B${rowIndex}`).value = logEntry.SelectorOrValue;
            ws.getCell(`C${rowIndex}`).value = logEntry.Method;
            ws.getCell(`D${rowIndex}`).value = logEntry.URL;
            ws.getCell(`E${rowIndex}`).value = logEntry.NetworkStatus;
            rowIndex++;

            // --- Write console messages to Console Captures sheet ---
            // Section header
            ws.mergeCells(`A${rowIndex}:E${rowIndex}`);
            ws.getCell(`A${rowIndex}`).value = "Browser Console Log";
            ws.getCell(`A${rowIndex}`).fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "FFBDD7EE" }//pale blue
            };
            ws.getCell(`A${rowIndex}`).alignment = { horizontal: "center" };
            ws.getCell(`A${rowIndex}`).font = { bold: true };

            rowIndex++;

            // Column headers
            ws.getCell(`A${rowIndex}`).value = "ConsoleType";
            ws.mergeCells(`B${rowIndex}:E${rowIndex}`);
            ws.getCell(`B${rowIndex}`).value = "Message";
            rowIndex++;

            // Fill console messages
            for (const msg of consoleMessages) {
                ws.getCell(`A${rowIndex}`).value = msg.type;
                ws.mergeCells(`B${rowIndex}:E${rowIndex}`);
                ws.getCell(`B${rowIndex}`).value = msg.text;
                rowIndex++;
            }

            // Adjust row height
            ws.getRow(rowIndex - 1).height = topOffset; // adjust last used row

            // --- Convert screenshot paths to relative paths from HTML file ---
            const relativeScreenshots = screenshots.map(s =>
                path.relative(path.dirname(htmlReportPath), s).replace(/\\/g, "/")
            );

            testResults.push({
                testCase: testCaseName,
                result: record.Result,
                duration: record.Duration,
                imgWidth: deviceProfile.viewport.width,
                imgHeight: deviceProfile.viewport.height,
                outcome: record.Outcome,
                screenshots: relativeScreenshots,
                deviceNames: deviceType || "default",
                networkLog: networkLogs.filter(r => r.TestCase === testCaseName)
            });
            runLog(`========== 🏁 Completed TestCase: (${testCaseName}) ==========\n`);
            console.log(`========== 🏁 ${chalk.bold("Completed TestCase:")} (${chalk.cyan(testCaseName)}) ==========\n`);
        }


        await browser.close();
        console.log(chalk.bold("✅ Browser closed"));
        runLog("✅ Browser closed");

        const endTime = new Date();
        const totalRunTime = endTime - startTime;
        const totalSeconds = Math.floor(totalRunTime / 1000);
        const minutes = Math.floor(totalSeconds / 60);
        const seconds = totalSeconds % 60;

        console.log(`🐰 ${chalk.bold("Total run time:")} ${chalk.greenBright.bold(`${minutes}min ${seconds}s`)}`);
        runLog(`🐇 Total run time: ${minutes}min ${seconds}s`);

        // --- Save Excel results ---
        await resultsWorkbook.xlsx.writeFile(excelResultPath);
        console.log(`📄 ${chalk.bold("Excel results saved:")} ${chalk.underline(excelResultPath)}`);
        runLog(`📄 Excel results saved: ${excelResultPath}`);

        // --- Generate HTML report inside same folder ---
        generateHTMLReport(file, testResults, htmlReportPath, totalSeconds, browserName);

        fs.writeFileSync(progressFile, totalTests.toString(), "utf8");
    }
    // #endregion ---
} catch (err) {
    console.error(chalk.redBright(`⚠️ Unexpected Error occured:${err}`));
    runLog(`⚠️ Unexpected Error occured:${err}`);
    killProgressBar();
}
