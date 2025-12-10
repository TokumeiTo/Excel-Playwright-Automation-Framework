import { runLog } from "./runLog.js";
import fs from "fs";
import path, { resolve } from "path";
import { waitForPaintToSettle } from "./wait.js";
import chalk from "chalk";
import { spinnerTick } from "./consoleLoading.js";

export async function highlightAndScreenshot(page, selectorStr, screenshotPath, outlineColor = "red") {
    const results = [];
    let outcomeMessage = "";

    try {
        // split multiple selectors by "&"
        const selectors = selectorStr
            .split("&")
            .map(s => s.trim())
            .filter(Boolean);

        if (selectors.length === 0) {
            outcomeMessage = "❌ No valid selector/s provided.";
            console.warn(chalk.yellowBright("⚠️ No valid selector/s provided"));
            runLog("⚠️ No valid selector/s provided")
            return { screenshots: results, outcomeMessage };
        }

        const highlightedEls = [];

        for (const selector of selectors) {
            try {
                await page.waitForSelector(selector, { timeout: 5000 });

                const el = await page.$(selector);
                if (!el) {
                    console.warn(chalk.red(`❌ Element not found for selector: ${selector}`));
                    runLog(`❌ Element not found for selector: ${selector}`);
                    continue;
                }

                // ✅ Check if visible
                const isVisible = await el.evaluate(e => {
                    const rect = e.getBoundingClientRect();
                    return rect.width > 0 && rect.height > 0;
                });
                if (!isVisible) console.warn(chalk.yellowBright("⚠️ Warning: Element not visible or detached."));

                const box = await el.boundingBox();
                if (!box) console.warn(chalk.yellowBright("⚠️ Warning: Element not visible or detached."));

                // ✅ Universal blocking detection
                const blockingInfo = await page.evaluate(({ x, y, width, height }) => {
                    const centerX = x + width / 2;
                    const centerY = y + height / 2;
                    const topElement = document.elementFromPoint(centerX, centerY);
                    if (!topElement) return null;

                    const style = window.getComputedStyle(topElement);
                    const tag = topElement.tagName;
                    const cls = topElement.className || "";
                    const attrs = topElement.getAttributeNames().join(" ");
                    const topZ = style.zIndex;

                    const looksLikePopup =
                        topElement.closest('[data-sonner-toast], [role="alert"], [role="status"], [data-type="error"], [data-type="success"], [aria-live], .toast, .snackbar, .alert, .popup, .modal, .dialog') !== null ||
                        (["fixed", "absolute"].includes(style.position) && Number(topZ) > 100);

                    return { topTag: tag, topClass: cls, topZ, looksLikePopup };
                }, box);

                let waitElapsed = 0;
                const maxWait = 10000;
                const waitStep = 200;

                if (blockingInfo?.looksLikePopup) {
                    runLog("👻 PopUp/s Appeared!");
                    while (waitElapsed < maxWait) {
                        const stillBlocked = await page.evaluate(() => {
                            const popup = document.querySelector(
                                '[data-sonner-toast][data-visible="true"], [role="alert"], [role="status"], .toast, .snackbar, .alert, .popup, .modal, .dialog'
                            );

                            if (!popup) return false;

                            const style = window.getComputedStyle(popup);
                            const isVisible =
                                style.display !== "none" &&
                                style.visibility !== "hidden" &&
                                popup.getAttribute("data-removed") !== "true" &&
                                popup.offsetParent !== null;

                            const z = Number(style.zIndex) || 0;
                            const isForeground = z > 100 || style.position === "fixed" || style.position === "absolute";

                            return isVisible && isForeground;
                        });

                        if (!stillBlocked) break;

                        spinnerTick("Waiting for popup/toast/modal to disappear...");
                        await page.waitForTimeout(waitStep);
                        waitElapsed += waitStep;
                    }

                    if (waitElapsed >= maxWait) {
                        outcomeMessage = `⚠️ Popup/s (z-index=${blockingInfo.topZ}) blocked element for > ${maxWait}ms. Screenshot may include overlay.`;
                        console.warn(chalk.yellow(outcomeMessage));
                        runLog(outcomeMessage);
                    } else {
                        outcomeMessage = `🍃 Popup/s disappeared after ${(waitElapsed / 1000).toFixed(2)}s.`;
                        console.log(outcomeMessage);
                        runLog(outcomeMessage);
                    }
                }

                // ✅ Highlight element
                const originalStyle = await el.evaluate(e => e.getAttribute("style") || "");
                highlightedEls.push({ el, originalStyle });

                await el.evaluate((e, color) => {
                    e.style.outline = `4px solid ${color}`;
                    e.style.outlineOffset = "5px";
                    e.style.boxShadow = `0 0 10px 3px ${color}`;
                }, outlineColor);
            } catch (innerErr) {
                console.warn(chalk.yellow(`⚠️ Failed to highlight selector ${selector}: ${innerErr.message}`));
                runLog(`⚠️ Failed to highlight selector ${selector}: ${innerErr.message}`);
            }
        }

        if (highlightedEls.length === 0) {
            outcomeMessage = "❌ No elements found to highlight.";
            console.warn(chalk.yellowBright("⚠️ No elements found to highlight"));
            runLog("⚠️ No elements found to highlight");
            return { screenshots: results, outcomeMessage };
        }

        // ✅ Take single screenshot of all highlighted elements
        if (screenshotPath) {
            fs.mkdirSync(path.dirname(screenshotPath), { recursive: true });
            await waitForPaintToSettle(page);
            await page.screenshot({ path: screenshotPath, fullPage: true });
            results.push(screenshotPath);
        }

        // ✅ Restore all element styles
        for (const { el, originalStyle } of highlightedEls) {
            try {
                await el.evaluate((e, style) => e.setAttribute("style", style), originalStyle);
            } catch (err) {
                outcomeMessage = `❌ Unexpected error occuredat restoring element style after highlight Err: ${err.message}`;
                console.warn(chalk.yellow(`⚠️ Unexpected error occured at Restoring element styles!! Err:\n ${err.message}`));
            }
        }
        outcomeMessage = `✅ Highlighted ${highlightedEls.length} element(s) successfully.`;

    } catch (err) {
        outcomeMessage = `❌ Highlight failed: ${err.message}`;
        console.warn(chalk.yellow(`⚠️ Highlight failed: ${err.message}`));
        runLog(`⚠️ Highlight failed: ${err.message}`);
    }

    return { screenshots: results, outcomeMessage };
}
