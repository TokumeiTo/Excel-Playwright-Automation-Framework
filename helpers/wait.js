export async function waitForPaintToSettle(page) {
    try {
        await page.evaluate(() =>
            new Promise(resolve => {
                requestAnimationFrame(() => {
                    requestAnimationFrame(resolve);
                });
            })
        );
    } catch (err) {
        await page.waitForTimeout(30);
    }
}
