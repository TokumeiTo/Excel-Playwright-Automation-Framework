export function rewriteStepErrorForUser(stepDescription, error) {
    if (!error) return `Step Issue: ${stepDescription} -> Unknown error occurred.`;
    let message;

    // Ensure we have a string to work with
    const errMsg = typeof error === "string" ? error : error?.message || JSON.stringify(error);

    // Handle net::ERR_CONNECTION_REFUSED and similar
    if (/ERR_CONNECTION_REFUSED/i.test(errMsg)) {
        const urlMatch = errMsg.match(/(http[s]?:\/\/[^\s"]+)/i);
        const foundURL = urlMatch ? urlMatch[1] : "the target URL";

        return `Failed to reach "${foundURL}"\n💡 NOTE FIX: Please check your link and ensure the server is running or your network connection is stable.`;
    }

    // Special handling for strict mode violation
    if (/strict mode violation/.test(errMsg) && /locator\('text=/.test(errMsg)) {
        const match = errMsg.match(/locator\('text=(.*?)'\)/);
        const expectedText = match ? match[1] : "the expected text";

        // Extract elements info
        const elementMatches = [...errMsg.matchAll(/(\d\)) (.*?) aka (.*?)\n/g)];
        const elementsInfo = elementMatches
            .map((m, i) => ` ${m[1]} ${m[2]} ${i === 0 ? "✔ Accurate" : ""}`)
            .join("\n");

        return `Expected text: ('${expectedText}') was found in multiple elements.\n Error: Strict-mode violation of selector('text=${expectedText}'). Found ${elementMatches.length} similar elements:\n${elementsInfo}\n💡 NOTE FIX: Please reconsider your text.`;
    } else {
        // fallback
        message = errMsg.replace(/\s+/g, " ").trim();
    }

    return message;
}
