import sharp from "sharp";
import fs from "fs";
import path from "path";

export async function getImageDimensions(fullShotPath) {
    if (!fs.existsSync(fullShotPath)) {
        console.warn("⚠️ Screenshot file missing:", fullShotPath);
        return { width: 0, height: 0 };
    }

    try {
        const metadata = await sharp(fullShotPath).metadata();
        return { width: metadata.width || 0, height: metadata.height || 0 };
    } catch (err) {
        console.error("⚠️ Failed to get image size via sharp:", fullShotPath, err);
        return { width: 0, height: 0 };
    }
}
