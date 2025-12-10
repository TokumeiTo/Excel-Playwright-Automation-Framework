import { devices } from 'playwright';
import { runLog } from './runLog.js';

const deviceMap = {
    // BlackBerry profiles
    'blackberry z30': devices['BlackBerry Z30'],
    'blackberry z30 landscape': devices['BlackBerry Z30 landscape'],
    'blackberry playbook': devices['BlackBerry PlayBook'],
    'blackberry playbook landscape': devices['BlackBerry PlayBook landscape'],

    // Samsung profiles
    'galaxy a55': devices['Galaxy A55'],
    'galaxy a55 landscape': devices['Galaxy A55 landscape'],
    'galaxy note 3': devices['Galaxy Note 3'],
    'galaxy note 3 landscape': devices['Galaxy Note 3 landscape'],
    'galaxy note 2': devices['Galaxy Note II'],
    'galaxy note 2 landscape': devices['Galaxy Note II landscape'],
    'galaxy s3': devices['Galaxy S III'],
    'galaxy s3 landscape': devices['Galaxy S III landscape'],
    'galaxy s24': devices['Galaxy S24'],
    'galaxy s24 landscape': devices['Galaxy S24 landscape'],
    'galaxy s5': devices['Galaxy S5'],
    'galaxy s5 landscape': devices['Galaxy S5 landscape'],
    'galaxy s8': devices['Galaxy S8'],
    'galaxy s8 landscape': devices['Galaxy S8 landscape'],
    'galaxy s9+': devices['Galaxy S9+'],
    'galaxy s9+ landscape': devices['Galaxy S9+ landscape'],
    'galaxy tab s4': devices['Galaxy Tab S4'],
    'galaxy tab s4 landscape': devices['Galaxy Tab S4 landscape'],

    // Kindle Fire profiles
    'kindle fire hdx': devices['Kindle Fire HDX'],
    'kindle fire hdx landscape': devices['Kindle Fire HDX landscape'],

    // LG profiles
    'lg optimus l70': devices['LG Optimus L70'],
    'lg optimus l70 landscape': devices['LG Optimus L70 landscape'],

    // MsLumia profiles
    'microsoft lumia 550': devices['Microsoft Lumia 550'],
    'microsoft lumia 550 landscape': devices['Microsoft Lumia 550 landscape'],
    'microsoft lumia 950': devices['Microsoft Lumia 950'],
    'microsoft lumia 950 landscape': devices['Microsoft Lumia 950 landscape'],

    // Nokia profiles
    'nokia lumia 520': devices['Nokia Lumia 520'],
    'nokia lumia 520 landscape': devices['Nokia Lumia 520 landscape'],
    'nokia n9': devices['Nokia N9'],
    'nokia n9 landscape': devices['Nokia N9 landscape'],

    // Moto profiles
    'moto g4': devices['Moto G4'],
    'moto g4 landscape': devices['Moto G4 landscape'],

    // Nexus profiles
    'nexus 4': devices['Nexus 4'],
    'nexus 4 landscape': devices['Nexus 4 landscape'],
    'nexus 5': devices['Nexus 5'],
    'nexus 5 landscape': devices['Nexus 5 landscape'],
    'nexus 5x': devices['Nexus 5X'],
    'nexus 5x landscape': devices['Nexus 5X landscape'],
    'nexus 6': devices['Nexus 6'],
    'nexus 6 landscape': devices['Nexus 6 landscape'],
    'nexus 6p': devices['Nexus 6P'],
    'nexus 6p landscape': devices['Nexus 6P landscape'],
    'nexus 7': devices['Nexus 7'],
    'nexus 7 landscape': devices['Nexus 7 landscape'],
    'nexus 10': devices['Nexus 10'],
    'nexus 10 landscape': devices['Nexus 10 landscape'],

    // Desktop profiles
    'desktop chrome': devices['Desktop Chrome'],
    'desktop chrome hidpi': devices['Desktop Chrome HiDPI'],
    'desktop firefox': devices['Desktop Firefox'],
    'desktop firefox hidpi': devices['Desktop Firefox HiDPI'],
    'desktop safari': devices['Desktop Safari'],
    'desktop safari hidpi': devices['Desktop Safari HiDPI'],
    'desktop edge': devices['Desktop Edge'],
    'desktop edge hidpi': devices['Desktop Edge HiDPI'],


    // Android profiles
    'pixel 2': devices['Pixel 2'],
    'pixel 2 landscape': devices['Pixel 2 landscape'],
    'pixcel 2 xl': devices['Pixel 2 XL'],
    'pixel 2 xl landscape': devices['Pixel 2 XL landscape'],
    'pixel 3': devices['Pixel 3'],
    'pixel 3 landscape': devices['Pixel 3 landscape'],
    'pixel 4': devices['Pixel 4'],
    'pixel 4 landscape': devices['Pixel 4 landscape'],
    'pixel 4a 5g': devices['Pixel 4a (5G)'],
    'pixel 4a 5g landscape': devices['Pixel 4a (5G) landscape'],
    'pixel 5': devices['Pixel 5'],
    'pixel 5 landscape': devices['Pixel 5 landscape'],
    'pixel 7': devices['Pixel 7'],
    'pixel 7 landscape': devices['Pixel 7 landscape'],

    // iPad profiles
    'ipad pro': {
        userAgent: "Mozilla/5.0 (iPad; CPU iPad OS 14_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0 Mobile/15E148 Safari/604.1",
        viewport: { width: 768, height: 1024 },
        deviceScaleFactor: 2,
        isMobile: true,
        hasTouch: true
    },
    'ipad pro landscape': {
        userAgent: "Mozilla/5.0 (iPad; CPU iPad OS 14_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0 Mobile/15E148 Safari/604.1",
        viewport: { width: 1024, height: 768 },
        deviceScaleFactor: 2,
        isMobile: true,
        hasTouch: true
    },
    'ipad (gen 11)': devices['iPad (gen 11)'],
    'ipad (gen 11) landscape': devices['iPad (gen 11) landscape'],
    'ipad (gen 5)': devices['iPad (gen 5)'],
    'ipad (gen 5) landscape': devices['iPad (gen 5) landscape'],
    'ipad (gen 6)': devices['iPad (gen 6)'],
    'ipad (gen 6) landscape': devices['iPad (gen 6) landscape'],
    'ipad (gen 7)': devices['iPad (gen 7)'],
    'ipad (gen 7) landscape': devices['iPad (gen 7) landscape'],
    'ipad mini': devices['iPad Mini'],
    'ipad mini landscape': devices['iPad Mini landscape'],
    'ipad pro 11': devices['iPad Pro 11'],
    'ipad pro 11 landscape': devices['iPad Pro 11 landscape'],

    // iPhone profiles
    'iphone 6': devices['iPhone 6'],
    'iphone 6 landscape': devices['iPhone 6 landscape'],
    'iphone 6 plus': devices['iPhone 6 Plus'],
    'iphone 6 plus landscape': devices['iPhone 6 Plus landscape'],
    'iphone 7': devices['iPhone 7'],
    'iphone 7 landscape': devices['iPhone 7 landscape'],
    'iphone 7 plus': devices['iPhone 7 Plus'],
    'iphone 7 plus landscape': devices['iPhone 7 Plus landscape'],
    'iphone 8': devices['iPhone 8'],
    'iphone 8 landscape': devices['iPhone 8 landscape'],
    'iphone 8 plus': devices['iPhone 8 Plus'],
    'iphone 8 plus landscape': devices['iPhone 8 Plus landscape'],
    'iphone se': devices['iPhone SE'],
    'iphone se landscape': devices['iPhone SE landscape'],
    'iphone se 3rd gen': devices['iPhone SE (3rd generation)'],
    'iphone se 3rd gen landscape': devices['iPhone SE (3rd generation) landscape'],
    'iphone x': devices['iPhone X'],
    'iphone x landscape': devices['iPhone X landscape'],
    'iphone xr': devices['iPhone XR'],
    'iphone xr landscape': devices['iPhone XR landscape'],
    'iphone 11': devices['iPhone 11'],
    'iphone 11 landscape': devices['iPhone 11 landscape'],
    'iphone 11 pro': devices['iPhone 11 Pro'],
    'iphone 11 pro landscape': devices['iPhone 11 Pro landscape'],
    'iphone 11 pro max': devices['iPhone 11 Pro Max'],
    'iphone 11 pro max landscape': devices['iPhone 11 Pro Max landscape'],
    'iphone 12': devices['iPhone 12'],
    'iphone 12 landscape': devices['iPhone 12 landscape'],
    'iphone 12 pro': devices['iPhone 12 Pro'],
    'iphone 12 pro landscape': devices['iPhone 12 Pro landscape'],
    'iphone 12 pro max': devices['iPhone 12 Pro Max'],
    'iphone 12 pro max landscape': devices['iPhone 12 Pro Max landscape'],
    'iphone 12 mini': devices['iPhone 12 Mini'],
    'iphone 12 mini landscape': devices['iPhone 12 Mini landscape'],
    'iphone 13': devices['iPhone 13'],
    'iphone 13 landscape': devices['iPhone 13 landscape'],
    'iphone 13 pro': devices['iPhone 13 Pro'],
    'iphone 13 pro landscape': devices['iPhone 13 Pro landscape'],
    'iphone 13 pro max': devices['iPhone 13 Pro Max'],
    'iphone 13 pro max landscape': devices['iPhone 13 Pro Max landscape'],
    'iphone 13 mini': devices['iPhone 13 Mini'],
    'iphone 13 mini landscape': devices['iPhone 13 Mini landscape'],
    'iphone 14': devices['iPhone 14'],
    'iphone 14 landscape': devices['iPhone 14 landscape'],
    'iphone 14 pro': devices['iPhone 14 Pro'],
    'iphone 14 pro landscape': devices['iPhone 14 Pro landscape'],
    'iphone 14 pro max': devices['iPhone 14 Pro Max'],
    'iphone 14 pro max landscape': devices['iPhone 14 Pro Max landscape'],
    'iphone 14 plus': devices['iPhone 14 Plus'],
    'iphone 14 plus landscape': devices['iPhone 14 Plus landscape'],
    'iphone 15': devices['iPhone 15'],
    'iphone 15 landscape': devices['iPhone 15 landscape'],
    'iphone 15 pro': devices['iPhone 15 Pro'],
    'iphone 15 pro landscape': devices['iPhone 15 Pro landscape'],
    'iphone 15 pro max': devices['iPhone 15 Pro Max'],
    'iphone 15 pro max landscape': devices['iPhone 15 Pro Max landscape'],
    'iphone 15 plus': devices['iPhone 15 Plus'],
    'iphone 15 plus landscape': devices['iPhone 15 Plus landscape'],

    // Add more mappings here as needed
};

function formatDeviceProfile(profile) {
    const { userAgent, viewport, deviceScaleFactor, isMobile, hasTouch } = profile;
    return JSON.stringify({ userAgent, viewport, deviceScaleFactor, isMobile, hasTouch });
}

// Device profile lookup function
export function getDeviceProfile(deviceType) {
    const normalizedDeviceType = (deviceType || "").trim().toLowerCase();

    let deviceProfile;

    if (!normalizedDeviceType) {
        runLog("📵 No device type specified, defaulting to 'Desktop Chrome'");
        deviceProfile = devices['Desktop Chrome'];
    } else {
        deviceProfile = deviceMap[normalizedDeviceType] || devices['Desktop Chrome'];
        if (!deviceMap[normalizedDeviceType]) {
            runLog(`📵 Unknown device type '${deviceType}', defaulting to 'Desktop Chrome'`);
        }
    }

    // Log the resolved device profile once
    runLog(`📃 Resolved device profile: ${formatDeviceProfile(deviceProfile)}`);

    return deviceProfile;
}
