const express = require('express');
const { google } = require('googleapis');
const cors = require('cors');
const path = require('path');
const fs = require('fs');

// Try to use puppeteer-extra with stealth plugin, fallback to regular puppeteer
let puppeteer;
try {
    const puppeteerExtra = require('puppeteer-extra');
    const StealthPlugin = require('puppeteer-extra-plugin-stealth');
    puppeteerExtra.use(StealthPlugin());
    puppeteer = puppeteerExtra;
    console.log('✅ Using puppeteer-extra with stealth plugin');
} catch (e) {
    puppeteer = require('puppeteer');
    console.log('⚠️  Using regular puppeteer (puppeteer-extra not available)');
}

const app = express();
app.use(cors());
app.use(express.json());

// Serve Index.html as the main page
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'Index.html'));
});

// Serve static files from /public directory (for any assets)
app.use(express.static(path.join(__dirname, 'public')));

const PORT = 3000;

// Google Sheets API setup
// You can use either OAuth2 or Service Account
// For simplicity, we'll use OAuth2 with a client ID/secret

// Option 1: OAuth2 Client (requires client ID and secret)
// Option 2: Service Account (requires service account JSON file)
// For now, we'll use a simple approach with OAuth2

let oauth2Client = null;

// Initialize OAuth2 client
function initOAuth2() {
    // You'll need to set these in environment variables or a config file
    const CLIENT_ID = process.env.GOOGLE_CLIENT_ID || '';
    const CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET || '';
    const REDIRECT_URI = `http://localhost:${PORT}/oauth2callback`;

    if (!CLIENT_ID || !CLIENT_SECRET) {
        console.log('⚠️  Google OAuth credentials not set. Using manual token method.');
        return null;
    }

    oauth2Client = new google.auth.OAuth2(
        CLIENT_ID,
        CLIENT_SECRET,
        REDIRECT_URI
    );

    return oauth2Client;
}

// Service Account approach (easier, no user interaction needed)
let serviceAccountAuth = null;

function initServiceAccount() {
    try {
        // Try to load service account credentials
        const serviceAccountPath = path.join(__dirname, 'service-account.json');
        const fs = require('fs');
        
        if (fs.existsSync(serviceAccountPath)) {
            const credentials = JSON.parse(fs.readFileSync(serviceAccountPath, 'utf8'));
            serviceAccountAuth = new google.auth.JWT(
                credentials.client_email,
                null,
                credentials.private_key,
                ['https://www.googleapis.com/auth/spreadsheets']
            );
            console.log('✅ Service Account initialized');
            return serviceAccountAuth;
        }
    } catch (error) {
        console.log('⚠️  Service Account not found, using OAuth2');
    }
    return null;
}

// Try to initialize authentication
try {
    const serviceAccount = initServiceAccount();
    if (!serviceAccount) {
        initOAuth2();
    }
} catch (error) {
    console.error('Error initializing authentication:', error);
    // Try OAuth2 as fallback
    try {
        initOAuth2();
    } catch (oauthError) {
        console.error('Error initializing OAuth2:', oauthError);
    }
}

// API endpoint to highlight cells
app.post('/api/highlight-cells', async (req, res) => {
    try {
        console.log('Received request:', JSON.stringify(req.body, null, 2));
        
        const { spreadsheetId, sheetId, cells } = req.body;

        console.log('Extracted:', { 
            spreadsheetId: spreadsheetId ? 'present' : 'missing',
            sheetId: sheetId !== undefined && sheetId !== null ? `present (${sheetId})` : 'missing',
            cells: cells ? `present (${cells.length} items)` : 'missing'
        });
        
        // Log first few cells to see all flags
        if (cells && cells.length > 0) {
            console.log('First 5 cells sample:', cells.slice(0, 5).map(c => ({
                row: c.rowIndex,
                col: c.colIndex,
                isValid: c.isValid,
                isValidType: typeof c.isValid,
                isHttps: c.isHttps,
                isHttp: c.isHttp
            })));
        }

        if (!spreadsheetId || typeof spreadsheetId !== 'string') {
            return res.status(400).json({ error: 'Missing or invalid required parameter: spreadsheetId' });
        }
        if (sheetId === undefined || sheetId === null || (typeof sheetId !== 'number' && typeof sheetId !== 'string')) {
            return res.status(400).json({ error: 'Missing or invalid required parameter: sheetId (must be a number)' });
        }
        if (!cells || !Array.isArray(cells) || cells.length === 0) {
            return res.status(400).json({ error: 'Missing or empty required parameter: cells (must be a non-empty array)' });
        }

        let auth = serviceAccountAuth || oauth2Client;

        if (!auth) {
            return res.status(500).json({ 
                error: 'Authentication not configured. Please set up Google credentials.' 
            });
        }

        // Ensure we have valid credentials
        if (oauth2Client && !oauth2Client.credentials) {
            return res.status(401).json({ 
                error: 'Not authenticated. Please authorize first.',
                authUrl: `/auth`
            });
        }

        const sheets = google.sheets({ version: 'v4', auth });

        // Prepare highlight requests
        // Support both old format (just cells) and new format (cells with color info)
        const requests = cells.map((cell, index) => {
            // Default to yellow if no color specified
            let color = { red: 1, green: 1, blue: 0 };
            
            // Log raw values received
            const rawIsHttps = cell.isHttps;
            const rawIsHttp = cell.isHttp;
            const isHttpsType = typeof rawIsHttps;
            const isHttpType = typeof rawIsHttp;
            
            // Check for HTTP/HTTPS verification colors - be very explicit
            let isHttps = false;
            let isHttp = false;
            
            if (rawIsHttps === true || rawIsHttps === 'true' || rawIsHttps === 1) {
                isHttps = true;
            }
            if (rawIsHttp === true || rawIsHttp === 'true' || rawIsHttp === 1) {
                isHttp = true;
            }
            
            // Check for model names validation (isValid flag) - highest priority
            const rawIsValid = cell.isValid;
            const hasIsValidProperty = 'isValid' in cell;
            const isValidType = typeof rawIsValid;
            
            // Determine color - priority: isValid > isHttps > isHttp
            // Check isValid first (for model names validation)
            if (hasIsValidProperty && rawIsValid !== undefined && rawIsValid !== null) {
                // Model names validation: valid = green, invalid = red
                const isValidBool = rawIsValid === true || rawIsValid === 'true' || rawIsValid === 1;
                if (isValidBool) {
                    color = { red: 0, green: 1, blue: 0 }; // Green for valid
                } else {
                    color = { red: 1, green: 0, blue: 0 }; // Red for invalid
                }
            } else if (isHttps) {
                // HTTPS - Green
                color = { red: 0, green: 1, blue: 0 };
            } else if (isHttp) {
                // HTTP - Red
                color = { red: 1, green: 0, blue: 0 };
            } else if (cell.color) {
                // If cell has explicit color property, use it
                color = cell.color;
            }
            
            // Log first 5 cells for debugging
            if (index < 5) {
                const isValidBool = hasIsValidProperty && rawIsValid !== undefined && rawIsValid !== null 
                    ? (rawIsValid === true || rawIsValid === 'true' || rawIsValid === 1)
                    : null;
                console.log(`Cell ${index} color decision:`, {
                    rowIndex: cell.rowIndex,
                    colIndex: cell.colIndex,
                    hasIsValidProperty: hasIsValidProperty,
                    rawIsValid: rawIsValid,
                    isValidType: isValidType,
                    isValidBool: isValidBool,
                    rawIsHttps: rawIsHttps,
                    rawIsHttp: rawIsHttp,
                    isHttps: isHttps,
                    isHttp: isHttp,
                    finalColor: color,
                    colorRed: color.red,
                    colorGreen: color.green,
                    colorBlue: color.blue
                });
            }
            
            return {
                repeatCell: {
                    range: {
                        sheetId: parseInt(sheetId),
                        startRowIndex: cell.rowIndex,
                        endRowIndex: cell.rowIndex + 1,
                        startColumnIndex: cell.colIndex,
                        endColumnIndex: cell.colIndex + 1
                    },
                    cell: {
                        userEnteredFormat: {
                            backgroundColor: color
                        }
                    },
                    fields: 'userEnteredFormat.backgroundColor'
                }
            };
        });

        // Apply highlights
        const response = await sheets.spreadsheets.batchUpdate({
            spreadsheetId: spreadsheetId,
            resource: { requests }
        });

        res.json({ 
            success: true, 
            message: `Successfully highlighted ${cells.length} cells`,
            updatedCells: cells.length
        });

    } catch (error) {
        console.error('Error highlighting cells:', error);
        res.status(500).json({ 
            error: error.message || 'Failed to highlight cells',
            details: error.response?.data || error
        });
    }
});

// Helper function to find system Chrome/Chromium executable
function findSystemChrome() {
    const possiblePaths = [];
    
    if (process.platform === 'darwin') {
        // macOS Chrome locations
        possiblePaths.push(
            '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
            '/Applications/Chromium.app/Contents/MacOS/Chromium',
            '/Applications/Google Chrome Canary.app/Contents/MacOS/Google Chrome Canary',
            process.env.HOME + '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
            process.env.HOME + '/Applications/Chromium.app/Contents/MacOS/Chromium'
        );
    } else if (process.platform === 'linux') {
        // Linux Chrome locations
        possiblePaths.push(
            '/usr/bin/google-chrome',
            '/usr/bin/google-chrome-stable',
            '/usr/bin/chromium',
            '/usr/bin/chromium-browser',
            '/snap/bin/chromium'
        );
    } else if (process.platform === 'win32') {
        // Windows Chrome locations
        const programFiles = process.env['ProgramFiles'] || 'C:\\Program Files';
        const programFilesX86 = process.env['ProgramFiles(x86)'] || 'C:\\Program Files (x86)';
        const localAppData = process.env.LOCALAPPDATA || process.env.USERPROFILE + '\\AppData\\Local';
        
        possiblePaths.push(
            `${programFiles}\\Google\\Chrome\\Application\\chrome.exe`,
            `${programFilesX86}\\Google\\Chrome\\Application\\chrome.exe`,
            `${localAppData}\\Google\\Chrome\\Application\\chrome.exe`,
            `${programFiles}\\Chromium\\Application\\chromium.exe`,
            `${localAppData}\\Chromium\\Application\\chromium.exe`
        );
    }
    
    // Check which paths exist
    for (const chromePath of possiblePaths) {
        try {
            if (fs.existsSync(chromePath)) {
                console.log(`✅ Found system Chrome at: ${chromePath}`);
                return chromePath;
            }
        } catch (e) {
            // Continue checking other paths
        }
    }
    
    return null;
}

// Serve static files from /public directory
app.use(express.static(path.join(__dirname, 'public')));

// API endpoint to count jvxBase_* iframes and frm1_HL_ elements
app.get('/api/count', async (req, res) => {
    const url = req.query.url;
    if (!url) return res.status(400).json({ ok: false, error: 'Missing ?url=' });

    let browser;
    let page;
    const maxRetries = 3;
    let lastError = null;

    try {
        // Try to launch browser with retries and fallback options
        const launchOptions = [
            // Option 1: New headless mode with minimal args
            {
                headless: 'new',
                args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage']
            },
            // Option 2: Old headless mode
            {
                headless: true,
                args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage']
            },
            // Option 3: Try with even fewer args
            {
                headless: true,
                args: ['--no-sandbox']
            },
            // Option 4: Absolute minimum
            {
                headless: true,
                args: []
            }
        ];

        // Try to find system Chrome as fallback
        const systemChromePath = findSystemChrome();
        if (systemChromePath) {
            // Add system Chrome options
            launchOptions.push(
                {
                    executablePath: systemChromePath,
                    headless: 'new',
                    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage']
                },
                {
                    executablePath: systemChromePath,
                    headless: true,
                    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage']
                }
            );
        }

        let browserLaunched = false;
        for (let attempt = 1; attempt <= maxRetries; attempt++) {
            for (let optIndex = 0; optIndex < launchOptions.length; optIndex++) {
                try {
                    const isSystemChrome = launchOptions[optIndex].executablePath ? ' (system Chrome)' : '';
                    console.log(`Browser launch attempt ${attempt}, option ${optIndex + 1}${isSystemChrome}...`);
                    browser = await puppeteer.launch(launchOptions[optIndex]);
                    console.log(`✅ Browser launched successfully with option ${optIndex + 1}${isSystemChrome}`);
                    browserLaunched = true;
                    break;
                } catch (launchError) {
                    lastError = launchError;
                    if (optIndex === launchOptions.length - 1 && attempt === maxRetries) {
                        // Last option of last attempt
                        const errorMsg = launchError.message || String(launchError);
                        let helpfulMessage = `Failed to launch browser after ${maxRetries} attempts with ${launchOptions.length} different configurations.\n\n`;
                        
                        if (process.platform === 'darwin') {
                            helpfulMessage += `On macOS, this is often caused by:\n`;
                            helpfulMessage += `1. Missing or corrupted Chromium installation\n`;
                            helpfulMessage += `2. macOS security restrictions\n\n`;
                            helpfulMessage += `Solutions:\n`;
                            helpfulMessage += `- Reinstall Puppeteer: npm uninstall puppeteer && npm install puppeteer\n`;
                            helpfulMessage += `- Or install Chromium via Homebrew: brew install chromium --no-quarantine\n`;
                            helpfulMessage += `- Or install Google Chrome from https://www.google.com/chrome/\n\n`;
                        }
                        
                        helpfulMessage += `Original error: ${errorMsg.substring(0, 500)}`;
                        throw new Error(helpfulMessage);
                    }
                }
            }
            if (browserLaunched) break;
            
            if (attempt < maxRetries) {
                await new Promise(r => setTimeout(r, 1000 * attempt));
            }
        }
        
        if (!browserLaunched) {
            throw new Error(`Failed to launch browser: ${lastError?.message || 'Unknown error'}`);
        }

        page = await browser.newPage();
        
        // Set longer timeout
        page.setDefaultNavigationTimeout(90000);
        page.setDefaultTimeout(90000);

        // Try navigation with retries
        let navigationSuccess = false;
        for (let attempt = 1; attempt <= maxRetries; attempt++) {
            try {
                console.log(`Navigation attempt ${attempt}/${maxRetries} for ${url}`);
                
                // Wait for network idle as specified
                await page.goto(url, { 
                    waitUntil: 'networkidle0', 
                    timeout: 90000 
                });
                
                navigationSuccess = true;
                console.log('Navigation successful (network idle)');
                break;
            } catch (navError) {
                lastError = navError;
                const errorMsg = navError.message || String(navError);
                console.log(`Navigation attempt ${attempt} failed:`, errorMsg);
                
                // If it's a connection error, retry
                if (errorMsg.includes('ECONNRESET') || 
                    errorMsg.includes('ECONNREFUSED') ||
                    errorMsg.includes('ETIMEDOUT') ||
                    errorMsg.includes('net::ERR')) {
                    if (attempt < maxRetries) {
                        console.log(`Connection error, retrying in ${attempt * 2} seconds...`);
                        await new Promise(r => setTimeout(r, attempt * 2000));
                        continue;
                    }
                }
                
                // For other errors, still retry once
                if (attempt < maxRetries) {
                    await new Promise(r => setTimeout(r, 2000));
                    continue;
                }
            }
        }

        if (!navigationSuccess) {
            throw new Error(`Failed to navigate to URL after ${maxRetries} attempts: ${lastError?.message || 'Unknown error'}`);
        }

        // Wait for creative JS injection - 6 seconds fixed delay as specified
        console.log('Waiting 6 seconds for Jivox JS creative injection...');
        await new Promise(r => setTimeout(r, 6000));

        // ----------------------------
        // 0) Find previewVariationTitle elements to identify creatives
        // NEW APPROACH: Find tagPreview previewFrameParent containers and get titles from each
        // ----------------------------
        let previewVariationTitles = [];
        try {
            // Find all tagPreview previewFrameParent containers and get titles from each
            const allFrames = page.frames();
            for (const f of allFrames) {
                try {
                    const frameTitles = await f.evaluate(() => {
                        try {
                            const containers = document.getElementsByClassName("tagPreview previewFrameParent");
                            const titles = [];
                            
                            Array.from(containers).forEach((container, containerIndex) => {
                                const titleEl = container.querySelector(".previewVariationTitle");
                                if (!titleEl) return;
                                
                                const text = (titleEl.innerText || titleEl.textContent || '').trim();
                                const cleanedText = text.replace(/^Preview\s+of\s+variation\s*/i, '').trim();
                                
                                if (cleanedText || text) {
                                    titles.push({
                                        text: cleanedText || text,
                                        originalText: text,
                                        index: containerIndex, // Use container index
                                        containerIndex: containerIndex,
                                        id: titleEl.id || '',
                                        className: titleEl.className || '',
                                        frameUrl: window.location.href
                                    });
                                }
                            });
                            
                            return titles;
                        } catch (e) {
                            console.log('Error in frame title evaluation:', e);
                            return [];
                        }
                    });
                    
                    if (frameTitles && frameTitles.length > 0) {
                        // Add all titles (don't deduplicate - each container is a separate creative)
                        frameTitles.forEach((title) => {
                            title.index = previewVariationTitles.length;
                            previewVariationTitles.push(title);
                        });
                    }
                } catch (frameError) {
                    // Skip cross-origin frames
                    continue;
                }
            }
        } catch (evalError) {
            console.log('Error evaluating previewVariationTitle:', evalError.message);
            previewVariationTitles = [];
        }

        console.log(`Found ${previewVariationTitles.length} previewVariationTitle elements from tagPreview previewFrameParent containers`);

        // ----------------------------
        // 0.5) Associate frm1_HL_ elements with their parent creative (previewVariationTitle)
        // ----------------------------
        // We need to find frm1_HL_ elements and associate them with their parent previewVariationTitle
        // by traversing the DOM structure
        const creativeElementMap = new Map(); // Map creative title to its frm1_HL_ elements
        
        try {
            // First, try to associate elements in the main page
            const mainPageAssociations = await page.evaluate(() => {
                const associations = [];
                const titleElements = document.getElementsByClassName("previewVariationTitle");
                
                Array.from(titleElements).forEach((titleEl, titleIdx) => {
                    const titleText = (titleEl.innerText || titleEl.textContent || '').trim();
                    const cleanedTitle = titleText.replace(/^Preview\s+of\s+variation\s*/i, '').trim();
                    
                    if (!cleanedTitle) return;
                    
                    // Find the closest parent container that might contain frm1_HL_ elements
                    // Look for frm1_HL_ elements in the same section/container as the title
                    let container = titleEl.parentElement;
                    let foundElements = [];
                    
                    // Try to find frm1_HL_ elements near this title
                    // Look in parent containers up to 5 levels up
                    for (let i = 0; i < 5 && container; i++) {
                        const frm1Elements = container.querySelectorAll('[id^="frm1_HL_"]');
                        if (frm1Elements.length > 0) {
                            foundElements = Array.from(frm1Elements).map(el => ({
                                id: el.id,
                                tagName: el.tagName
                            }));
                            break;
                        }
                        container = container.parentElement;
                    }
                    
                    // Also check siblings and next elements
                    if (foundElements.length === 0) {
                        let nextSibling = titleEl.nextElementSibling;
                        for (let i = 0; i < 3 && nextSibling; i++) {
                            const frm1Elements = nextSibling.querySelectorAll('[id^="frm1_HL_"]');
                            if (frm1Elements.length > 0) {
                                foundElements = Array.from(frm1Elements).map(el => ({
                                    id: el.id,
                                    tagName: el.tagName
                                }));
                                break;
                            }
                            nextSibling = nextSibling.nextElementSibling;
                        }
                    }
                    
                    // Deduplicate by ID
                    const uniqueElements = [];
                    const seenIds = new Set();
                    foundElements.forEach(el => {
                        if (el.id && !seenIds.has(el.id)) {
                            seenIds.add(el.id);
                            uniqueElements.push(el);
                        }
                    });
                    
                    if (cleanedTitle) {
                        associations.push({
                            title: cleanedTitle,
                            elements: uniqueElements
                        });
                    }
                });
                
                return associations;
            });
            
            // Store associations
            mainPageAssociations.forEach(assoc => {
                if (!creativeElementMap.has(assoc.title)) {
                    creativeElementMap.set(assoc.title, []);
                }
                creativeElementMap.get(assoc.title).push(...assoc.elements);
            });
            
        } catch (evalError) {
            console.log('Error associating elements with creatives:', evalError.message);
        }

        // ============================================
        // 3. DETECTION LOGIC
        // ============================================
        
        // A. Detect jvxBase_* iframes
        // Scans the main document DOM
        // Collects all <iframe> elements whose id starts with "jvxBase_"
        let jvxFrames = [];
        try {
            jvxFrames = await page.evaluate(() => {
                try {
                    return Array.from(document.querySelectorAll('iframe'))
                        .filter(f => f.id && f.id.startsWith('jvxBase_'))
                        .map(f => {
                            return { 
                                id: f.id, 
                                src: f.src || ''
                            };
                        });
                } catch (e) {
                    return [];
                }
            });
            console.log(`Found ${jvxFrames.length} jvxBase_* iframes via DOM query`);
        } catch (evalError) {
            console.log('Error evaluating jvxBase_ frames:', evalError.message);
            jvxFrames = [];
        }

        // First, find all creative variations from the main page and map them to jvxBase_ iframes
        const creativeVariationMap = new Map(); // Map jvxBase_ iframe ID to creative variation name
        try {
            const variations = await page.evaluate(() => {
                // Find all elements with class "previewVariationTitle"
                const titleElements = document.querySelectorAll(".previewVariationTitle");
                const containers = document.getElementsByClassName("tagPreview previewFrameParent");
                const results = [];
                
                // First approach: Find previewVariationTitle within tagPreview previewFrameParent containers
                Array.from(containers).forEach((container, containerIndex) => {
                    const titleEl = container.querySelector(".previewVariationTitle");
                    if (!titleEl) {
                        return;
                    }
                    
                    // Get the exact text from previewVariationTitle element
                    const text = (titleEl.innerText || titleEl.textContent || titleEl.text || '').trim();
                    // Remove "Preview of variation" prefix if present, but keep the rest
                    const cleanedText = text.replace(/^Preview\s+of\s+variation\s*/i, '').trim();
                    const finalText = cleanedText || text;
                    
                    if (finalText) {
                        // Find the jvxBase_ iframe within this container
                        const jvxBaseIframe = container.querySelector('iframe[id^="jvxBase_"]');
                        const jvxBaseId = jvxBaseIframe ? (jvxBaseIframe.id || '') : '';
                        
                        if (jvxBaseId) {
                            results.push({
                                jvxBaseId: jvxBaseId,
                                creativeVariation: finalText,
                                rawText: text // Keep original for debugging
                            });
                        }
                    }
                });
                
                // If no results from containers, try direct query of all previewVariationTitle elements
                if (results.length === 0 && titleElements.length > 0) {
                    console.log(`Found ${titleElements.length} previewVariationTitle elements directly, trying to match with jvxBase_ iframes`);
                    Array.from(titleElements).forEach((titleEl, idx) => {
                        const text = (titleEl.innerText || titleEl.textContent || titleEl.text || '').trim();
                        const cleanedText = text.replace(/^Preview\s+of\s+variation\s*/i, '').trim();
                        const finalText = cleanedText || text;
                        
                        if (finalText) {
                            // Try to find nearest jvxBase_ iframe (could be parent, sibling, or in same container)
                            let container = titleEl.closest('.tagPreview, .previewFrameParent, [class*="preview"]');
                            if (!container) {
                                container = titleEl.parentElement;
                            }
                            
                            if (container) {
                                const jvxBaseIframe = container.querySelector('iframe[id^="jvxBase_"]');
                                const jvxBaseId = jvxBaseIframe ? (jvxBaseIframe.id || '') : '';
                                
                                if (jvxBaseId) {
                                    results.push({
                                        jvxBaseId: jvxBaseId,
                                        creativeVariation: finalText,
                                        rawText: text
                                    });
                                }
                            }
                        }
                    });
                }
                
                return results;
            });
            
            variations.forEach(v => {
                creativeVariationMap.set(v.jvxBaseId, v.creativeVariation);
                console.log(`Mapped: ${v.jvxBaseId} -> "${v.creativeVariation}"`);
            });
            
            console.log(`Found ${creativeVariationMap.size} creative variations mapped to jvxBase_ iframes`);
            console.log(`Creative variation map entries:`, Array.from(creativeVariationMap.entries()));
        } catch (e) {
            console.log('Error finding creative variations:', e.message);
        }

        // B. Scan all frames (including nested iframes)
        // Iterates through every frame context using page.frames()
        // Inside each frame: Searches for elements whose IDs start with frm1_HL_
        const frames = page.frames();
        const hlMatches = [];

        console.log(`Scanning ${frames.length} frames for frm1_HL_* elements...`);

        for (const f of frames) {
            try {
                let frameUrl, frameName;
                
                try {
                    frameUrl = f.url();
                    frameName = f.name() || '';
                } catch (urlError) {
                    console.log('Error getting frame URL info:', urlError.message);
                    continue;
                }

                // Try to find the jvxBase_ ID from the main page that corresponds to this frame
                // by checking iframe elements in the main page
                let matchedJvxBaseId = null;
                try {
                    matchedJvxBaseId = await page.evaluate((frameUrl, frameName) => {
                        // Find all jvxBase_ iframes
                        const iframes = document.querySelectorAll('iframe[id^="jvxBase_"]');
                        for (const iframe of iframes) {
                            const iframeId = iframe.id;
                            const iframeSrc = iframe.src || '';
                            const iframeName = iframe.name || '';
                            
                            // Match by name, ID, or URL
                            if (frameName && (iframeName === frameName || iframeId === frameName)) {
                                return iframeId;
                            }
                            if (frameUrl && iframeSrc && (iframeSrc === frameUrl || frameUrl.includes(iframeId) || iframeSrc.includes(frameUrl))) {
                                return iframeId;
                            }
                        }
                        return null;
                    }, frameUrl, frameName);
                    
                    if (matchedJvxBaseId) {
                        console.log(`Matched frame ${frameName || frameUrl} to jvxBase_ ID: ${matchedJvxBaseId}`);
                    }
                } catch (matchError) {
                    console.log('Error matching frame to jvxBase_ ID:', matchError.message);
                }

                // Search for frm1_HL_*, frm2_HL_*, frm3_HL_*, frm4_HL_* and frm1_SL_*, frm2_SL_*, frm3_SL_*, frm4_SL_* elements in this frame
                // Also find creative variation name (previewVariationTitle)
                let matches = null;
                let matches2 = null;
                let matches3 = null;
                let matches4 = null;
                let matchesSL1 = null;
                let matchesSL2 = null;
                let matchesSL3 = null;
                let matchesSL4 = null;
                let creativeVariation = '';
                
                try {
                    const allMatches = await f.evaluate(() => {
                        try {
                            // Mazda model names to check for line breaks
                            // Only check models with spaces (can be split across lines)
                            // Single-word models like "MAZDA6e" and "MAZDA3" cannot be split
                            const mazdaModels = [
                                'MAZDA CX-60',
                                'MAZDA CX-30',
                                'MAZDA2 HYBRID',
                                'MAZDA CX-80',
                                'MAZDA MX-5',
                                'MAZDA CX-5',
                                'MAZDA MX-30'
                            ];
                            
                            // Function to build normalized text and mapping
                            function buildNormalizedTextAndMap(element) {
                                const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT, null, false);
                                let node;
                                let normalized = "";
                                const mapping = [];
                                let lastWasSpace = false;

                                while ((node = walker.nextNode())) {
                                    const txt = node.nodeValue || "";
                                    for (let i = 0; i < txt.length; i++) {
                                        const ch = txt[i];
                                        if (/\s/.test(ch)) {
                                            if (!lastWasSpace && normalized.length > 0) {
                                                normalized += " ";
                                                mapping.push({ node, offset: i });
                                                lastWasSpace = true;
                                            }
                                        } else {
                                            normalized += ch;
                                            mapping.push({ node, offset: i });
                                            lastWasSpace = false;
                                        }
                                    }
                                }

                                if (normalized.length > 0 && normalized[0] === " ") {
                                    normalized = normalized.slice(1);
                                    mapping.shift();
                                }
                                if (normalized.length > 0 && normalized[normalized.length - 1] === " ") {
                                    normalized = normalized.slice(0, -1);
                                    mapping.pop();
                                }

                                return { normalized, mapping };
                            }

                            // Function to create range for phrase
                            function createRangeForPhraseRobust(element, phrase, occurrence = 0, caseInsensitive = false) {
                                if (!element || !phrase) return null;

                                const { normalized, mapping } = buildNormalizedTextAndMap(element);
                                if (normalized.length === 0) return null;

                                const targetNorm = phrase.replace(/\s+/g, " ").trim();
                                const searchNormalized = caseInsensitive ? normalized.toLowerCase() : normalized;
                                const targetToSearch = caseInsensitive ? targetNorm.toLowerCase() : targetNorm;

                                let idx = -1;
                                let from = 0;
                                for (let i = 0; i <= occurrence; i++) {
                                    idx = searchNormalized.indexOf(targetToSearch, from);
                                    if (idx === -1) break;
                                    from = idx + targetToSearch.length;
                                }
                                if (idx === -1) return null;

                                const startIndex = idx;
                                const endIndex = idx + targetToSearch.length - 1;

                                const startMap = mapping[startIndex];
                                const endMap = mapping[endIndex];

                                if (!startMap || !endMap) return null;

                                const startNode = startMap.node;
                                const startOffset = startMap.offset;
                                const endNode = endMap.node;
                                const endOffset = endMap.offset + 1;

                                const range = document.createRange();
                                try {
                                    range.setStart(startNode, startOffset);
                                    range.setEnd(endNode, endOffset);
                                } catch (e) {
                                    return null;
                                }

                                return { range, startNode, startOffset, endNode, endOffset };
                            }

                            // Function to check if phrase is on single line
                            function isPhraseSingleLineRobust(element, phrase, options = {}) {
                                const { occurrence = 0, caseInsensitive = false } = options;
                                const res = createRangeForPhraseRobust(element, phrase, occurrence, caseInsensitive);

                                if (!res) {
                                    return { found: false, singleLine: null, rectsCount: 0, rects: [] };
                                }

                                const rectList = Array.from(res.range.getClientRects());
                                const rectsCount = rectList.length;
                                const singleLine = rectsCount === 1;

                                return { found: true, singleLine, rectsCount, rects: rectList };
                            }

                            // Function to check element for Mazda model names and their line status
                            function checkElementForBrokenModels(element) {
                                const issues = [];
                                const allModels = []; // Track all found models for debugging
                                
                                for (const model of mazdaModels) {
                                    const result = isPhraseSingleLineRobust(element, model, { caseInsensitive: true });
                                    if (result.found) {
                                        allModels.push({
                                            model: model,
                                            found: true,
                                            singleLine: result.singleLine,
                                            rectsCount: result.rectsCount
                                        });
                                        
                                        // Only flag as broken if it's on multiple lines
                                        if (!result.singleLine) {
                                            issues.push({
                                                model: model,
                                                found: true,
                                                singleLine: false,
                                                rectsCount: result.rectsCount
                                            });
                                        }
                                    }
                                }
                                
                                // Log for debugging
                                if (allModels.length > 0) {
                                    console.log(`Element ${element.id}: Found ${allModels.length} model(s):`, allModels.map(m => `${m.model} (${m.singleLine ? 'single line' : m.rectsCount + ' lines'})`));
                                }
                                
                                return issues;
                            }

                            const processElements = (elements) => {
                                return elements.map(el => {
                                    try {
                                        const brokenModels = checkElementForBrokenModels(el);
                                        // Also check all models to get their status (for reporting)
                                        const allModelStatuses = [];
                                        for (const model of mazdaModels) {
                                            const result = isPhraseSingleLineRobust(el, model, { caseInsensitive: true });
                                            if (result.found) {
                                                allModelStatuses.push({
                                                    model: model,
                                                    found: true,
                                                    singleLine: result.singleLine,
                                                    rectsCount: result.rectsCount
                                                });
                                            }
                                        }
                                        return {
                                            id: el.id,
                                            outerHTML: el.outerHTML ? el.outerHTML.substring(0, 500) : '',
                                            brokenModels: brokenModels.length > 0 ? brokenModels : null,
                                            allModelStatuses: allModelStatuses.length > 0 ? allModelStatuses : null
                                        };
                                    } catch (e) {
                                        console.log(`Error processing element ${el.id}:`, e.message);
                                        return {
                                            id: el.id,
                                            outerHTML: '',
                                            brokenModels: null,
                                            allModelStatuses: null
                                        };
                                    }
                                });
                            };
                            
                            // Find HL elements
                            const frm1Elements = Array.from(document.querySelectorAll('[id^="frm1_HL_"]'));
                            const frm2Elements = Array.from(document.querySelectorAll('[id^="frm2_HL_"]'));
                            const frm3Elements = Array.from(document.querySelectorAll('[id^="frm3_HL_"]'));
                            const frm4Elements = Array.from(document.querySelectorAll('[id^="frm4_HL_"]'));
                            
                            // Find SL elements
                            const frm1SLElements = Array.from(document.querySelectorAll('[id^="frm1_SL_"]'));
                            const frm2SLElements = Array.from(document.querySelectorAll('[id^="frm2_SL_"]'));
                            const frm3SLElements = Array.from(document.querySelectorAll('[id^="frm3_SL_"]'));
                            const frm4SLElements = Array.from(document.querySelectorAll('[id^="frm4_SL_"]'));
                            
                            return {
                                frm1: processElements(frm1Elements),
                                frm2: processElements(frm2Elements),
                                frm3: processElements(frm3Elements),
                                frm4: processElements(frm4Elements),
                                frm1SL: processElements(frm1SLElements),
                                frm2SL: processElements(frm2SLElements),
                                frm3SL: processElements(frm3SLElements),
                                frm4SL: processElements(frm4SLElements)
                            };
                        } catch (e) {
                            return { frm1: [], frm2: [], frm3: [], frm4: [], frm1SL: [], frm2SL: [], frm3SL: [], frm4SL: [] };
                        }
                    });
                    
                    matches = allMatches.frm1;
                    matches2 = allMatches.frm2;
                    matches3 = allMatches.frm3;
                    matches4 = allMatches.frm4;
                    matchesSL1 = allMatches.frm1SL;
                    matchesSL2 = allMatches.frm2SL;
                    matchesSL3 = allMatches.frm3SL;
                    matchesSL4 = allMatches.frm4SL;
                    
                    // PRIORITY 1: Get creative variation from the map using matched jvxBase_ ID (most reliable)
                    // This ensures we use the previewVariationTitle value from the main page
                    if (matchedJvxBaseId) {
                        creativeVariation = creativeVariationMap.get(matchedJvxBaseId) || '';
                        if (creativeVariation) {
                            console.log(`✅ Matched creative variation "${creativeVariation}" to frame ${frameName} using jvxBase_ ID: ${matchedJvxBaseId}`);
                        }
                    }
                    
                    // PRIORITY 2: Try by frame name (in case frame name is the jvxBase_ ID)
                    if (!creativeVariation && frameName) {
                        creativeVariation = creativeVariationMap.get(frameName) || '';
                        if (creativeVariation) {
                            console.log(`✅ Matched creative variation "${creativeVariation}" to frame ${frameName} by frame name`);
                        }
                    }
                    
                    // PRIORITY 3: Try matching by URL or partial ID match
                    if (!creativeVariation && frameUrl) {
                        for (const [jvxBaseId, variation] of creativeVariationMap.entries()) {
                            if (frameUrl.includes(jvxBaseId) || (frameName && frameName.includes(jvxBaseId)) || (matchedJvxBaseId && matchedJvxBaseId === jvxBaseId)) {
                                creativeVariation = variation;
                                console.log(`✅ Matched creative variation "${variation}" to frame ${frameName} via URL/ID matching (${jvxBaseId})`);
                                break;
                            }
                        }
                    }
                    
                    // PRIORITY 4: Fallback - try to find it in the frame itself (less reliable, but better than nothing)
                    if (!creativeVariation) {
                        try {
                            const frameVariation = await f.evaluate(() => {
                                const titleEl = document.querySelector('.previewVariationTitle');
                                if (titleEl) {
                                    const text = (titleEl.innerText || titleEl.textContent || '').trim();
                                    return text.replace(/^Preview\s+of\s+variation\s*/i, '').trim() || text;
                                }
                                return '';
                            });
                            creativeVariation = frameVariation || '';
                            if (creativeVariation) {
                                console.log(`⚠️ Found creative variation "${creativeVariation}" inside frame ${frameName} (fallback method)`);
                            }
                        } catch (e) {
                            // Ignore errors
                        }
                    }
                    
                    if (!creativeVariation) {
                        console.log(`❌ Could not find creative variation for frame ${frameName || frameUrl}, matchedJvxBaseId=${matchedJvxBaseId || 'none'}`);
                    }
                    
                    console.log(`Frame ${frameName || frameUrl}: creativeVariation="${creativeVariation}", isJvxBase=${(frameName && frameName.startsWith('jvxBase_')) || (matchedJvxBaseId !== null)}, matchedJvxBaseId=${matchedJvxBaseId || 'none'}`);
                } catch (evalError) {
                    // Cross-origin or other access error
                    console.log(`Cannot access frame ${frameUrl} (likely cross-origin):`, evalError.message);
                    matches = null;
                    matches2 = null;
                    matches3 = null;
                    matches4 = null;
                    matchesSL1 = null;
                    matchesSL2 = null;
                    matchesSL3 = null;
                    matchesSL4 = null;
                    creativeVariation = '';
                }

                // Only add frames that:
                // 1. Are jvxBase_ frames AND have a valid creative variation (from previewVariationTitle)
                // 2. OR have elements found (even if not jvxBase_)
                // Check if it's a jvxBase_ frame (by name or matched ID)
                const isJvxBaseFrame = (frameName && frameName.startsWith('jvxBase_')) || (matchedJvxBaseId !== null);
                const hasValidCreativeVariation = creativeVariation && creativeVariation !== 'N/A' && creativeVariation !== '';
                const hasElements = (matches && matches.length > 0) || 
                                   (matches2 && matches2.length > 0) || 
                                   (matches3 && matches3.length > 0) || 
                                   (matches4 && matches4.length > 0) ||
                                   (matchesSL1 && matchesSL1.length > 0) ||
                                   (matchesSL2 && matchesSL2.length > 0) ||
                                   (matchesSL3 && matchesSL3.length > 0) ||
                                   (matchesSL4 && matchesSL4.length > 0);
                
                // Add if it's a jvxBase_ frame (we'll filter invalid creative variations later)
                // OR if it has elements found
                if (isJvxBaseFrame || hasElements) {
                    // Check if we already have this creative variation (deduplicate)
                    const existingIndex = creativeVariation && creativeVariation !== 'N/A' && creativeVariation !== '' 
                        ? hlMatches.findIndex(m => m.creativeVariation === creativeVariation)
                        : -1;
                    
                    if (existingIndex >= 0 && creativeVariation && creativeVariation !== 'N/A' && creativeVariation !== '') {
                        // Merge elements into existing entry
                        const existing = hlMatches[existingIndex];
                        existing.matches = [...(existing.matches || []), ...(matches || [])];
                        existing.matches2 = [...(existing.matches2 || []), ...(matches2 || [])];
                        existing.matches3 = [...(existing.matches3 || []), ...(matches3 || [])];
                        existing.matches4 = [...(existing.matches4 || []), ...(matches4 || [])];
                        existing.matchesSL1 = [...(existing.matchesSL1 || []), ...(matchesSL1 || [])];
                        existing.matchesSL2 = [...(existing.matchesSL2 || []), ...(matchesSL2 || [])];
                        existing.matchesSL3 = [...(existing.matchesSL3 || []), ...(matchesSL3 || [])];
                        existing.matchesSL4 = [...(existing.matchesSL4 || []), ...(matchesSL4 || [])];
                        console.log(`Merged elements into existing entry for "${creativeVariation}"`);
                    } else {
                        hlMatches.push({
                            frameUrl: frameUrl,
                            frameName: frameName,
                            creativeVariation: creativeVariation || 'N/A',
                            matches: matches || [],
                            matches2: matches2 || [],
                            matches3: matches3 || [],
                            matches4: matches4 || [],
                            matchesSL1: matchesSL1 || [],
                            matchesSL2: matchesSL2 || [],
                            matchesSL3: matchesSL3 || [],
                            matchesSL4: matchesSL4 || []
                        });
                        const totalCount = (matches?.length || 0) + (matches2?.length || 0) + (matches3?.length || 0) + (matches4?.length || 0) +
                                         (matchesSL1?.length || 0) + (matchesSL2?.length || 0) + (matchesSL3?.length || 0) + (matchesSL4?.length || 0);
                        console.log(`Added frame ${frameName || frameUrl} (${creativeVariation || 'N/A'}): HL: frm1=${matches?.length || 0}, frm2=${matches2?.length || 0}, frm3=${matches3?.length || 0}, frm4=${matches4?.length || 0} | SL: frm1=${matchesSL1?.length || 0}, frm2=${matchesSL2?.length || 0}, frm3=${matchesSL3?.length || 0}, frm4=${matchesSL4?.length || 0} (total: ${totalCount})`);
                    }
                } else {
                    console.log(`Skipped frame ${frameName || frameUrl}: not jvxBase_ and no elements found`);
                }
            } catch (err) {
                console.log('Error scanning frame:', err.message);
                // Continue with other frames even if one fails
            }
        }

        // 4. Fallback Detection (Safety Net)
        // If no jvxBase_* iframes are found via DOM queries:
        // Perform a regex scan on raw HTML
        if (!jvxFrames || jvxFrames.length === 0) {
            try {
                console.log('No jvxBase_* iframes found via DOM, trying regex fallback...');
                const html = await page.content();
                const ids = Array.from(html.matchAll(/id=["'](jvxBase_[^"']+)["']/g)).map(m => m[1]);
                jvxFrames = ids.map(id => ({ id, src: '' }));
                console.log(`Fallback regex found ${jvxFrames.length} jvxBase_* iframe IDs`);
            } catch (htmlError) {
                console.log('Error getting page content for fallback:', htmlError.message);
                // Keep jvxFrames as empty array
            }
        }

        // Close browser after detection
        try {
            await browser.close();
        } catch (closeError) {
            console.log('Error closing browser (non-critical):', closeError.message);
        }

        // ============================================
        // 5. API RESPONSE STRUCTURE
        // ============================================
        // Return structured JSON as specified:
        // {
        //   "ok": true,
        //   "count": <number_of_jvxBase_iframes>,
        //   "frames": [
        //     { "id": "jvxBase_xxx", "src": "..." }
        //   ],
        //   "hlMatches": [
        //     {
        //       "frameUrl": "...",
        //       "frameName": "...",
        //       "matches": [
        //         {
        //           "id": "frm1_HL_xxx",
        //           "outerHTML": "<div ...>"
        //         }
        //       ]
        //     }
        //   ]
        // }
        
        // Filter out entries with invalid creative variations (N/A or empty) and deduplicate
        // But keep entries even if creative variation is missing if they have elements
        const validHlMatches = [];
        const seenCreativeVariations = new Set();
        
        (hlMatches || []).forEach(match => {
            const creativeVar = match.creativeVariation || '';
            const hasElements = (match.matches && match.matches.length > 0) ||
                               (match.matches2 && match.matches2.length > 0) ||
                               (match.matches3 && match.matches3.length > 0) ||
                               (match.matches4 && match.matches4.length > 0) ||
                               (match.matchesSL1 && match.matchesSL1.length > 0) ||
                               (match.matchesSL2 && match.matchesSL2.length > 0) ||
                               (match.matchesSL3 && match.matchesSL3.length > 0) ||
                               (match.matchesSL4 && match.matchesSL4.length > 0);
            
            const isValidCreativeVar = creativeVar && 
                                     creativeVar !== 'N/A' && 
                                     creativeVar !== '' &&
                                     creativeVar.trim() !== '';
            
            // Include if it has a valid creative variation (and not duplicate)
            // OR if it has elements (even without creative variation)
            if (isValidCreativeVar) {
                if (!seenCreativeVariations.has(creativeVar)) {
                    seenCreativeVariations.add(creativeVar);
                    validHlMatches.push(match);
                } else {
                    // Merge into existing entry
                    const existingIndex = validHlMatches.findIndex(m => m.creativeVariation === creativeVar);
                    if (existingIndex >= 0) {
                        const existing = validHlMatches[existingIndex];
                        existing.matches = [...(existing.matches || []), ...(match.matches || [])];
                        existing.matches2 = [...(existing.matches2 || []), ...(match.matches2 || [])];
                        existing.matches3 = [...(existing.matches3 || []), ...(match.matches3 || [])];
                        existing.matches4 = [...(existing.matches4 || []), ...(match.matches4 || [])];
                        existing.matchesSL1 = [...(existing.matchesSL1 || []), ...(match.matchesSL1 || [])];
                        existing.matchesSL2 = [...(existing.matchesSL2 || []), ...(match.matchesSL2 || [])];
                        existing.matchesSL3 = [...(existing.matchesSL3 || []), ...(match.matchesSL3 || [])];
                        existing.matchesSL4 = [...(existing.matchesSL4 || []), ...(match.matchesSL4 || [])];
                    }
                }
            } else if (hasElements) {
                // Include entries with elements even if no creative variation
                validHlMatches.push(match);
            }
        });
        
        console.log(`Filtered hlMatches: ${hlMatches.length} total, ${validHlMatches.length} with valid creative variations or elements`);
        console.log(`Valid creative variations:`, Array.from(seenCreativeVariations));
        console.log(`All hlMatches creative variations:`, hlMatches.map(m => m.creativeVariation));
        
        // Collect all broken model issues from valid elements (after filtering)
        const brokenModels = [];
        const allModelStatuses = []; // Track all models found for comprehensive reporting
        (validHlMatches || []).forEach(match => {
            const allElementArrays = [
                ...(match.matches || []),
                ...(match.matches2 || []),
                ...(match.matches3 || []),
                ...(match.matches4 || []),
                ...(match.matchesSL1 || []),
                ...(match.matchesSL2 || []),
                ...(match.matchesSL3 || []),
                ...(match.matchesSL4 || [])
            ];
            
            allElementArrays.forEach(element => {
                if (element) {
                    // Track all model statuses (for comprehensive reporting)
                    if (element.allModelStatuses && element.allModelStatuses.length > 0) {
                        allModelStatuses.push({
                            creativeVariation: match.creativeVariation || 'N/A',
                            elementId: element.id || 'Unknown',
                            elementType: element.id && element.id.includes('_HL_') ? 'HL' : 'SL',
                            modelStatuses: element.allModelStatuses,
                            outerHTML: element.outerHTML || ''
                        });
                    }
                    
                    // Track broken models separately (for the broken models table)
                    if (element.brokenModels && element.brokenModels.length > 0) {
                        brokenModels.push({
                            creativeVariation: match.creativeVariation || 'N/A',
                            elementId: element.id || 'Unknown',
                            elementType: element.id && element.id.includes('_HL_') ? 'HL' : 'SL',
                            brokenModels: element.brokenModels,
                            outerHTML: element.outerHTML || ''
                        });
                    }
                }
            });
        });
        
        console.log(`Found ${allModelStatuses.length} elements with Mazda model names`);
        console.log(`Found ${brokenModels.length} elements with broken Mazda model names (split across lines)`);
        if (allModelStatuses.length > 0) {
            console.log('All model statuses:', allModelStatuses.map(ams => ({
                creativeVariation: ams.creativeVariation,
                elementId: ams.elementId,
                models: ams.modelStatuses.map(m => `${m.model} (${m.singleLine ? 'single line' : m.rectsCount + ' lines'})`)
            })));
        }
        if (brokenModels.length > 0) {
            console.log('Broken models details:', brokenModels.map(bm => ({
                creativeVariation: bm.creativeVariation,
                elementId: bm.elementId,
                models: bm.brokenModels.map(m => m.model)
            })));
        }
        
        const response = {
            ok: true,
            count: jvxFrames.length,
            frames: jvxFrames || [],
            hlMatches: validHlMatches,
            brokenModels: brokenModels,
            allModelStatuses: allModelStatuses
        };
        
        console.log('Sending response:', {
            ok: response.ok,
            count: response.count,
            framesCount: response.frames.length,
            hlMatchesCount: response.hlMatches.length
        });
        
        // Debug: Log what we're sending
        if (response.hlMatches.length > 0) {
            console.log('Sample hlMatch:', JSON.stringify(response.hlMatches[0], null, 2));
        } else {
            console.log('⚠️ No hlMatches found! This will show "No results found"');
            console.log(`Total frames scanned: ${frames.length}`);
            console.log(`Creative variation map size: ${creativeVariationMap.size}`);
            console.log(`Creative variation map entries:`, Array.from(creativeVariationMap.entries()));
        }
        
        return res.json(response);
        const allHLElements = [];
        const allHLElements2 = [];
        const allHLElements3 = [];
        const allHLElements4 = [];
        const uniqueElementIds = new Set();
        const uniqueElementIds2 = new Set();
        const uniqueElementIds3 = new Set();
        const uniqueElementIds4 = new Set();
        const usedElementIds = new Set(); // Track which frm1 elements have been assigned to creatives
        const usedElementIds2 = new Set(); // Track which frm2 elements have been assigned to creatives
        const usedElementIds3 = new Set(); // Track which frm3 elements have been assigned to creatives
        const usedElementIds4 = new Set(); // Track which frm4 elements have been assigned to creatives
        
        // Collect all frm1_HL_*, frm2_HL_*, frm3_HL_*, and frm4_HL_* elements BY JVXBASE_ IFRAME ID
        // Each jvxBase_ iframe has its own elements, so we need to track which iframe ID they came from
        const elementsByJvxBaseId = new Map(); // Map jvxBase_* iframe ID to elements
        
        console.log(`\n=== COLLECTING ELEMENTS BY JVXBASE_ IFRAME ID ===`);
        
        // Collect elements and group them by jvxBase_* iframe ID (not index!)
        hlMatches.forEach((hlMatch, hlMatchIndex) => {
            // Get the jvxBase_* iframe ID from the frame name
            let jvxBaseId = '';
            if (hlMatch.frameName && hlMatch.frameName.startsWith('jvxBase_')) {
                jvxBaseId = hlMatch.frameName; // Use the full frame name as ID (e.g., "jvxBase_6978809ad0eb8")
            } else if (hlMatch.frameUrl) {
                // Try to extract from URL or frame info
                // Sometimes the frame name might be in the URL or other properties
                const urlMatch = hlMatch.frameUrl.match(/jvxBase_[\w]+/);
                if (urlMatch) {
                    jvxBaseId = urlMatch[0];
                }
            }
            
            // If we still don't have an ID, skip this match (it's not a jvxBase_ frame)
            if (!jvxBaseId) {
                console.log(`Skipping frame ${hlMatchIndex}: Not a jvxBase_ frame (frameName: ${hlMatch.frameName || 'N/A'})`);
                return;
            }
            
            console.log(`Processing jvxBase_ frame: ${jvxBaseId}`);
            
            // Initialize the map entry for this jvxBase_ ID if it doesn't exist
            if (!elementsByJvxBaseId.has(jvxBaseId)) {
                elementsByJvxBaseId.set(jvxBaseId, { frm1: [], frm2: [], frm3: [], frm4: [] });
            }
            
            // Collect frm1_HL_* elements and store by jvxBase_ ID
            if (hlMatch.matches && hlMatch.matches.length > 0) {
                hlMatch.matches.forEach(match => {
                    if (match.id && !uniqueElementIds.has(match.id)) {
                        uniqueElementIds.add(match.id);
                        const elementWithFrame = {
                            ...match,
                            frameUrl: hlMatch.frameUrl || '',
                            frameName: hlMatch.frameName || '',
                            jvxBaseId: jvxBaseId
                        };
                        allHLElements.push(elementWithFrame);
                        elementsByJvxBaseId.get(jvxBaseId).frm1.push(elementWithFrame);
                        console.log(`  Added frm1 element ${match.id} to ${jvxBaseId}`);
                    }
                });
            }
            
            // Collect frm2_HL_* elements and store by jvxBase_ ID
            if (hlMatch.matches2 && hlMatch.matches2.length > 0) {
                hlMatch.matches2.forEach(match => {
                    if (match.id && !uniqueElementIds2.has(match.id)) {
                        uniqueElementIds2.add(match.id);
                        const elementWithFrame = {
                            ...match,
                            frameUrl: hlMatch.frameUrl || '',
                            frameName: hlMatch.frameName || '',
                            jvxBaseId: jvxBaseId
                        };
                        allHLElements2.push(elementWithFrame);
                        elementsByJvxBaseId.get(jvxBaseId).frm2.push(elementWithFrame);
                        console.log(`  Added frm2 element ${match.id} to ${jvxBaseId}`);
                    }
                });
            }
            
            // Collect frm3_HL_* elements and store by jvxBase_ ID
            if (hlMatch.matches3 && hlMatch.matches3.length > 0) {
                hlMatch.matches3.forEach(match => {
                    if (match.id && !uniqueElementIds3.has(match.id)) {
                        uniqueElementIds3.add(match.id);
                        const elementWithFrame = {
                            ...match,
                            frameUrl: hlMatch.frameUrl || '',
                            frameName: hlMatch.frameName || '',
                            jvxBaseId: jvxBaseId
                        };
                        allHLElements3.push(elementWithFrame);
                        elementsByJvxBaseId.get(jvxBaseId).frm3.push(elementWithFrame);
                        console.log(`  Added frm3 element ${match.id} to ${jvxBaseId}`);
                    }
                });
            }
            
            // Collect frm4_HL_* elements and store by jvxBase_ ID
            if (hlMatch.matches4 && hlMatch.matches4.length > 0) {
                hlMatch.matches4.forEach(match => {
                    if (match.id && !uniqueElementIds4.has(match.id)) {
                        uniqueElementIds4.add(match.id);
                        const elementWithFrame = {
                            ...match,
                            frameUrl: hlMatch.frameUrl || '',
                            frameName: hlMatch.frameName || '',
                            jvxBaseId: jvxBaseId
                        };
                        allHLElements4.push(elementWithFrame);
                        elementsByJvxBaseId.get(jvxBaseId).frm4.push(elementWithFrame);
                        console.log(`  Added frm4 element ${match.id} to ${jvxBaseId}`);
                    }
                });
            }
        });
        
        console.log(`\nElements grouped by jvxBase_ iframe ID:`);
        elementsByJvxBaseId.forEach((elements, jvxBaseId) => {
            console.log(`  ${jvxBaseId}: frm1=${elements.frm1.length}, frm2=${elements.frm2.length}, frm3=${elements.frm3.length}, frm4=${elements.frm4.length}`);
        });

        console.log(`\n=== ELEMENT COLLECTION SUMMARY ===`);
        console.log(`Collected ${allHLElements.length} unique frm1_HL_* elements from ${hlMatches.length} frames`);
        console.log(`Collected ${allHLElements2.length} unique frm2_HL_* elements from ${hlMatches.length} frames`);
        console.log(`Collected ${allHLElements3.length} unique frm3_HL_* elements from ${hlMatches.length} frames`);
        console.log(`Collected ${allHLElements4.length} unique frm4_HL_* elements from ${hlMatches.length} frames`);
        console.log(`\nAvailable frm1 element IDs (${allHLElements.length} total):`, allHLElements.map((el, idx) => `${idx}:${el.id}`));
        console.log(`Available frm2 element IDs (${allHLElements2.length} total):`, allHLElements2.map((el, idx) => `${idx}:${el.id}`));
        console.log(`Available frm3 element IDs (${allHLElements3.length} total):`, allHLElements3.map((el, idx) => `${idx}:${el.id}`));
        console.log(`Available frm4 element IDs (${allHLElements4.length} total):`, allHLElements4.map((el, idx) => `${idx}:${el.id}`));
        console.log(`=====================================\n`);

        // NEW APPROACH: Search for frm1_HL_, frm2_HL_, frm3_HL_, and frm4_HL_ elements in the same context as each creative title
        // This will be done by searching in ALL frames for each title
        const creativeElementAssociations = new Map(); // Map creative title to frm1 element IDs
        const creativeElementAssociations2 = new Map(); // Map creative title to frm2 element IDs
        const creativeElementAssociations3 = new Map(); // Map creative title to frm3 element IDs
        const creativeElementAssociations4 = new Map(); // Map creative title to frm4 element IDs
        const titleFrameMap = new Map(); // Map creative title to its frame URL
        
        // Search for elements in the context of each creative title across ALL frames
        // NEW APPROACH: Find tagPreview previewFrameParent containers and search within each
        const allFramesForSearch = page.frames();
        for (const f of allFramesForSearch) {
            try {
                const frameUrl = f.url();
                const frameAssociations = await f.evaluate(() => {
                    const results = [];
                    
                    // Find all tagPreview previewFrameParent containers
                    const previewContainers = document.getElementsByClassName("tagPreview previewFrameParent");
                    console.log(`Found ${previewContainers.length} tagPreview previewFrameParent containers`);
                    
                    Array.from(previewContainers).forEach((container, containerIndex) => {
                        // Find previewVariationTitle within this container
                        const titleEl = container.querySelector(".previewVariationTitle");
                        if (!titleEl) {
                            console.log(`Container ${containerIndex}: No previewVariationTitle found`);
                            return;
                        }
                        
                        const titleText = (titleEl.innerText || titleEl.textContent || '').trim();
                        const cleanedTitle = titleText.replace(/^Preview\s+of\s+variation\s*/i, '').trim();
                        
                        if (!cleanedTitle) {
                            console.log(`Container ${containerIndex}: Empty title`);
                            return;
                        }
                        
                        // CRITICAL: Find the jvxBase_* iframe ID within this container
                        // This is the key to matching elements correctly
                        const jvxBaseIframe = container.querySelector('iframe[id^="jvxBase_"]');
                        const jvxBaseIframeId = jvxBaseIframe ? (jvxBaseIframe.id || '') : '';
                        
                        console.log(`Container ${containerIndex} (${cleanedTitle}): Found jvxBase_ iframe ID: ${jvxBaseIframeId || 'NOT FOUND'}`);
                        
                        // Skip if we already have this title in results (but allow duplicates if they're in different containers)
                        // We'll handle deduplication later based on container index
                        
                        let foundElementIds = [];
                        let foundElementIds2 = [];
                        let foundElementIds3 = [];
                        let foundElementIds4 = [];
                        
                        // Search for elements WITHIN this tagPreview previewFrameParent container
                        // IMPORTANT: Elements are likely in iframes (jvxBase_*), so we need to search those too
                        const searchForElements = (prefix) => {
                            let found = [];
                            
                            // Strategy 1: Search within the container first (most specific)
                            const elementsInContainer = container.querySelectorAll(`[id^="${prefix}"]`);
                            if (elementsInContainer.length > 0) {
                                found = Array.from(elementsInContainer).map(el => el.id).filter(id => id);
                                console.log(`Container ${containerIndex} (${cleanedTitle}): Found ${found.length} ${prefix} elements directly in container`);
                                return found;
                            }
                            
                            // Strategy 2: Search in iframes within this container
                            // Find all iframes (especially jvxBase_*) within this container
                            const iframesInContainer = container.querySelectorAll('iframe');
                            console.log(`Container ${containerIndex} (${cleanedTitle}): Found ${iframesInContainer.length} iframes in container`);
                            
                            // Note: We can't access iframe content from main frame due to cross-origin restrictions
                            // So we'll use index-based matching instead
                            // But first, let's try to find elements by searching the container's DOM tree
                            
                            // Strategy 3: Search from titleEl upward within container
                            let current = titleEl;
                            for (let i = 0; i < 30 && current; i++) {
                                // Check if we've gone outside the container
                                if (!container.contains(current)) break;
                                
                                const elements = current.querySelectorAll(`[id^="${prefix}"]`);
                                if (elements.length > 0) {
                                    found = Array.from(elements).map(el => el.id).filter(id => id);
                                    console.log(`Container ${containerIndex} (${cleanedTitle}): Found ${found.length} ${prefix} elements in DOM tree`);
                                    break;
                                }
                                current = current.parentElement;
                                if (!current) break;
                            }
                            
                            // Strategy 4: Check next siblings within container
                            if (found.length === 0) {
                                let sibling = titleEl.nextElementSibling;
                                for (let i = 0; i < 30 && sibling; i++) {
                                    if (!container.contains(sibling)) break;
                                    const elements = sibling.querySelectorAll(`[id^="${prefix}"]`);
                                    if (elements.length > 0) {
                                        found = Array.from(elements).map(el => el.id).filter(id => id);
                                        break;
                                    }
                                    sibling = sibling.nextElementSibling;
                                    if (!sibling) break;
                                }
                            }
                            
                            // Strategy 5: Check previous siblings within container
                            if (found.length === 0) {
                                let sibling = titleEl.previousElementSibling;
                                for (let i = 0; i < 30 && sibling; i++) {
                                    if (!container.contains(sibling)) break;
                                    const elements = sibling.querySelectorAll(`[id^="${prefix}"]`);
                                    if (elements.length > 0) {
                                        found = Array.from(elements).map(el => el.id).filter(id => id);
                                        break;
                                    }
                                    sibling = sibling.previousElementSibling;
                                    if (!sibling) break;
                                }
                            }
                            
                            // Strategy 6: If no elements found in container DOM, return empty
                            // We'll rely on index-based matching in the association function
                            // This is because elements are in iframes which we can't access from main frame
                            if (found.length === 0) {
                                console.log(`Container ${containerIndex} (${cleanedTitle}): No ${prefix} elements found in container DOM (likely in iframe)`);
                            }
                            
                            return found;
                        };
                        
                        foundElementIds = searchForElements('frm1_HL_');
                        foundElementIds2 = searchForElements('frm2_HL_');
                        foundElementIds3 = searchForElements('frm3_HL_');
                        foundElementIds4 = searchForElements('frm4_HL_');
                        
                        console.log(`Container ${containerIndex} (${cleanedTitle}): frm1=${foundElementIds.length}, frm2=${foundElementIds2.length}, frm3=${foundElementIds3.length}, frm4=${foundElementIds4.length}`);
                        
                        results.push({
                            title: cleanedTitle,
                            elementIds: [...new Set(foundElementIds)],
                            elementIds2: [...new Set(foundElementIds2)],
                            elementIds3: [...new Set(foundElementIds3)],
                            elementIds4: [...new Set(foundElementIds4)],
                            frameUrl: window.location.href,
                            titleIndex: containerIndex, // Use container index instead of title index
                            containerIndex: containerIndex,
                            jvxBaseIframeId: jvxBaseIframeId // Store the jvxBase_* iframe ID for this container
                        });
                    });
                    
                    return results;
                });
                
                // Store associations for all element types
                // Use a composite key: title + containerIndex to handle multiple containers with same title
                // Also store the jvxBase_ iframe ID for matching
                frameAssociations.forEach(assoc => {
                    const key = `${assoc.title}__container_${assoc.containerIndex || assoc.titleIndex}`;
                    
                    // Store frm1 associations
                    creativeElementAssociations.set(key, assoc.elementIds || []);
                    
                    // Store frm2 associations
                    creativeElementAssociations2.set(key, assoc.elementIds2 || []);
                    
                    // Store frm3 associations
                    creativeElementAssociations3.set(key, assoc.elementIds3 || []);
                    
                    // Store frm4 associations
                    creativeElementAssociations4.set(key, assoc.elementIds4 || []);
                    
                    if (assoc.frameUrl) {
                        titleFrameMap.set(key, assoc.frameUrl);
                    }
                    
                    // Store the jvxBase_ iframe ID for this container
                    if (assoc.jvxBaseIframeId) {
                        // Create a map: container key → jvxBase_ iframe ID
                        if (!titleFrameMap.has(`jvxBase_${key}`)) {
                            titleFrameMap.set(`jvxBase_${key}`, assoc.jvxBaseIframeId);
                        }
                        console.log(`Container "${assoc.title}" (index ${assoc.containerIndex}) → jvxBase_ iframe: ${assoc.jvxBaseIframeId}`);
                    }
                });
            } catch (frameError) {
                // Skip cross-origin frames
                continue;
            }
        }
        
        console.log(`Found associations for ${creativeElementAssociations.size} creatives`);

        // Helper function to associate elements for a given type
        // NEW: Match container to its jvxBase_* iframe ID, then get elements from that iframe
        const associateElementsForType = (titleText, titleFrame, idx, containerIndex, elementType, allElements, usedElementSet, creativeAssociationsMap, prefix, elementsByJvxBaseId, titleFrameMap) => {
            let creativeElements = [];
            const unusedElements = allElements.filter(el => !usedElementSet.has(el.id));
            
            console.log(`[${elementType}] Associating for "${titleText}" (container ${containerIndex}, idx ${idx})`);
            console.log(`[${elementType}] Total elements: ${allElements.length}, Unused: ${unusedElements.length}`);
            
            // Strategy 1: jvxBase_ iframe ID matching (MOST RELIABLE)
            // Match container to its jvxBase_* iframe ID, then get elements from that iframe
            const containerKey = `${titleText}__container_${containerIndex}`;
            const jvxBaseId = titleFrameMap.get(`jvxBase_${containerKey}`);
            
            if (jvxBaseId && elementsByJvxBaseId && elementsByJvxBaseId.has(jvxBaseId)) {
                const frameElements = elementsByJvxBaseId.get(jvxBaseId);
                
                // Get elements for this element type from this jvxBase_ iframe
                let frameElementsForType = [];
                if (elementType === 'frm1') {
                    frameElementsForType = frameElements.frm1 || [];
                } else if (elementType === 'frm2') {
                    frameElementsForType = frameElements.frm2 || [];
                } else if (elementType === 'frm3') {
                    frameElementsForType = frameElements.frm3 || [];
                } else if (elementType === 'frm4') {
                    frameElementsForType = frameElements.frm4 || [];
                }
                
                console.log(`[Strategy 1 - jvxBase_ ID Match] Container "${titleText}" (${containerIndex}) → jvxBase_ iframe: ${jvxBaseId} → Found ${frameElementsForType.length} ${elementType} elements`);
                
                // Get the first unused element from this iframe
                for (const frameElement of frameElementsForType) {
                    if (!usedElementSet.has(frameElement.id)) {
                        creativeElements = [frameElement];
                        usedElementSet.add(frameElement.id);
                        console.log(`[Strategy 1 - jvxBase_ ID Match] ✅ Container "${titleText}" (${containerIndex}) → ${jvxBaseId} → Element: ${frameElement.id}`);
                        break;
                    }
                }
                
                if (creativeElements.length === 0 && frameElementsForType.length > 0) {
                    console.log(`[Strategy 1 - jvxBase_ ID Match] ⚠️ Container "${titleText}" (${containerIndex}) → All elements from ${jvxBaseId} already used`);
                } else if (frameElementsForType.length === 0) {
                    console.log(`[Strategy 1 - jvxBase_ ID Match] ⚠️ Container "${titleText}" (${containerIndex}) → No ${elementType} elements found in ${jvxBaseId}`);
                }
            } else {
                if (!jvxBaseId) {
                    console.log(`[Strategy 1 - jvxBase_ ID Match] ⚠️ Container "${titleText}" (${containerIndex}) → No jvxBase_ iframe ID found for container key: ${containerKey}`);
                } else {
                    console.log(`[Strategy 1 - jvxBase_ ID Match] ⚠️ Container "${titleText}" (${containerIndex}) → jvxBase_ iframe ${jvxBaseId} not found in elementsByJvxBaseId`);
                    console.log(`[Strategy 1] Available jvxBase_ IDs:`, Array.from(elementsByJvxBaseId.keys()));
                }
            }
            
            // Strategy 2: Fallback to index-based matching if frame-based didn't work
            if (creativeElements.length === 0 && allElements.length > 0 && containerIndex !== undefined && containerIndex >= 0) {
                // Try to get the element at the exact container index
                if (containerIndex < allElements.length) {
                    const elementAtIndex = allElements[containerIndex];
                    if (elementAtIndex && !usedElementSet.has(elementAtIndex.id)) {
                        creativeElements = [elementAtIndex];
                        usedElementSet.add(elementAtIndex.id);
                        console.log(`[Strategy 2 - Index Match] ✅ Container ${containerIndex} → Element at index ${containerIndex}: ${elementAtIndex.id}`);
                    } else {
                        console.log(`[Strategy 2 - Index Match] ⚠️ Container ${containerIndex} → Element at index ${containerIndex} already used`);
                    }
                }
            }
            
            // Strategy 3: Try DOM association (only if frame/index-based didn't work)
            if (creativeElements.length === 0) {
                const associationKey = `${titleText}__container_${containerIndex !== undefined ? containerIndex : idx}`;
                console.log(`[Strategy 3 - DOM Association] Checking key: ${associationKey}`);
                console.log(`[Strategy 3] Association map has key: ${creativeAssociationsMap.has(associationKey)}`);
                
                if (creativeAssociationsMap.has(associationKey)) {
                    const associatedIds = creativeAssociationsMap.get(associationKey);
                    console.log(`[Strategy 3] Associated IDs for this key:`, associatedIds);
                    
                    // Find first unused element from associations
                    for (const elementId of associatedIds) {
                        let fullElement = allElements.find(el => 
                            el.id === elementId && 
                            !usedElementSet.has(el.id)
                        );
                        
                        if (fullElement) {
                            creativeElements = [fullElement];
                            usedElementSet.add(fullElement.id);
                            console.log(`[Strategy 3 - DOM Association] ✅ Found via key ${associationKey}: ${fullElement.id}`);
                            break;
                        } else {
                            console.log(`[Strategy 3] Element ${elementId} not found or already used`);
                        }
                    }
                } else {
                    console.log(`[Strategy 3] No association found for key: ${associationKey}`);
                    console.log(`[Strategy 3] Available keys:`, Array.from(creativeAssociationsMap.keys()));
                }
            }
            
            // Strategy 4: Pattern matching by number - only unused elements
            if (creativeElements.length === 0) {
                const creativeNumber = titleText.match(/\d+/); // Extract "10" from "Char_Limit-10"
                
                if (creativeNumber) {
                    const numStr = creativeNumber[0];
                    console.log(`[Strategy 4 - Pattern Match] Looking for number "${numStr}" in element IDs`);
                    
                    // Find ALL unused elements containing this number
                    const matchingElements = unusedElements.filter(el => {
                        if (!el.id) return false;
                        return el.id.includes(numStr);
                    });
                    
                    console.log(`[Strategy 4] Found ${matchingElements.length} elements containing "${numStr}":`, matchingElements.map(el => el.id));
                    
                    if (matchingElements.length > 0) {
                        // Prefer exact match patterns
                        const exactMatch = matchingElements.find(el => {
                            const id = el.id;
                            return id.includes(`-${numStr}`) ||
                                   id.includes(`_${numStr}`) ||
                                   id.includes(`${numStr}-`) ||
                                   id.includes(`${numStr}_`) ||
                                   id.includes(`x${numStr}`) ||
                                   id.includes(`${numStr}x`) ||
                                   id.endsWith(numStr) ||
                                   id.includes(`${prefix}${numStr}`);
                        });
                        
                        const selectedElement = exactMatch || matchingElements[0];
                        creativeElements = [selectedElement];
                        usedElementSet.add(selectedElement.id);
                        console.log(`[Strategy 4 - Pattern Match] ✅ Selected: ${selectedElement.id}`);
                    }
                } else {
                    console.log(`[Strategy 4] No number found in title: "${titleText}"`);
                }
            }
            
            // Strategy 5: Same frame elements by index (only unused)
            if (creativeElements.length === 0 && titleFrame) {
                const sameFrameElements = unusedElements.filter(el => 
                    el.frameUrl === titleFrame
                );
                
                if (sameFrameElements.length > 0) {
                    const frameTitleIndex = previewVariationTitles
                        .filter(t => (titleFrameMap.get(t.text) || '') === titleFrame)
                        .findIndex(t => t.text === titleText);
                    
                    if (frameTitleIndex >= 0 && sameFrameElements.length > frameTitleIndex) {
                        creativeElements = [sameFrameElements[frameTitleIndex]];
                        usedElementSet.add(sameFrameElements[frameTitleIndex].id);
                    } else if (sameFrameElements.length > 0) {
                        creativeElements = [sameFrameElements[0]];
                        usedElementSet.add(sameFrameElements[0].id);
                    }
                }
            }
            
            // Strategy 6: Sequential assignment by global index (only unused) - GUARANTEED TO ASSIGN
            if (creativeElements.length === 0) {
                if (unusedElements.length > 0) {
                    // Assign by creative index - ensure each creative gets a unique element
                    const elementIndex = idx % unusedElements.length;
                    const selectedElement = unusedElements[elementIndex];
                    creativeElements = [selectedElement];
                    usedElementSet.add(selectedElement.id);
                        console.log(`[Strategy 6 - Sequential] Creative ${idx} → Element ${elementIndex}: ${selectedElement.id}`);
                }
            }
            
            // Deduplicate by ID
            const finalUniqueElements = [];
            const seenElementIds = new Set();
            creativeElements.forEach(el => {
                if (el.id && !seenElementIds.has(el.id)) {
                    seenElementIds.add(el.id);
                    finalUniqueElements.push(el);
                }
            });
            
            if (finalUniqueElements.length === 0) {
                console.log(`[${elementType}] ⚠️ No elements found for "${titleText}" after all strategies`);
            } else {
                console.log(`[${elementType}] ✅ Final result for "${titleText}":`, finalUniqueElements.map(el => el.id));
            }
            
            return finalUniqueElements;
        };

        // If we have previewVariationTitles, create a creative entry for each
        // NEW APPROACH: Treat each creative as a separate DOM - find its iframe and get elements directly
        if (previewVariationTitles.length > 0) {
            console.log(`\n=== STARTING CREATIVE ASSOCIATION (Direct DOM Access) ===`);
            console.log(`Total creatives: ${previewVariationTitles.length}`);
            console.log(`=====================================\n`);
            
            // First, get all frames and create a mapping of iframe IDs to Puppeteer frames
            const allFrames = page.frames();
            const frameMap = new Map(); // Map iframe ID to Puppeteer frame
            
            // Try to map frames by evaluating the main page to find all jvxBase_ iframes
            try {
                const iframeInfo = await page.evaluate(() => {
                    const iframes = document.querySelectorAll('iframe[id^="jvxBase_"]');
                    return Array.from(iframes).map(iframe => ({
                        id: iframe.id || '',
                        name: iframe.name || '',
                        src: iframe.src || ''
                    }));
                });
                
                console.log(`Found ${iframeInfo.length} jvxBase_ iframes in main page`);
                
                // Try to match each iframe to a Puppeteer frame
                // Match by index: iframe at index N should map to frame at index N+1 (main frame is 0)
                for (let i = 0; i < iframeInfo.length; i++) {
                    const iframe = iframeInfo[i];
                    let matchedFrame = null;
                    
                    // Method 1: Try by name
                    if (iframe.name) {
                        matchedFrame = allFrames.find(f => {
                            try {
                                return f.name() === iframe.name || f.name() === iframe.id;
                            } catch (e) {
                                return false;
                            }
                        });
                    }
                    
                    // Method 2: Try by URL
                    if (!matchedFrame && iframe.src) {
                        matchedFrame = allFrames.find(f => {
                            try {
                                const url = f.url();
                                return url && (url === iframe.src || url.includes(iframe.id));
                            } catch (e) {
                                return false;
                            }
                        });
                    }
                    
                    // Method 3: Try by index (main frame is 0, first iframe should be at index 1, etc.)
                    if (!matchedFrame && i + 1 < allFrames.length) {
                        matchedFrame = allFrames[i + 1];
                    }
                    
                    if (matchedFrame) {
                        frameMap.set(iframe.id, matchedFrame);
                        console.log(`✅ Mapped iframe ${iframe.id} to Puppeteer frame (index ${i + 1})`);
                    } else {
                        console.log(`⚠️ Could not map iframe ${iframe.id} to Puppeteer frame`);
                    }
                }
            } catch (e) {
                console.log(`Error creating frame map: ${e.message}`);
            }
            
            // Process each creative separately - treat each as its own DOM
            for (let idx = 0; idx < previewVariationTitles.length; idx++) {
                const title = previewVariationTitles[idx];
                const containerIndex = title.containerIndex !== undefined ? title.containerIndex : idx;
                
                console.log(`\n[Creative ${idx + 1}] Processing: "${title.text}" (container ${containerIndex})`);
                
                // Step 1: Find the jvxBase_ iframe ID for this container
                let jvxBaseIframeId = '';
                let jvxBaseFrame = null;
                
                try {
                    // Find the container and its jvxBase_ iframe in the main page
                    const containerInfo = await page.evaluate((containerIdx) => {
                        const containers = document.getElementsByClassName("tagPreview previewFrameParent");
                        if (containerIdx >= containers.length) return null;
                        
                        const container = containers[containerIdx];
                        const jvxBaseIframe = container.querySelector('iframe[id^="jvxBase_"]');
                        
                        return {
                            jvxBaseIframeId: jvxBaseIframe ? (jvxBaseIframe.id || '') : '',
                            hasIframe: !!jvxBaseIframe
                        };
                    }, containerIndex);
                    
                    if (containerInfo && containerInfo.jvxBaseIframeId) {
                        jvxBaseIframeId = containerInfo.jvxBaseIframeId;
                        console.log(`[Creative ${idx + 1}] Found jvxBase_ iframe ID: ${jvxBaseIframeId}`);
                        
                        // Step 2: Get the Puppeteer frame from our map
                        jvxBaseFrame = frameMap.get(jvxBaseIframeId);
                        
                        if (!jvxBaseFrame) {
                            // Fallback: Try to find frame by iterating through all frames
                            console.log(`[Creative ${idx + 1}] Frame not in map, trying fallback methods...`);
                            for (const frame of allFrames) {
                                try {
                                    const frameName = frame.name();
                                    const frameUrl = frame.url();
                                    
                                    if (frameName === jvxBaseIframeId || 
                                        (frameUrl && frameUrl.includes(jvxBaseIframeId))) {
                                        jvxBaseFrame = frame;
                                        frameMap.set(jvxBaseIframeId, frame); // Cache it
                                        console.log(`[Creative ${idx + 1}] Found frame via fallback`);
                                        break;
                                    }
                                } catch (e) {
                                    continue;
                                }
                            }
                        }
                    } else {
                        console.log(`[Creative ${idx + 1}] ⚠️ No jvxBase_ iframe found in container ${containerIndex}`);
                    }
                } catch (e) {
                    console.log(`[Creative ${idx + 1}] Error finding iframe: ${e.message}`);
                }
                
                // Step 3: Access the iframe's DOM and get elements directly
                let creativeElements = [];
                let creativeElements2 = [];
                let creativeElements3 = [];
                let creativeElements4 = [];
                
                if (jvxBaseFrame) {
                    try {
                        console.log(`[Creative ${idx + 1}] Accessing iframe ${jvxBaseIframeId} to get elements...`);
                        
                        const iframeElements = await jvxBaseFrame.evaluate(() => {
                            const processElements = (elements) => {
                                return Array.from(elements).map(el => {
                                    try {
                                        return {
                                            id: el.id || '',
                                            tagName: el.tagName || '',
                                            className: el.className || '',
                                            innerText: (el.innerText || el.textContent || '').substring(0, 100),
                                            outerHTML: (el.outerHTML || '').substring(0, 200)
                                        };
                                    } catch (e) {
                                        return { id: el.id || '', tagName: el.tagName || '' };
                                    }
                                });
                            };
                            
                            const frm1Elements = document.querySelectorAll('[id^="frm1_HL_"]');
                            const frm2Elements = document.querySelectorAll('[id^="frm2_HL_"]');
                            const frm3Elements = document.querySelectorAll('[id^="frm3_HL_"]');
                            const frm4Elements = document.querySelectorAll('[id^="frm4_HL_"]');
                            
                            return {
                                frm1: processElements(frm1Elements),
                                frm2: processElements(frm2Elements),
                                frm3: processElements(frm3Elements),
                                frm4: processElements(frm4Elements)
                            };
                        });
                        
                        creativeElements = iframeElements.frm1 || [];
                        creativeElements2 = iframeElements.frm2 || [];
                        creativeElements3 = iframeElements.frm3 || [];
                        creativeElements4 = iframeElements.frm4 || [];
                        
                        console.log(`[Creative ${idx + 1}] ✅ Found in iframe ${jvxBaseIframeId}:`);
                        console.log(`  - frm1: ${creativeElements.length} elements`, creativeElements.map(el => el.id));
                        console.log(`  - frm2: ${creativeElements2.length} elements`, creativeElements2.map(el => el.id));
                        console.log(`  - frm3: ${creativeElements3.length} elements`, creativeElements3.map(el => el.id));
                        console.log(`  - frm4: ${creativeElements4.length} elements`, creativeElements4.map(el => el.id));
                    } catch (iframeError) {
                        console.log(`[Creative ${idx + 1}] ⚠️ Cannot access iframe ${jvxBaseIframeId}: ${iframeError.message}`);
                        console.log(`[Creative ${idx + 1}] Error details:`, iframeError.stack);
                    }
                } else {
                    console.log(`[Creative ${idx + 1}] ⚠️ Could not find Puppeteer frame for ${jvxBaseIframeId}`);
                }
                
                creatives.push({
                    index: creatives.length + 1,
                    title: title.text,
                    originalTitle: title.originalText,
                    jvxBaseIframeId: jvxBaseIframeId || '', // Store the jvxBase_ iframe ID
                    hasFrm1HLElements: creativeElements.length > 0,
                    frm1HLElements: creativeElements,
                    hasFrm2HLElements: creativeElements2.length > 0,
                    frm2HLElements: creativeElements2,
                    hasFrm3HLElements: creativeElements3.length > 0,
                    frm3HLElements: creativeElements3,
                    hasFrm4HLElements: creativeElements4.length > 0,
                    frm4HLElements: creativeElements4
                });
            }
        } else {
            // If no previewVariationTitles found, create a single creative entry if we have elements
            if (allHLElements.length > 0 || allHLElements2.length > 0 || allHLElements3.length > 0 || allHLElements4.length > 0 || hlMatches.length > 0) {
                // Take only the first unique element for each type
                const firstElement = allHLElements.length > 0 ? [allHLElements[0]] : [];
                const firstElement2 = allHLElements2.length > 0 ? [allHLElements2[0]] : [];
                const firstElement3 = allHLElements3.length > 0 ? [allHLElements3[0]] : [];
                const firstElement4 = allHLElements4.length > 0 ? [allHLElements4[0]] : [];
                
                creatives.push({
                    index: 1,
                    title: 'Creative (No title found)',
                    originalTitle: '',
                    hasFrm1HLElements: firstElement.length > 0,
                    frm1HLElements: firstElement,
                    hasFrm2HLElements: firstElement2.length > 0,
                    frm2HLElements: firstElement2,
                    hasFrm3HLElements: firstElement3.length > 0,
                    frm3HLElements: firstElement3,
                    hasFrm4HLElements: firstElement4.length > 0,
                    frm4HLElements: firstElement4
                });
            }
        }
        
        // Close browser after processing all creatives
        try {
            await browser.close();
        } catch (closeError) {
            console.log('Error closing browser (non-critical):', closeError.message);
        }
        
        console.log(`\n=== FINAL SUMMARY ===`);
        console.log(`Total creatives: ${creatives.length}`);
        creatives.forEach((c, i) => {
            console.log(`Creative ${i + 1}: "${c.title}"`);
            console.log(`  frm1_HL_: ${c.hasFrm1HLElements ? '✅' : '❌'} (${c.frm1HLElements.length} elements)`);
            if (c.frm1HLElements.length > 0) {
                console.log(`    Element IDs:`, c.frm1HLElements.map(el => el.id));
            }
            console.log(`  frm2_HL_: ${c.hasFrm2HLElements ? '✅' : '❌'} (${c.frm2HLElements.length} elements)`);
            if (c.frm2HLElements.length > 0) {
                console.log(`    Element IDs:`, c.frm2HLElements.map(el => el.id));
            }
            console.log(`  frm3_HL_: ${c.hasFrm3HLElements ? '✅' : '❌'} (${c.frm3HLElements.length} elements)`);
            if (c.frm3HLElements.length > 0) {
                console.log(`    Element IDs:`, c.frm3HLElements.map(el => el.id));
            }
            console.log(`  frm4_HL_: ${c.hasFrm4HLElements ? '✅' : '❌'} (${c.frm4HLElements.length} elements)`);
            if (c.frm4HLElements.length > 0) {
                console.log(`    Element IDs:`, c.frm4HLElements.map(el => el.id));
            }
        });
        console.log(`====================\n`);

        // Build response safely
        try {
            const response = { 
                ok: true, 
                count: jvxFrames.length, 
                frames: jvxFrames || [], 
                hlMatches: hlMatches || [],
                frameStructure: frameStructure || [],
                previewVariationTitles: previewVariationTitles || [],
                creatives: creatives || [],
                summary: {
                    totalFrames: frames.length,
                    jvxBaseFrames: jvxFrames.length,
                    framesWithHLMatches: hlMatches.length,
                    totalHLMatches: hlMatches.reduce((sum, h) => sum + (h.matches ? h.matches.length : 0), 0),
                    totalCreatives: creatives.length
                }
            };
            
            console.log('Sending successful response with:', {
                jvxBaseCount: jvxFrames.length,
                hlMatchesCount: hlMatches.length,
                frameStructureCount: frameStructure.length
            });
            
            return res.json(response);
        } catch (responseError) {
            console.error('Error building response:', responseError);
            throw new Error(`Error building response: ${responseError.message}`);
        }

    } catch (err) {
        console.error('Error in /api/count:', err);
        console.error('Error stack:', err.stack);
        console.error('Error details:', {
            message: err.message,
            name: err.name,
            code: err.code
        });
        
        if (browser) {
            try { 
                await browser.close(); 
            } catch (e) {
                console.error('Error closing browser:', e);
            }
        }
        
        // Provide more helpful error messages
        let errorMessage = err.message || 'Unknown error';
        if (errorMessage.includes('ECONNRESET')) {
            errorMessage = 'Connection was reset by the server. The website may be blocking automated requests, the URL may be unreachable, or there may be network issues. Please try again or check if the URL is accessible.';
        } else if (errorMessage.includes('ECONNREFUSED')) {
            errorMessage = 'Connection refused. The server may be down or the URL may be incorrect.';
        } else if (errorMessage.includes('ETIMEDOUT') || errorMessage.includes('timeout')) {
            errorMessage = 'Request timed out. The website took too long to respond. Please try again.';
        } else if (errorMessage.includes('net::ERR')) {
            errorMessage = `Network error: ${errorMessage}. Please check if the URL is correct and accessible.`;
        } else if (errorMessage.includes('Protocol error') || errorMessage.includes('Target closed')) {
            errorMessage = 'Browser connection lost. This may happen if the page takes too long to load. Please try again.';
        } else if (errorMessage.includes('Navigation failed')) {
            errorMessage = `Navigation failed: ${errorMessage}. The page may be blocking automated access or may require authentication.`;
        }
        
        // Log the full error for debugging
        console.error('Sending error response:', errorMessage);
        
        return res.status(500).json({ ok: false, error: errorMessage });
    }
});

// Get all sheets from a spreadsheet
app.get('/api/get-sheets/:spreadsheetId', async (req, res) => {
    try {
        const { spreadsheetId } = req.params;
        
        if (!spreadsheetId) {
            return res.status(400).json({ error: 'Missing spreadsheetId' });
        }

        // Get authentication
        const auth = serviceAccountAuth || oauth2Client;
        if (!auth) {
            return res.status(401).json({ 
                error: 'Authentication not configured. Please set up Google credentials.' 
            });
        }

        const sheets = google.sheets({ version: 'v4', auth });

        // Get spreadsheet metadata
        const response = await sheets.spreadsheets.get({
            spreadsheetId: spreadsheetId,
            fields: 'sheets.properties'
        });

        const sheetList = response.data.sheets.map(sheet => ({
            sheetId: sheet.properties.sheetId,
            title: sheet.properties.title,
            index: sheet.properties.index
        }));

        res.json({ 
            success: true, 
            sheets: sheetList 
        });

    } catch (error) {
        console.error('Error getting sheets:', error);
        res.status(500).json({ 
            error: error.message || 'Failed to get sheets',
            details: error.response?.data || error
        });
    }
});

// OAuth2 authorization endpoint
app.get('/auth', (req, res) => {
    if (!oauth2Client) {
        return res.status(500).json({ error: 'OAuth2 not configured' });
    }

    const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
    const authUrl = oauth2Client.generateAuthUrl({
        access_type: 'offline',
        scope: SCOPES,
        prompt: 'consent'
    });

    res.redirect(authUrl);
});

// OAuth2 callback
app.get('/oauth2callback', async (req, res) => {
    try {
        const { code } = req.query;
        if (!code) {
            return res.status(400).send('No authorization code provided');
        }

        const { tokens } = await oauth2Client.getToken(code);
        oauth2Client.setCredentials(tokens);

        res.send(`
            <html>
                <body>
                    <h1>✅ Authorization Successful!</h1>
                    <p>You can now close this window and return to the application.</p>
                    <script>setTimeout(() => window.close(), 2000);</script>
                </body>
            </html>
        `);
    } catch (error) {
        res.status(500).send(`Error: ${error.message}`);
    }
});

// Serve static files (CSS, JS, images, etc.) - but only if no API route matched
app.use(express.static(__dirname));

// Serve the HTML file (must be last, after all API routes)
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'Index.html'));
});

app.listen(PORT, () => {
    console.log(`🚀 Server running at http://localhost:${PORT}`);
    console.log(`📝 Open http://localhost:${PORT} in your browser`);
    
    if (!serviceAccountAuth && !oauth2Client) {
        console.log('\n⚠️  SETUP REQUIRED:');
        console.log('Option 1 - Service Account (Recommended):');
        console.log('  1. Create a service account in Google Cloud Console');
        console.log('  2. Download the JSON key file');
        console.log('  3. Save it as "service-account.json" in this directory');
        console.log('  4. Share your Google Sheet with the service account email');
        console.log('\nOption 2 - OAuth2:');
        console.log('  1. Set GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET environment variables');
        console.log('  2. Visit http://localhost:3000/auth to authorize');
    }
});

