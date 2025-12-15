const { firefox } = require('playwright');

class FordScraper {
  constructor() {
    this.browser = null;
    this.page = null;
    this.baseUrl = 'https://www.ford.com/support/recalls-details/';
  }

    async initialize() {
        try {
            console.log('Launching Ford scraper with Firefox...');
            this.browser = await firefox.launch({
                headless: true, // Browser runs in background
                args: [
                    '--no-sandbox', 
                    '--disable-setuid-sandbox'
                ]
            });
            
            console.log('Creating new page...');
            this.page = await this.browser.newPage({
                userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/115.0'
            });
            
            // Set viewport
            await this.page.setViewportSize({ width: 1920, height: 1080 });
            
            console.log('Ford scraper initialized successfully - Firefox browser running in headless mode');
            console.log('Browser is ready and waiting for VIN scraping to begin...');
            console.log('Current page URL:', await this.page.url());
            return true;
        } catch (error) {
            console.error('Error initializing Ford scraper:', error);
            return false;
        }
    }

    async extractRecallNumberFromPanel() {
        try {
            // Wait a bit for panel to fully load
            await this.page.waitForTimeout(1500);
            
            // Method 1: Look for recall number in the opened panel using the recall-info-piece-data class
            // This is the most reliable method - the recall number appears in this element when panel is open
            const recallDataElements = await this.page.$$('.recall-info-piece-data');
            
            for (const element of recallDataElements) {
                try {
                    // Check if element is visible (panel is open)
                    const isVisible = await element.isVisible();
                    if (!isVisible) continue;
                    
                    const text = await element.textContent();
                    if (text && text.trim()) {
                        const recallPattern = /^\d{2}[SC]\d{2}\/\d{2}V\d+$/;
                        if (recallPattern.test(text.trim())) {
                            const recallNumber = text.trim().split('/')[0];
                            return recallNumber;
                        }
                    }
                } catch (e) {
                    continue;
                }
            }
            
            // Method 2: Look for recall number in expanded button's parent/sibling content
            // Find the currently expanded button and look for recall number nearby
            try {
                const expandedButton = await this.page.$('button.recalls-info-button[aria-expanded="true"]');
                if (expandedButton) {
                    // Get the parent container's text content
                    const expandedContentText = await expandedButton.evaluate((el) => {
                        // Find the parent container
                        let parent = el.closest('[class*="recall"], [class*="panel"], [class*="accordion"]');
                        if (!parent) parent = el.parentElement;
                        // Go up a few levels to find the content container
                        while (parent && !parent.textContent.includes('/')) {
                            parent = parent.parentElement;
                        }
                        return parent ? parent.textContent : '';
                    });
                    
                    if (expandedContentText) {
                        const recallPattern = /\b(\d{2}[SC]\d{2})\/\d{2}V\d+\b/;
                        const match = expandedContentText.match(recallPattern);
                        if (match) {
                            return match[1];
                        }
                    }
                }
            } catch (e) {
                // Continue to next method
            }
            
            // Method 3: Fallback - Search visible page content for recall number pattern
            // This is less precise but should catch the recall number if other methods fail
            const pageText = await this.page.textContent('body');
            if (pageText) {
                const globalRecallPattern = /\b(\d{2}[SC]\d{2})\/\d{2}V\d+\b/g;
                const matches = [...pageText.matchAll(globalRecallPattern)];
                
                // Return the first match found
                if (matches.length > 0) {
                    return matches[0][1];
                }
            }
            
            return null;
        } catch (e) {
            console.log('Error extracting recall number from panel:', e.message);
            return null;
        }
    }

    async scrapeVinRecallData(vinNumber, retryCount = 0) {
        const MAX_RETRIES = 2; // Maximum number of retries after timeout
        
        try {
            if (!this.page) {
                throw new Error('Scraper not initialized. Call initialize() first.');
            }

            console.log(`\n=== STARTING FORD SCRAPING FOR VIN: ${vinNumber} ===`);
            if (retryCount > 0) {
                console.log(`🔄 Retry attempt ${retryCount} for VIN: ${vinNumber}`);
            }
            console.log(`Current page URL before navigation: ${await this.page.url()}`);
            console.log(`Scraping Ford recall data for VIN: ${vinNumber}`);

            // Navigate to Ford's recall lookup page
            console.log(`Navigating to: ${this.baseUrl}`);
            await this.page.goto(this.baseUrl, { 
                waitUntil: 'networkidle',
                timeout: 30000
            });

            console.log('Page loaded, waiting for elements...');
            await this.page.waitForTimeout(2000);

            // Find VIN input field using data-testid (most reliable selector)
            console.log('Waiting for VIN input field...');
            const vinInput = await this.page.waitForSelector('[data-testid="vin-search-text-field"]', {
                state: 'visible',
                timeout: 10000
            }).catch(async () => {
                // Fallback: try alternative selectors
                console.log('Primary selector failed, trying fallback selectors...');
                const fallbackSelectors = [
                    'input[data-testid*="vin"]',
                    'input[id*="vin"]',
                    'input[aria-labelledby*="vin"]'
                ];
                
                for (const selector of fallbackSelectors) {
                    try {
                        const element = await this.page.$(selector);
                        if (element && await element.isVisible()) {
                            console.log(`Found VIN input using fallback: ${selector}`);
                            return element;
                        }
                    } catch (e) {
                        continue;
                    }
                }
                return null;
            });

            if (!vinInput) {
                throw new Error('Could not find VIN input field');
            }

            console.log('Found VIN input field');

            // Use locator to avoid stale element references
            const vinInputLocator = this.page.locator('[data-testid="vin-search-text-field"]');
            
            // Clear the field first by selecting all and deleting
            await vinInputLocator.click({ clickCount: 3 });
            await this.page.keyboard.press('Backspace');
            await this.page.waitForTimeout(200);

            // Type VIN character by character (instead of fill/copy-paste)
            console.log(`Typing VIN: ${vinNumber}`);
            await vinInputLocator.type(vinNumber, { delay: 50 }); // Small delay between characters
            console.log(`VIN entered: ${vinNumber}`);

            // Wait a moment for validation and button to become enabled
            await this.page.waitForTimeout(1000);

            // Check for VIN validation error
            const errorElement = await this.page.$('[data-testid="vin-search-text-field-error"]');
            if (errorElement && await errorElement.isVisible()) {
                const errorText = await errorElement.textContent();
                console.log(`❌ Invalid VIN detected: ${errorText}`);
                return {
                    vin: vinNumber,
                    success: false,
                    error: `Invalid VIN: ${errorText?.trim() || 'Enter a valid 17-character Ford VIN'}`,
                    scrapedAt: new Date().toISOString()
                };
            }

            // Find and click the Search button
            // The button has aria-label="Search" and is initially disabled
            const submitButton = await this.page.waitForSelector('button[aria-label="Search"]:not([disabled])', {
                state: 'visible',
                timeout: 5000
            }).catch(async () => {
                // Fallback: try to find any enabled search button
                const buttons = await this.page.$$('button');
                for (const button of buttons) {
                    try {
                        const ariaLabel = await button.getAttribute('aria-label');
                        const disabled = await button.getAttribute('disabled');
                        const text = await button.textContent();
                        if ((ariaLabel === 'Search' || text?.includes('Search')) && !disabled) {
                            return button;
                        }
                    } catch (e) {
                        continue;
                    }
                }
                return null;
            });

            if (!submitButton) {
                // Check again for error in case button didn't enable due to invalid VIN
                const errorElement = await this.page.$('[data-testid="vin-search-text-field-error"]');
                if (errorElement && await errorElement.isVisible()) {
                    const errorText = await errorElement.textContent();
                    console.log(`❌ Invalid VIN detected: ${errorText}`);
                    return {
                        vin: vinNumber,
                        success: false,
                        error: `Invalid VIN: ${errorText?.trim() || 'Enter a valid 17-character Ford VIN'}`,
                        scrapedAt: new Date().toISOString()
                    };
                }
                throw new Error('Could not find Search button or button is still disabled');
            }

            await submitButton.click();
            console.log('Search button clicked');

            // Wait a moment and check for any validation errors that might appear after submission
            await this.page.waitForTimeout(1000);
            const errorAfterSubmit = await this.page.$('[data-testid="vin-search-text-field-error"]');
            if (errorAfterSubmit && await errorAfterSubmit.isVisible()) {
                const errorText = await errorAfterSubmit.textContent();
                console.log(`❌ Invalid VIN detected after submission: ${errorText}`);
                return {
                    vin: vinNumber,
                    success: false,
                    error: `Invalid VIN: ${errorText?.trim() || 'Enter a valid 17-character Ford VIN'}`,
                    scrapedAt: new Date().toISOString()
                };
            }

            // Wait for results to load
            await this.page.waitForTimeout(4000);

            // Extract recall information
            const recallData = await this.extractRecallData();

            return {
                vin: vinNumber,
                success: true,
                recallData: recallData,
                scrapedAt: new Date().toISOString()
            };

        } catch (error) {
            // Check if it's a timeout error and we haven't exceeded max retries
            const isTimeoutError = error.name === 'TimeoutError' || 
                                   error.message.includes('Timeout') || 
                                   error.message.includes('timeout');
            
            if (isTimeoutError && retryCount < MAX_RETRIES) {
                console.log(`\n⏱️  Timeout error detected for VIN ${vinNumber}. Restarting browser and retrying...`);
                console.log(`   Retry attempt: ${retryCount + 1} of ${MAX_RETRIES}`);
                
                try {
                    // Close the browser
                    console.log('Closing browser due to timeout...');
                    await this.close();
                    
                    // Wait a moment before restarting
                    await new Promise(resolve => setTimeout(resolve, 2000));
                    
                    // Reinitialize the browser
                    console.log('Reinitializing browser...');
                    const reinitialized = await this.initialize();
                    
                    if (!reinitialized) {
                        console.error('❌ Failed to reinitialize browser after timeout');
                        return {
                            vin: vinNumber,
                            success: false,
                            error: 'Failed to reinitialize browser after timeout',
                            scrapedAt: new Date().toISOString()
                        };
                    }
                    
                    console.log('✅ Browser restarted successfully. Retrying VIN scraping...');
                    
                    // Retry the scraping with incremented retry count
                    return await this.scrapeVinRecallData(vinNumber, retryCount + 1);
                    
                } catch (retryError) {
                    console.error(`❌ Error during browser restart/retry for VIN ${vinNumber}:`, retryError);
                    return {
                        vin: vinNumber,
                        success: false,
                        error: `Timeout and retry failed: ${retryError.message}`,
                        scrapedAt: new Date().toISOString()
                    };
                }
            }
            
            // If not a timeout error, or max retries exceeded, return error
            console.error(`Error scraping Ford data for VIN ${vinNumber}:`, error);
            return {
                vin: vinNumber,
                success: false,
                error: error.message,
                scrapedAt: new Date().toISOString()
            };
        }
    }

    async extractCustomerSatisfactionPrograms(uniqueRecalls, recalls) {
        try {
            console.log('Checking for Customer Satisfaction Programs section...');
            
            // Look for the "Customer Satisfaction Programs" section header
            const cspHeader = await this.page.$('[data-testid="button-csp-section-header"]');
            
            if (!cspHeader) {
                console.log('No Customer Satisfaction Programs section found');
                return;
            }
            
            console.log('Found Customer Satisfaction Programs section');
            
            // Extract campaign numbers directly from button data-testid attributes
            // CSP buttons have data-testid="button-{CAMPAIGN_NUMBER}" format
            // e.g., button-22L05, button-24N08, button-25N09
            const allButtons = await this.page.$$('button[data-testid^="button-"]');
            const cspButtons = [];
            
            for (const button of allButtons) {
                try {
                    const testId = await button.getAttribute('data-testid');
                    // Exclude the section header and recall buttons
                    if (testId && 
                        testId !== 'button-csp-section-header' && 
                        testId !== 'button-safety-recalls-section-header') {
                        const potentialCampaignNumber = testId.replace('button-', '');
                        // Campaign number pattern: 2 digits, letter, 2 digits (e.g., 22L05, 24N08, 25N09)
                        const campaignPattern = /^\d{2}[A-Za-z]\d{2}$/;
                        if (campaignPattern.test(potentialCampaignNumber)) {
                            cspButtons.push(button);
                        }
                    }
                } catch (e) {
                    continue;
                }
            }
            
            // Filter out recall buttons (they have YYV###### format, not YYL## format)
            const recallPattern = /^\d{2}V\d{3,6}$/;
            const filteredCspButtons = [];
            
            for (const button of cspButtons) {
                const testId = await button.getAttribute('data-testid');
                if (testId) {
                    const number = testId.replace('button-', '');
                    // If it matches recall pattern, skip it (it's a recall, not CSP)
                    if (!recallPattern.test(number)) {
                        filteredCspButtons.push(button);
                    }
                }
            }
            
            console.log(`Found ${filteredCspButtons.length} Customer Satisfaction Program button(s)`);
            
            // Extract campaign numbers directly from data-testid attributes
            for (let i = 0; i < filteredCspButtons.length; i++) {
                try {
                    const button = filteredCspButtons[i];
                    const testId = await button.getAttribute('data-testid');
                    
                    if (testId && testId.startsWith('button-')) {
                        // Extract campaign number from data-testid (format: button-22L05)
                        const campaignNumber = testId.replace('button-', '').toUpperCase();
                        
                        // Validate it's a campaign number format (2 digits, letter, 2 digits)
                        const campaignPattern = /^\d{2}[A-Za-z]\d{2}$/;
                        
                        if (campaignPattern.test(campaignNumber)) {
                            if (!uniqueRecalls.has(campaignNumber)) {
                                uniqueRecalls.add(campaignNumber);
                                console.log(`Found CSP campaign number: ${campaignNumber}`);
                                recalls.push({
                                    recallNumber: campaignNumber,
                                    element: 'customer-satisfaction-program',
                                    fullRecallNumber: campaignNumber,
                                    type: 'Safety'
                                });
                            } else {
                                console.log(`Campaign number ${campaignNumber} already found, skipping duplicate`);
                            }
                        }
                    }
                } catch (e) {
                    console.log(`Error processing CSP button ${i + 1}:`, e.message);
                    continue;
                }
            }
            
        } catch (error) {
            console.log('Error extracting Customer Satisfaction Programs:', error.message);
        }
    }

    async extractCampaignNumberFromPanel() {
        try {
            // Wait a bit for panel to fully load
            await this.page.waitForTimeout(1500);
            
            // Look for campaign number in the opened panel using the recall-info-piece-data class
            // Campaign numbers are like "22L05" (2 digits, letter, 2 digits)
            const recallDataElements = await this.page.$$('.recall-info-piece-data');
            
            for (const element of recallDataElements) {
                try {
                    const isVisible = await element.isVisible();
                    if (!isVisible) continue;
                    
                    const text = await element.textContent();
                    if (text && text.trim()) {
                        // Campaign number pattern: 2 digits, letter, 2 digits (e.g., 22L05, 24N08)
                        const campaignPattern = /^\d{2}[A-Za-z]\d{2}$/i;
                        if (campaignPattern.test(text.trim())) {
                            return text.trim().toUpperCase(); // Normalize to uppercase
                        }
                    }
                } catch (e) {
                    continue;
                }
            }
            
            return null;
        } catch (e) {
            console.log('Error extracting campaign number from panel:', e.message);
            return null;
        }
    }

    async extractRecallData() {
        try {
            // Wait a bit for the page to fully load after search
            await this.page.waitForTimeout(3000);
            
            // Check for the "no recall information" message
            const noRecallMessage = await this.page.$('text="We are not able to retrieve recall information for the VIN you have entered"');
            if (noRecallMessage) {
                console.log('Found "no recall information" message');
                return {
                    recalls: [{
                        recallNumber: 'No recall information',
                        element: 'no-recall-message',
                        type: 'Recall'
                    }],
                    recallCount: 0,
                    hasRecalls: false
                };
            }

            // Use a Set to store unique recall numbers
            const uniqueRecalls = new Set();
            const recalls = [];
            
            // Method 1: Extract recall numbers directly from button data-testid attributes
            // Ford recall buttons have data-testid="button-{RECALL_NUMBER}" format
            // e.g., button-24V684, button-25V019, button-09V399000
            try {
                // Wait for buttons to be available
                await this.page.waitForTimeout(2000);
                
                // Find all buttons with data-testid starting with "button-"
                // Filter out the section header button and only get recall buttons
                const allButtons = await this.page.$$('button[data-testid^="button-"]');
                const recallButtons = [];
                
                for (const button of allButtons) {
                    const testId = await button.getAttribute('data-testid');
                    // Exclude the section header and only include buttons that match recall number pattern
                    if (testId && testId !== 'button-safety-recalls-section-header') {
                        const potentialRecallNumber = testId.replace('button-', '');
                        // Validate it looks like a recall number (YYV###### format)
                        const recallPattern = /^\d{2}V\d{3,6}$/;
                        if (recallPattern.test(potentialRecallNumber)) {
                            recallButtons.push(button);
                        }
                    }
                }
                
                console.log(`Found ${recallButtons.length} recall button(s)`);
                
                if (recallButtons.length >= 1) {
                    // Extract recall numbers directly from data-testid attributes
                    for (let i = 0; i < recallButtons.length; i++) {
                        try {
                            const button = recallButtons[i];
                            const testId = await button.getAttribute('data-testid');
                            
                            if (testId && testId.startsWith('button-')) {
                                // Extract recall number from data-testid (format: button-24V684)
                                const recallNumber = testId.replace('button-', '');
                                
                                // Validate it's a recall number format (YYV###### or similar NHTSA format)
                                // Pattern: 2 digits, V, followed by 3-6 digits
                                const recallPattern = /^\d{2}V\d{3,6}$/;
                                
                                if (recallPattern.test(recallNumber)) {
                                    if (!uniqueRecalls.has(recallNumber)) {
                                        uniqueRecalls.add(recallNumber);
                                        console.log(`Found recall number: ${recallNumber}`);
                                        recalls.push({
                                            recallNumber: recallNumber,
                                            element: 'recall-button-data-testid',
                                            fullRecallNumber: recallNumber,
                                            type: 'Recall'
                                        });
                                    } else {
                                        console.log(`Recall number ${recallNumber} already found, skipping duplicate`);
                                    }
                                } else {
                                    console.log(`Skipping button with testid "${testId}" - doesn't match recall number pattern`);
                                }
                            }
                        } catch (e) {
                            console.log(`Error processing recall button ${i + 1}:`, e.message);
                            continue;
                        }
                    }
                }
            } catch (e) {
                console.log('Error finding/clicking recall buttons:', e);
            }
            
            // Method 2: Fallback - Extract from button IDs and aria-controls attributes
            if (recalls.length === 0) {
                try {
                    // Look for buttons with IDs matching the pattern button-{RECALL_NUMBER}
                    const buttonsWithIds = await this.page.$$('button[id^="button-"]');
                    
                    for (const button of buttonsWithIds) {
                        try {
                            const buttonId = await button.getAttribute('id');
                            if (buttonId && buttonId.startsWith('button-') && buttonId !== 'button-safety-recalls-section-header') {
                                const recallNumber = buttonId.replace('button-', '');
                                // Validate NHTSA recall format: YYV######
                                const recallPattern = /^\d{2}V\d{3,6}$/;
                                
                                if (recallPattern.test(recallNumber) && !uniqueRecalls.has(recallNumber)) {
                                    uniqueRecalls.add(recallNumber);
                                    console.log(`Found recall number from button ID: ${recallNumber}`);
                                    recalls.push({
                                        recallNumber: recallNumber,
                                        element: 'recall-button-id',
                                        fullRecallNumber: recallNumber,
                                        type: 'Recall'
                                    });
                                }
                            }
                        } catch (e) {
                            continue;
                        }
                    }
                } catch (e) {
                    console.log('Error extracting from button IDs:', e);
                }
            }
            
            // Method 3: Fallback - Search entire page content for NHTSA recall number patterns
            if (recalls.length === 0) {
                try {
                    const pageContent = await this.page.textContent('body');
                    if (pageContent) {
                        // Pattern to find NHTSA recall numbers: YYV###### (2 digits, V, 3-6 digits)
                        const globalRecallPattern = /\b(\d{2}V\d{3,6})\b/g;
                        let match;
                        
                        while ((match = globalRecallPattern.exec(pageContent)) !== null) {
                            const recallNumber = match[1];
                            if (!uniqueRecalls.has(recallNumber)) {
                                uniqueRecalls.add(recallNumber);
                                console.log(`Found recall number (page scan): ${recallNumber}`);
                                recalls.push({
                                    recallNumber: recallNumber,
                                    element: 'page-content-scan',
                                    fullRecallNumber: recallNumber,
                                    type: 'Recall'
                                });
                            }
                        }
                    }
                } catch (e) {
                    console.log('Error scanning page content for recalls:', e);
                }
            }
            
            // Extract Customer Satisfaction Programs (campaign numbers)
            await this.extractCustomerSatisfactionPrograms(uniqueRecalls, recalls);
            
            // Log summary of found recalls
            if (recalls.length > 0) {
                const validRecalls = recalls.filter(r => 
                    r.recallNumber && 
                    r.recallNumber !== 'No recall information' && 
                    r.recallNumber !== 'No recall information available'
                );
                console.log(`📋 Total unique recalls found: ${validRecalls.length}`);
                if (validRecalls.length > 1) {
                    console.log(`   Multiple recalls detected: ${validRecalls.map(r => r.recallNumber).join(', ')}`);
                }
            }

            // If no specific recall elements found, check page content for other indicators
            if (recalls.length === 0) {
                try {
                    const pageContent = await this.page.textContent('body');
                    if (pageContent) {
                        // Check for the specific "no recall" message in page content
                        if (pageContent.includes('We are not able to retrieve recall information for the VIN you have entered')) {
                            recalls.push({
                                recallNumber: 'No recall information available',
                                element: 'no-recall-message',
                                type: 'Recall'
                            });
                        } else {
                            // Look for other recall-related content
                            const recallKeywords = ['recall', 'safety', 'defect', 'campaign', 'service bulletin'];
                            const hasRecallInfo = recallKeywords.some(keyword => 
                                pageContent.toLowerCase().includes(keyword)
                            );

                            if (hasRecallInfo) {
                                recalls.push({
                                    recallNumber: 'Recall information may be available but could not be extracted',
                                    element: 'general-content',
                                    type: 'Recall'
                                });
                            } else {
                                recalls.push({
                                    recallNumber: 'No recall information available',
                                    element: 'no-recall-detected',
                                    type: 'Recall'
                                });
                            }
                        }
                    }
                } catch (e) {
                    console.log('Could not extract general page content');
                    recalls.push({
                        recallNumber: 'No recall information available',
                        element: 'extraction-error',
                        type: 'Recall'
                    });
                }
            }

            return {
                recalls: recalls,
                recallCount: recalls.length,
                hasRecalls: recalls.length > 0 && !recalls.some(r => r.recallNumber.includes('No recall information available'))
            };

        } catch (error) {
            console.error('Error extracting recall data:', error);
            return {
                recalls: [{
                    recallNumber: 'No recall information available',
                    element: 'extraction-error',
                    type: 'Recall'
                }],
                recallCount: 0,
                hasRecalls: false,
                error: error.message
            };
        }
    }

    async scrapeMultipleVins(vinNumbers) {
        const results = [];
        
        for (const vin of vinNumbers) {
            try {
                const result = await this.scrapeVinRecallData(vin);
                results.push(result);
                
                // Add delay between requests to be respectful
                await this.page.waitForTimeout(2000);
                
            } catch (error) {
                results.push({
                    vin: vin,
                    success: false,
                    error: error.message,
                    scrapedAt: new Date().toISOString()
                });
            }
        }
        
        return results;
    }

    async close() {
        try {
            if (this.browser) {
                await this.browser.close();
                this.browser = null;
                this.page = null;
                console.log('Ford scraper closed successfully');
            }
        } catch (error) {
            console.error('Error closing Ford scraper:', error);
            // Reset references even if close fails
            this.browser = null;
            this.page = null;
        }
    }
}

module.exports = FordScraper;
