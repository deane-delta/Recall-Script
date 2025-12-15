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

    async scrapeVinRecallData(vinNumber) {
        try {
            if (!this.page) {
                throw new Error('Scraper not initialized. Call initialize() first.');
            }

            console.log(`\n=== STARTING FORD SCRAPING FOR VIN: ${vinNumber} ===`);
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

            // Find VIN input field
            const vinInput = await this.page.$('#vin-field-vin-selector');
            if (!vinInput) {
                throw new Error('Could not find VIN input field');
            }

            console.log('Found VIN input field');

            // Enter VIN
            await vinInput.click();
            await vinInput.fill(''); // Clear field
            await vinInput.fill(vinNumber);
            console.log(`VIN entered: ${vinNumber}`);

            // Find and click the Search button
            const submitButton = await this.page.$('[data-test-id="vin-submit-button"]');
            if (!submitButton) {
                throw new Error('Could not find Search button');
            }

            await submitButton.click();
            console.log('Search button clicked');

            // Wait for results to load
            await this.page.waitForTimeout(5000);

            // Extract recall information
            const recallData = await this.extractRecallData();

            return {
                vin: vinNumber,
                success: true,
                recallData: recallData,
                scrapedAt: new Date().toISOString()
            };

        } catch (error) {
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
            
            // Look for the "Customer Satisfaction Programs" header
            const cspHeader = await this.page.$('header.recalls-safety-heading:has-text("Customer Satisfaction Programs")');
            
            if (!cspHeader) {
                console.log('No Customer Satisfaction Programs section found');
                return;
            }
            
            console.log('Found Customer Satisfaction Programs section');
            
            // Get the position of the CSP header and find the next section header
            const sectionInfo = await this.page.evaluate(() => {
                const headers = Array.from(document.querySelectorAll('header.recalls-safety-heading'));
                const cspHeader = headers.find(h => h.textContent.includes('Customer Satisfaction Programs'));
                if (!cspHeader) return null;
                
                const cspY = cspHeader.getBoundingClientRect().y;
                const nextHeader = headers.find(h => {
                    const hY = h.getBoundingClientRect().y;
                    return hY > cspY && h !== cspHeader;
                });
                
                return {
                    cspY: cspY,
                    nextHeaderY: nextHeader ? nextHeader.getBoundingClientRect().y : Infinity
                };
            });
            
            if (!sectionInfo) {
                console.log('Could not determine CSP section boundaries');
                return;
            }
            
            // Find all buttons and filter to only those in the CSP section
            // Re-find buttons to ensure we have fresh references (in case they were clicked during regular recall extraction)
            const allButtons = await this.page.$$('button.recalls-info-button');
            const freshCspButtons = [];
            
            for (const button of allButtons) {
                try {
                    const buttonY = await button.evaluate((btn) => btn.getBoundingClientRect().y);
                    if (buttonY > sectionInfo.cspY && buttonY < sectionInfo.nextHeaderY) {
                        freshCspButtons.push(button);
                    }
                } catch (e) {
                    continue;
                }
            }
            
            console.log(`Processing ${freshCspButtons.length} Customer Satisfaction Program button(s)`);
            
            for (let i = 0; i < freshCspButtons.length; i++) {
                try {
                    const button = freshCspButtons[i];
                    
                    // Scroll button into view
                    await button.scrollIntoViewIfNeeded();
                    await this.page.waitForTimeout(500);
                    
                    // Check if button is already expanded
                    const isExpanded = await button.getAttribute('aria-expanded');
                    
                    // If not expanded, click to open
                    if (isExpanded !== 'true') {
                        await button.click();
                        console.log(`Clicked CSP button ${i + 1} of ${freshCspButtons.length} to open`);
                        await this.page.waitForTimeout(2000);
                    } else {
                        console.log(`CSP button ${i + 1} already expanded, extracting campaign number`);
                    }
                    
                    // Extract campaign number from the opened panel
                    const campaignNumber = await this.extractCampaignNumberFromPanel();
                    
                    if (campaignNumber && !uniqueRecalls.has(campaignNumber)) {
                        uniqueRecalls.add(campaignNumber);
                        console.log(`Found CSP campaign number: ${campaignNumber}`);
                        recalls.push({
                            recallNumber: campaignNumber,
                            element: 'customer-satisfaction-program',
                            fullRecallNumber: campaignNumber,
                            type: 'Safety'
                        });
                    } else if (campaignNumber) {
                        console.log(`Campaign number ${campaignNumber} already found, skipping duplicate`);
                    }
                    
                    // Close the panel if needed (for next button)
                    if (i < freshCspButtons.length - 1) {
                        const stillExpanded = await button.getAttribute('aria-expanded');
                        if (stillExpanded === 'true') {
                            await button.click();
                            await this.page.waitForTimeout(500);
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
            
            // Method 1: Check for recall buttons and click each one to extract recall numbers
            try {
                // Wait for buttons to be available
                await this.page.waitForTimeout(2000);
                
                const recallButtons = await this.page.$$('button.recalls-info-button');
                console.log(`Found ${recallButtons.length} recall button(s)`);
                
                if (recallButtons.length >= 1) {
                    // Click each button to open recall details and extract recall numbers
                    for (let i = 0; i < recallButtons.length; i++) {
                        try {
                            // Wait a bit before clicking to ensure page is ready
                            await this.page.waitForTimeout(1000);
                            
                            // Get all buttons again (they might have changed after previous clicks)
                            const currentButtons = await this.page.$$('button.recalls-info-button');
                            if (i >= currentButtons.length) {
                                console.log(`Button index ${i} no longer available, skipping`);
                                continue;
                            }
                            
                            const button = currentButtons[i];
                            
                            // Scroll button into view
                            await button.scrollIntoViewIfNeeded();
                            await this.page.waitForTimeout(500);
                            
                            // Check if button is already expanded (panel already open)
                            const isExpanded = await button.getAttribute('aria-expanded');
                            
                            // Click the button to open/close recall details
                            await button.click();
                            console.log(`Clicked recall button ${i + 1} of ${recallButtons.length}`);
                            
                            // Wait for recall details panel to load
                            await this.page.waitForTimeout(2000);
                            
                            // Extract recall number from the opened panel
                            const recallNumber = await this.extractRecallNumberFromPanel();
                            
                            if (recallNumber && !uniqueRecalls.has(recallNumber)) {
                                uniqueRecalls.add(recallNumber);
                                console.log(`Found recall number: ${recallNumber}`);
                                recalls.push({
                                    recallNumber: recallNumber,
                                    element: 'recalls-info-button',
                                    fullRecallNumber: recallNumber,
                                    type: 'Recall'
                                });
                            } else if (recallNumber) {
                                console.log(`Recall number ${recallNumber} already found, skipping duplicate`);
                            }
                            
                            // If panel is open, close it by clicking again (if needed for next button)
                            // Check if we need to close it before moving to next button
                            if (i < recallButtons.length - 1) {
                                // Check if button is still expanded
                                const stillExpanded = await button.getAttribute('aria-expanded');
                                if (stillExpanded === 'true') {
                                    // Close the panel by clicking again
                                    await button.click();
                                    await this.page.waitForTimeout(500);
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
            
            // Method 2: Fallback - Look for recall information using the specific class
            if (recalls.length === 0) {
                const recallDataElements = await this.page.$$('.recall-info-piece-data');
                
                if (recallDataElements.length > 0) {
                    console.log(`Found ${recallDataElements.length} recall data elements`);
                    
                    // Look for recall numbers in the elements
                    for (const element of recallDataElements) {
                        try {
                            const text = await element.textContent();
                            if (text && text.trim().length > 0) {
                                // Check if this looks like a recall number (pattern: XXSXX/XXVXXXXXX or XXCXX/XXVXXX)
                                const recallPattern = /^\d{2}[SC]\d{2}\/\d{2}V\d+$/; // Pattern like 10S13/10V385000, 25C42/25V543
                                
                                if (recallPattern.test(text.trim())) {
                                    // Extract the first part before the slash (e.g., 10S13 from 10S13/10V385000)
                                    const recallNumber = text.trim().split('/')[0];
                                    if (!uniqueRecalls.has(recallNumber)) {
                                        uniqueRecalls.add(recallNumber);
                                        console.log(`Found recall number: ${recallNumber} from ${text.trim()}`);
                                        recalls.push({
                                            recallNumber: recallNumber,
                                            element: 'recall-info-piece-data',
                                            fullRecallNumber: text.trim(),
                                            type: 'Recall'
                                        });
                                    }
                                }
                            }
                        } catch (e) {
                            console.log('Error extracting text from recall element:', e);
                            continue;
                        }
                    }
                }
            }
            
            // Method 3: Fallback - Search entire page content for all recall number patterns
            if (recalls.length === 0) {
                try {
                    const pageContent = await this.page.textContent('body');
                    if (pageContent) {
                        // Pattern to find all recall numbers in the format XXSXX/XXVXXXXXX or XXCXX/XXVXXX
                        const globalRecallPattern = /\b(\d{2}[SC]\d{2})\/\d{2}V\d+\b/g;
                        let match;
                        
                        while ((match = globalRecallPattern.exec(pageContent)) !== null) {
                            const recallNumber = match[1]; // Extract the first part (e.g., 10S13)
                            if (!uniqueRecalls.has(recallNumber)) {
                                uniqueRecalls.add(recallNumber);
                                console.log(`Found recall number (page scan): ${recallNumber}`);
                                recalls.push({
                                    recallNumber: recallNumber,
                                    element: 'page-content-scan',
                                    fullRecallNumber: match[0], // Full match like 10S13/10V385000
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
                console.log('Ford scraper closed successfully');
            }
        } catch (error) {
            console.error('Error closing Ford scraper:', error);
        }
    }
}

module.exports = FordScraper;
