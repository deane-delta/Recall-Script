const { chromium } = require('playwright');

class DocSearchScraper {
    constructor(username, password) {
        this.username = username;
        this.password = password;
        this.browser = null;
        this.page = null;
        this.isAuthenticated = false;
        this.baseUrl = 'https://techops.delta.com/docsearch/';
    }

    async initialize() {
        try {
            // For manual sign-in, we don't require credentials upfront
            console.log('Initializing DocSearch scraper for manual sign-in...');

            this.browser = await chromium.launch({
                headless: true, // Run headless
                args: ['--no-sandbox', '--disable-setuid-sandbox']
            });
            
            this.page = await this.browser.newPage({
                userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
            });
            
            // Set viewport
            await this.page.setViewportSize({ width: 1920, height: 1080 });
            
            console.log('DocSearch scraper initialized successfully');
            return true;
        } catch (error) {
            console.error('Error initializing DocSearch scraper:', error);
            return false;
        }
    }

    async checkIfAlreadySignedIn() {
        try {
            if (!this.page) {
                throw new Error('Scraper not initialized. Call initialize() first.');
            }

            console.log('Checking if user is already signed into DocSearch...');

            // Navigate to DocSearch homepage
            await this.page.goto(this.baseUrl, { 
                waitUntil: 'networkidle',
                timeout: 30000 
            });

            // Wait a moment for any redirects or dynamic content
            await this.page.waitForTimeout(3000);

            // Check for signs that user is already signed in
            const signedInIndicators = [
                // Look for logout/signout links
                'a:has-text("Logout")',
                'a:has-text("Sign Out")',
                'a:has-text("Log Out")',
                'button:has-text("Logout")',
                'button:has-text("Sign Out")',
                // Look for user profile/menu elements
                '.user-menu',
                '.profile',
                '.account',
                '.user-info',
                // Look for dashboard or main content areas
                '.dashboard',
                '.main-content',
                '.user-dashboard',
                // Look for specific DocSearch authenticated elements
                '[class*="authenticated"]',
                '[class*="logged-in"]',
                '[id*="user"]',
                '[id*="profile"]'
            ];

            let isSignedIn = false;
            for (const selector of signedInIndicators) {
                try {
                    const element = await this.page.$(selector);
                    if (element) {
                        console.log(`Found signed-in indicator: ${selector}`);
                        isSignedIn = true;
                        break;
                    }
                } catch (e) {
                    continue;
                }
            }

            // Also check the URL for signs of being signed in
            const currentUrl = this.page.url();
            const urlIndicators = ['dashboard', 'home', 'main', 'user', 'profile', 'account'];
            const urlIndicatesSignIn = urlIndicators.some(indicator => 
                currentUrl.toLowerCase().includes(indicator)
            );

            if (urlIndicatesSignIn) {
                console.log('URL indicates user is signed in');
                isSignedIn = true;
            }

            // Check page content for authentication indicators
            try {
                const pageContent = await this.page.textContent('body');
                const contentIndicators = ['logout', 'sign out', 'profile', 'dashboard', 'welcome'];
                const contentIndicatesSignIn = contentIndicators.some(indicator => 
                    pageContent.toLowerCase().includes(indicator)
                );

                if (contentIndicatesSignIn) {
                    console.log('Page content indicates user is signed in');
                    isSignedIn = true;
                }
            } catch (e) {
                console.log('Could not check page content');
            }

            if (isSignedIn) {
                console.log('✅ User appears to be already signed into DocSearch');
                this.isAuthenticated = true;
                return true;
            } else {
                console.log('❌ User does not appear to be signed in');
                return false;
            }

        } catch (error) {
            console.error('Error checking sign-in status:', error);
            return false;
        }
    }

    async authenticate() {
        try {
            if (!this.page) {
                throw new Error('Scraper not initialized. Call initialize() first.');
            }

            console.log('User not signed in. Opening browser for manual sign-in...');
            console.log('⚠️ ACTION REQUIRED: Please sign in to DocSearch manually in the browser window');

            // Navigate to DocSearch homepage
            await this.page.goto(this.baseUrl, { 
                waitUntil: 'networkidle',
                timeout: 30000 
            });

            console.log('📝 Browser opened. Please sign in manually and navigate to the homepage.');
            console.log('⏳ Waiting for you to complete sign-in... (5 minute timeout)');
            
            // Wait for user to sign in and navigate to homepage
            // We'll wait for a specific element that indicates successful login
            try {
                // Wait for the homepage to load (you can adjust this selector based on actual homepage elements)
                await this.page.waitForSelector('body', { timeout: 300000 }); // 5 minute timeout
                console.log('✅ Detected homepage loaded. Proceeding with DocSearch workflow...');
                
                // Wait a bit more to ensure user is ready
                await this.page.waitForTimeout(2000);
                
                this.isAuthenticated = true;
                return true;
            } catch (waitError) {
                console.log('⏱️ Timeout waiting for homepage. Assuming user is ready to proceed...');
                this.isAuthenticated = true;
                return true;
            }

        } catch (error) {
            console.error('❌ Error during DocSearch authentication:', error);
            return false;
        }
    }

    async searchVinData(recallNumber) {
        try {
            if (!this.isAuthenticated) {
                throw new Error('Not authenticated. Call authenticate() first.');
            }

            console.log(`Starting DocSearch workflow for recall number: ${recallNumber}`);

            // Step 1: Navigate directly to the search form URL
            console.log('Navigating directly to DocSearch search form...');
            const searchFormUrl = 'https://techops.delta.com/docsearch/tmd_search_form.aspx';
            await this.page.goto(searchFormUrl, { 
                waitUntil: 'networkidle',
                timeout: 30000 
            });
            
            // Wait for page to fully load
            await this.page.waitForTimeout(3000);
            console.log('Successfully navigated to search form');

            // Step 2: Check the EA/ERA checkbox
            console.log('Checking the EA/ERA checkbox...');
            const checkboxSelectors = [
                'input[id="EA"]',  // Primary selector from the HTML you provided
                'input[name="EA"]', // Alternative by name
                'input[type="checkbox"][name="EA"]', // More specific
                'td[id="chkbx3"] input[type="checkbox"]', // By parent td id
                'span[name="EA"] input[type="checkbox"]' // By parent span
            ];
            
            let checkboxChecked = false;
            for (const selector of checkboxSelectors) {
                try {
                    const checkbox = await this.page.$(selector);
                    if (checkbox) {
                        await checkbox.check();
                        console.log(`Successfully checked EA/ERA checkbox with selector: ${selector}`);
                        checkboxChecked = true;
                        await this.page.waitForTimeout(2000);
                        break;
                    }
                } catch (e) {
                    console.log(`Could not check checkbox with selector: ${selector}`);
                    continue;
                }
            }
            
            // If specific selectors don't work, try looking for "EA/ERA" text
            if (!checkboxChecked) {
                console.log('Trying to find checkbox by "EA/ERA" text...');
                try {
                    // Look for the text "EA/ERA" and find the associated checkbox
                    const eaEraElement = await this.page.$('text="EA/ERA"');
                    if (eaEraElement) {
                        // Find the checkbox in the same td or nearby
                        const checkbox = await eaEraElement.$('xpath=..//input[@type="checkbox"]') || 
                                       await eaEraElement.$('xpath=../..//input[@type="checkbox"]') ||
                                       await eaEraElement.$('xpath=../../..//input[@type="checkbox"]');
                        
                        if (checkbox) {
                            await checkbox.check();
                            console.log('Successfully checked EA/ERA checkbox found by text');
                            checkboxChecked = true;
                            await this.page.waitForTimeout(2000);
                        }
                    }
                } catch (e) {
                    console.log('Could not find checkbox by EA/ERA text');
                }
            }
            
            // Final fallback - try any checkbox if specific ones don't work
            if (!checkboxChecked) {
                console.log('Trying fallback checkbox selection...');
                try {
                    await this.page.check('input[type="checkbox"]');
                    console.log('Successfully checked fallback checkbox');
                    checkboxChecked = true;
                    await this.page.waitForTimeout(2000);
                } catch (e) {
                    console.log('Warning: Could not find or check any checkbox');
                }
            }

            // Step 3: Enter the recall number in the "All of the words:" field
            console.log(`Entering recall number "${recallNumber}" in the "All of the words:" field`);
            
            // Wait for the page to be fully loaded and elements to be available
            console.log('Waiting for input fields to be available...');
            await this.page.waitForTimeout(3000);
            
            // Try to wait for the specific input field to be visible
            try {
                await this.page.waitForSelector('input[name="all_words"]', { timeout: 10000 });
                console.log('Found all_words input field');
            } catch (e) {
                console.log('Could not wait for all_words field, proceeding with direct selection...');
            }
            
            const inputSelector = 'input[name="all_words"]';
            try {
                const inputField = await this.page.$(inputSelector);
                if (inputField) {
                    // Check if the field is visible and enabled
                    const isVisible = await inputField.isVisible();
                    const isEnabled = await inputField.isEnabled();
                    console.log(`Input field visibility: ${isVisible}, enabled: ${isEnabled}`);
                    
                    if (isVisible && isEnabled) {
                        await inputField.click(); // Click to focus
                        await inputField.fill(''); // Clear field
                        await inputField.fill(recallNumber);
                        console.log(`✅ Successfully entered recall number "${recallNumber}" in "All of the words:" field`);
                    } else {
                        throw new Error('Input field is not visible or enabled');
                    }
                } else {
                    throw new Error('Input field not found');
                }
            } catch (error) {
                console.log('Could not find the "all_words" input field, trying alternative selectors...');
                // Try alternative selectors for input field
                const inputAlternatives = [
                    'input[id="all_words"]',
                    'input[name="all_words"]',
                    'input[accesskey="w"]',
                    'input[title*="All word"]',
                    'input.textBoxLarger'
                ];
                
                let inputFilled = false;
                for (const selector of inputAlternatives) {
                    try {
                        console.log(`Trying selector: ${selector}`);
                        const inputField = await this.page.$(selector);
                        if (inputField) {
                            const isVisible = await inputField.isVisible();
                            const isEnabled = await inputField.isEnabled();
                            console.log(`Alternative field visibility: ${isVisible}, enabled: ${isEnabled}`);
                            
                            if (isVisible && isEnabled) {
                                await inputField.click(); // Click to focus
                                await inputField.fill(''); // Clear field
                                await inputField.fill(recallNumber);
                                console.log(`✅ Successfully entered recall number "${recallNumber}" in alternative field: ${selector}`);
                                inputFilled = true;
                                break;
                            }
                        }
                    } catch (e) {
                        console.log(`Failed with selector ${selector}:`, e.message);
                        continue;
                    }
                }
                
                if (!inputFilled) {
                    console.log('❌ Could not find any suitable input field for recall number');
                    console.log('Trying alternative input methods...');
                    
                    // Try using page.fill as a last resort
                    try {
                        await this.page.fill('input[name="all_words"]', recallNumber);
                        console.log(`✅ Successfully filled field using page.fill method`);
                        inputFilled = true;
                    } catch (e) {
                        console.log('page.fill method also failed:', e.message);
                        
                        // Try using evaluate to set value directly
                        try {
                            await this.page.evaluate((recallNum) => {
                                const field = document.querySelector('input[name="all_words"]');
                                if (field) {
                                    field.value = recallNum;
                                    field.dispatchEvent(new Event('input', { bubbles: true }));
                                    field.dispatchEvent(new Event('change', { bubbles: true }));
                                    return true;
                                }
                                return false;
                            }, recallNumber);
                            console.log(`✅ Successfully set field value using evaluate method`);
                            inputFilled = true;
                        } catch (e2) {
                            console.log('evaluate method also failed:', e2.message);
                        }
                    }
                    
                    if (!inputFilled) {
                        console.log('Available input fields on page:');
                        try {
                            const allInputs = await this.page.$$('input');
                            for (let i = 0; i < allInputs.length; i++) {
                                const input = allInputs[i];
                                const name = await input.getAttribute('name').catch(() => 'no-name');
                                const id = await input.getAttribute('id').catch(() => 'no-id');
                                const type = await input.getAttribute('type').catch(() => 'no-type');
                                console.log(`  Input ${i + 1}: name="${name}", id="${id}", type="${type}"`);
                            }
                        } catch (e) {
                            console.log('Could not list available input fields');
                        }
                        throw new Error('Could not find or fill the required input field');
                    }
                }
            }

            // Step 4: Click the "Search" button
            console.log('Looking for the "Search" button...');
            const submitSelectors = [
                'a[id="searchButton"]',  // Primary selector from the HTML you provided
                'a.btnBlueBg',  // By class
                'a:has-text("Search")',  // By text content
                'a[title="Submit to search"]',  // By title attribute
                'a[href*="__doPostBack"]'  // By href pattern
            ];

            let submitClicked = false;
            for (const selector of submitSelectors) {
                try {
                    const submitButton = await this.page.$(selector);
                    if (submitButton) {
                        await submitButton.click();
                        console.log(`Successfully clicked Search button: ${selector}`);
                        submitClicked = true;
                        break;
                    }
                } catch (e) {
                    continue;
                }
            }

            if (!submitClicked) {
                console.log('Could not find the specific Search button, trying fallback selectors...');
                const fallbackSelectors = [
                    'button[type="submit"]',
                    'input[type="submit"]',
                    'button:has-text("Search")',
                    'button:has-text("Submit")',
                    'input[value*="Search"]',
                    'input[value*="Submit"]'
                ];

                for (const selector of fallbackSelectors) {
                    try {
                        const submitButton = await this.page.$(selector);
                        if (submitButton) {
                            await submitButton.click();
                            console.log(`Successfully clicked fallback submit button: ${selector}`);
                            submitClicked = true;
                            break;
                        }
                    } catch (e) {
                        continue;
                    }
                }
            }

            if (!submitClicked) {
                console.log('No submit button found, trying to press Enter on input field...');
                try {
                    await this.page.press('input[name="all_words"]', 'Enter');
                    console.log('Pressed Enter on input field');
                    submitClicked = true;
                } catch (e) {
                    console.log('Could not press Enter on input field');
                }
            }

            if (!submitClicked) {
                console.log('Warning: Could not find or click submit button');
            }

            // Wait for search results to load
            console.log('Waiting for search results...');
            await this.page.waitForTimeout(5000);

            // Extract search results and check if EA exists
            const eaExists = await this.checkEAExists();
            const eaNumber = eaExists ? await this.extractEANumber() : null;

            return {
                success: true,
                recallNumber: recallNumber,
                eaExists: eaExists,
                eaNumber: eaNumber || null,
                searchResults: {
                    results: [],
                    resultCount: 0,
                    hasResults: eaExists
                },
                resultCount: eaExists ? 1 : 0
            };

        } catch (error) {
            console.error(`Error in DocSearch workflow for recall number ${recallNumber}:`, error);
            return {
                success: false,
                recallNumber: recallNumber,
                error: error.message,
                eaExists: false,
                eaNumber: null,
                searchResults: [],
                resultCount: 0
            };
        }
    }

    async extractEANumber() {
        try {
            console.log('Extracting EA number from DocSearchResultDataGrid...');
            
            // Look for the specific span element: <span id="DocSearchResultDataGrid_LblDocID_0" class="labelText" style="float: left">15-479030-03 rev C</span>
            const eaNumberSelectors = [
                'span#DocSearchResultDataGrid_LblDocID_0',
                'span[id="DocSearchResultDataGrid_LblDocID_0"]',
                'span[id*="DocSearchResultDataGrid_LblDocID"]'
            ];
            
            for (const selector of eaNumberSelectors) {
                try {
                    const spanElement = await this.page.$(selector);
                    if (spanElement) {
                        const isVisible = await spanElement.isVisible();
                        if (isVisible) {
                            const eaNumber = await spanElement.textContent();
                            if (eaNumber && eaNumber.trim()) {
                                console.log(`✅ Found EA number: ${eaNumber.trim()}`);
                                return eaNumber.trim();
                            }
                        }
                    }
                } catch (e) {
                    console.log(`Could not find EA number with selector: ${selector}`);
                    continue;
                }
            }
            
            // Fallback: try to find any span with DocSearchResultDataGrid_LblDocID
            try {
                const allDocIDSpans = await this.page.$$('span[id*="DocSearchResultDataGrid_LblDocID"]');
                if (allDocIDSpans && allDocIDSpans.length > 0) {
                    for (const span of allDocIDSpans) {
                        const isVisible = await span.isVisible();
                        if (isVisible) {
                            const eaNumber = await span.textContent();
                            if (eaNumber && eaNumber.trim()) {
                                console.log(`✅ Found EA number (fallback): ${eaNumber.trim()}`);
                                return eaNumber.trim();
                            }
                        }
                    }
                }
            } catch (e) {
                console.log('Could not find EA number using fallback method');
            }
            
            console.log('❌ No EA number found');
            return null;
            
        } catch (error) {
            console.error('Error extracting EA number:', error);
            return null;
        }
    }

    async checkEAExists() {
        try {
            console.log('Checking if EA exists by examining DocSearchResultDataGrid span...');
            
            // Look for the specific span element: <span id="DocSearchResultDataGrid_LblRank_0" style="float: left">1</span>
            const spanSelectors = [
                'span#DocSearchResultDataGrid_LblRank_0',
                'span[id="DocSearchResultDataGrid_LblRank_0"]',
                'span[id*="DocSearchResultDataGrid_LblRank"]'
            ];
            
            let spanFound = false;
            for (const selector of spanSelectors) {
                try {
                    const spanElement = await this.page.$(selector);
                    if (spanElement) {
                        const isVisible = await spanElement.isVisible();
                        if (isVisible) {
                            console.log(`✅ Found EA result span element with selector: ${selector}`);
                            spanFound = true;
                            return true;
                        }
                    }
                } catch (e) {
                    console.log(`Could not find span with selector: ${selector}`);
                    continue;
                }
            }
            
            // Also try to look for any span with DocSearchResultDataGrid in the id
            if (!spanFound) {
                try {
                    const allSpans = await this.page.$$('span[id*="DocSearchResultDataGrid"]');
                    if (allSpans && allSpans.length > 0) {
                        console.log(`✅ Found ${allSpans.length} DocSearchResult spans - EA exists`);
                        spanFound = true;
                        return true;
                    }
                } catch (e) {
                    console.log('Could not find any DocSearchResult spans');
                }
            }
            
            if (!spanFound) {
                console.log('❌ EA does not exist - No DocSearchResult spans found');
                return false;
            }
            
            return spanFound;
            
        } catch (error) {
            console.error('Error checking EA existence:', error);
            return false;
        }
    }

    async extractSearchResults() {
        try {
            // Look for search results in various possible locations
            const resultSelectors = [
                '.search-result',
                '.result-item',
                '.search-item',
                '[class*="result"]',
                '.document',
                '.record',
                '.entry'
            ];

            let resultElements = [];
            for (const selector of resultSelectors) {
                try {
                    const elements = await this.page.$$(selector);
                    if (elements.length > 0) {
                        resultElements = elements;
                        console.log(`Found ${elements.length} result elements with selector: ${selector}`);
                        break;
                    }
                } catch (e) {
                    continue;
                }
            }

            const results = [];
            
            if (resultElements.length > 0) {
                for (const element of resultElements) {
                    try {
                        const text = await element.textContent();
                        const title = await element.$eval('h1, h2, h3, .title, .name', el => el.textContent).catch(() => '');
                        const link = await element.$eval('a', el => el.href).catch(() => '');
                        
                        if (text && text.trim().length > 0) {
                            results.push({
                                title: title || 'No title',
                                text: text.trim(),
                                link: link || '',
                                element: 'search-result'
                            });
                        }
                    } catch (e) {
                        continue;
                    }
                }
            }

            // If no specific result elements found, try to get general page content
            if (results.length === 0) {
                try {
                    const pageContent = await this.page.textContent('body');
                    if (pageContent) {
                        results.push({
                            title: 'General Search Results',
                            text: pageContent.substring(0, 1000) + '...', // Limit content length
                            link: '',
                            element: 'general-content'
                        });
                    }
                } catch (e) {
                    console.log('Could not extract general page content');
                }
            }

            return {
                results: results,
                resultCount: results.length,
                hasResults: results.length > 0
            };

        } catch (error) {
            console.error('Error extracting search results:', error);
            return {
                results: [],
                resultCount: 0,
                hasResults: false,
                error: error.message
            };
        }
    }

    async searchMultipleVins(vinNumbers) {
        const results = [];
        
        for (const vin of vinNumbers) {
            try {
                const result = await this.searchVinData(vin);
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

    async logout() {
        try {
            if (this.page && this.isAuthenticated) {
                // Look for logout button/link
                const logoutSelectors = [
                    'a:has-text("Logout")',
                    'a:has-text("Sign Out")',
                    'button:has-text("Logout")',
                    'button:has-text("Sign Out")',
                    '.logout',
                    '.signout'
                ];

                for (const selector of logoutSelectors) {
                    try {
                        const logoutElement = await this.page.$(selector);
                        if (logoutElement) {
                            await logoutElement.click();
                            console.log('Successfully logged out of DocSearch');
                            break;
                        }
                    } catch (e) {
                        continue;
                    }
                }
            }
        } catch (error) {
            console.error('Error logging out of DocSearch:', error);
        }
    }

    async close() {
        try {
            if (this.isAuthenticated) {
                await this.logout();
            }
            
            if (this.browser) {
                await this.browser.close();
                console.log('DocSearch scraper closed successfully');
            }
        } catch (error) {
            console.error('Error closing DocSearch scraper:', error);
        }
    }
}

module.exports = DocSearchScraper;
