# VIN Recall Scraper

A Node.js application that allows users to upload Excel files containing VIN numbers, scrapes recall data from Ford's website and DocSearch, and exports the results to a downloadable Excel file.

## Features

- Upload CSV, XLS, or XLSX files
- **Smart VIN Detection**: Auto-detect VIN column or manually specify (A, B, C, etc.)
- Extract VIN numbers from uploaded files
- Scrape recall data from Ford's website
- Authenticate and scrape additional data from DocSearch
- Export processed data to downloadable Excel file
- **Dynamic Column Selection**: See available columns and their names

## Tech Stack

- **Backend**: Node.js with Express
- **Web Scraping**: Playwright
- **Excel Processing**: xlsx library
- **Frontend**: HTML/CSS/JavaScript

## Prerequisites

Before running this application, you need to install Node.js:

1. **Install Node.js** (if not already installed):
   - Download from [nodejs.org](https://nodejs.org/)
   - Choose the LTS version for Windows
   - Run the installer and follow the setup wizard
   - Restart your terminal/command prompt after installation

2. **Verify installation**:
   ```bash
   node --version
   npm --version
   ```

## Installation

1. Install dependencies:
```bash
npm install
```

2. Install Playwright browsers:
```bash
npx playwright install
```

3. Create a `.env` file with your DocSearch credentials:
   - Copy `env.example` to `.env`
   - Replace the placeholder values with your actual DocSearch credentials:
   ```
   DOCSEARCH_USERNAME=your_actual_username
   DOCSEARCH_PASSWORD=your_actual_password
   ```

## Usage

1. Start the server:
```bash
npm start
```

2. Open your browser and navigate to `http://localhost:3000`

3. Upload an Excel file containing VIN numbers

4. **Select VIN Column**: Choose auto-detect or specify which column (A, B, C, etc.) contains VIN numbers

5. Click "Process File" and wait for the scraping process to complete

6. Download the processed Excel file with recall data

## Column Selection

The application offers two ways to find VIN numbers:

- **Auto-detect (Recommended)**: Automatically searches for columns with VIN-like data or common VIN column names
- **Manual Selection**: Choose a specific column (A, B, C, etc.) where VIN numbers are located

When you upload a file, the application will show you all available columns with their names, making it easy to identify the correct column.

## File Structure

```
├── server.js          # Main server file
├── public/            # Static files
│   ├── index.html     # Frontend interface
│   ├── style.css      # Styling
│   └── script.js      # Frontend JavaScript
├── uploads/           # Temporary file storage
├── downloads/         # Generated Excel files
└── scraper/           # Scraping modules
    ├── fordScraper.js
    └── docsearchScraper.js
```
