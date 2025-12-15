@echo off
echo VIN Recall Scraper Setup Script
echo ================================

echo.
echo Checking for Node.js installation...
node --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Node.js is not installed!
    echo.
    echo Please install Node.js first:
    echo 1. Go to https://nodejs.org/
    echo 2. Download the LTS version for Windows
    echo 3. Run the installer
    echo 4. Restart your command prompt
    echo 5. Run this setup script again
    echo.
    pause
    exit /b 1
)

echo Node.js is installed!
node --version

echo.
echo Checking for npm...
npm --version
if %errorlevel% neq 0 (
    echo ERROR: npm is not available!
    pause
)

echo.
echo Installing project dependencies...
npm install
if %errorlevel% neq 0 (
    echo ERROR: Failed to install dependencies!
    pause
    exit /b 1
)

echo.
echo Installing Playwright browsers...
npx playwright install
if %errorlevel% neq 0 (
    echo ERROR: Failed to install Playwright browsers!
    pause
    exit /b 1
)

echo.
echo Setting up environment file...
if not exist .env (
    copy env.example .env
    echo Created .env file from template
    echo IMPORTANT: Please edit .env file and add your DocSearch credentials!
) else (
    echo .env file already exists
)

echo.
echo Setup completed successfully!
echo.
echo Next steps:
echo 1. Edit .env file and add your DocSearch username and password
echo 2. Run: npm start
echo 3. Open http://localhost:3000 in your browser
echo.
pause
