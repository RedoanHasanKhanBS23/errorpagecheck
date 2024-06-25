const { chromium, firefox, webkit } = require('playwright');
const ExcelJS = require('exceljs');

async function openUrlsAndVerifyElements() {
    const workbook = new ExcelJS.Workbook();
    try {
        await workbook.xlsx.readFile('urls.xlsx'); // Replace 'urls.xlsx' with your file name

        const worksheet = workbook.getWorksheet(1);
        const rowCount = worksheet.rowCount;

        for (let i = 1; i <= rowCount; i++) { // Ensure we iterate through all rows
            const urlCell = worksheet.getCell(`A${i}`);
            const url = urlCell.value;

            const elementSelectorCell = worksheet.getCell(`B${i}`);
            const elementSelector = elementSelectorCell.value;

            if (typeof url === 'string' && url.trim() !== '' && typeof elementSelector === 'string' && elementSelector.trim() !== '') { // Check if URL and element selector are valid
                const browsers = [chromium, firefox, webkit]; // List of supported browsers
                let result = 'Failed';

                for (const browserType of browsers) {
                    const browser = await browserType.launch({ headless: false });
                    const page = await browser.newPage();
                    try {
                        console.log(`Navigating to URL: ${url} using ${browserType.name()}`);
                        await page.goto(url, { waitUntil: 'networkidle' });

                        // Take a screenshot for inspection
                        await page.screenshot({ path: `screenshot-${i}-${browserType.name()}.png` });

                        // Evaluate the presence of the element in the DOM
                        const elementExists = await page.evaluate((selector) => {
                            return document.querySelector(selector) !== null;
                        }, elementSelector);

                        if (elementExists) {
                            result = 'Success';
                            break;
                        }
                    } catch (error) {
                        console.error(`Error navigating to ${url} using ${browserType.name()}: ${error.message}`);
                    } finally {
                        await browser.close();
                    }
                }

                console.log(`Element check for URL ${url}: ${result}`);
                worksheet.getCell(`C${i}`).value = result; // Update column C with the result
            } else {
                worksheet.getCell(`C${i}`).value = 'Invalid URL or Selector';
            }
        }

        // Save the updated workbook
        await workbook.xlsx.writeFile('urls.xlsx');
        console.log('Successfully saved the updated Excel file.');
    } catch (error) {
        console.error('Error reading or writing Excel file:', error.message);
    }
}

openUrlsAndVerifyElements().then(() => console.log('Finished opening URLs and verifying element presence'));
