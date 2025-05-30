const express = require('express');
const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');
const app = express();
const PORT = 3000;

// Middleware to log requests
app.use((req, res, next) => {
    console.log(`${new Date().toISOString()} - ${req.method} ${req.url}`);
    next();
});

// Route to trigger scraping
app.get('/scrape-jobs', async (req, res) => {
    try {
        console.log("Starting job scraping process...");
        
        // Step 1: Scrape job data
        const jobData = await scrapeJobs();
        if (!jobData || jobData.length === 0) {
            throw new Error('No jobs found or scraping failed');
        }
        console.log(`Successfully scraped ${jobData.length} jobs`);
        
        // Step 2: Export to Excel
        const excelPath = await exportToExcel(jobData);
        console.log(`Excel file saved to: ${excelPath}`);
        
        // Step 3: Respond with results
        res.json({
            success: true,
            message: 'Scraping completed successfully',
            stats: {
                totalJobs: jobData.length,
                companies: [...new Set(jobData.map(job => job.company))].length
            },
            excelLink: 'https://excel.cloud.microsoft/open/onedrive/?docId=1D68E6C0AFABEAC2%21sc29a48d4ffe2427ba084b2d182cfc5ec&driveId=1D68E6C0AFABEAC2',
            dataSample: jobData.slice(0, 3) // Return first 3 jobs as sample
        });

    } catch (error) {
        console.error("Scraping error:", error);
        res.status(500).json({ 
            success: false,
            error: error.message,
            details: process.env.NODE_ENV === 'development' ? error.stack : undefined
        });
    }
});

// Job scraping function
async function scrapeJobs() {
    let browser;
    try {
        // Launch browser with additional options for reliability
        browser = await puppeteer.launch({ 
            headless: "new",
            args: [
                '--no-sandbox',
                '--disable-setuid-sandbox',
                '--disable-dev-shm-usage'
            ],
            timeout: 30000
        });

        const page = await browser.newPage();
        
        // Set realistic user agent and viewport
        await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36');
        await page.setViewport({ width: 1366, height: 768 });

        console.log("Navigating to jobs page...");
        await page.goto('https://www.timesjobs.com/candidate/job-search.html?searchType=Home_Search&from=submit&asKey=OFF&txtKeywords=&cboPresFuncArea=35', {
            waitUntil: 'networkidle2',
            timeout: 60000
        });

        // Wait for job listings with timeout
        console.log("Waiting for job listings...");
        await page.waitForSelector('.job-bx', { timeout: 15000 });

        // Add delay to mimic human behavior (using Promise instead of waitForTimeout)
        await new Promise(resolve => setTimeout(resolve, 2000 + Math.random() * 3000));

        // Extract job data
        console.log("Extracting job data...");
        const jobs = await page.evaluate(() => {
            const jobElements = Array.from(document.querySelectorAll('.job-bx'));
            return jobElements.map(job => {
                return {
                    title: job.querySelector('h2 a')?.innerText.trim() || 'N/A',
                    company: job.querySelector('.joblist-comp-name')?.innerText.trim().replace(/\s+/g, ' ') || 'N/A',
                    location: job.querySelector('.top-jd-dtl li:first-child')?.innerText.trim() || 'N/A',
                    jobType: job.querySelector('.job-type')?.innerText.trim() || 'N/A',
                    postedDate: job.querySelector('.posted-date')?.innerText.trim() || 'N/A',
                    description: job.querySelector('.list-job-dtl li:first-child')?.innerText.trim().replace(/\s+/g, ' ') || 'N/A'
                };
            });
        });

        return jobs;

    } catch (error) {
        console.error("Error during scraping:", error);
        throw error;
    } finally {
        if (browser) {
            await browser.close();
        }
    }
}

// Excel export function
async function exportToExcel(jobData) {
    try {
        console.log("Creating Excel workbook...");
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Tech Jobs');

        // Add headers with styling
        worksheet.columns = [
            { header: 'Job Title', key: 'title', width: 30 },
            { header: 'Company', key: 'company', width: 25 },
            { header: 'Location', key: 'location', width: 20 },
            { header: 'Job Type', key: 'jobType', width: 15 },
            { header: 'Posted Date', key: 'postedDate', width: 15 },
            { header: 'Description', key: 'description', width: 50 }
        ];

        // Style header row
        worksheet.getRow(1).font = { bold: true };
        worksheet.getRow(1).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFD3D3D3' }
        };

        // Add data rows
        jobData.forEach(job => {
            worksheet.addRow(job);
        });

        // Auto-filter
        worksheet.autoFilter = {
            from: 'A1',
            to: 'F1'
        };

        // Save file
        const fileName = `jobs_${new Date().toISOString().split('T')[0]}.xlsx`;
        await workbook.xlsx.writeFile(fileName);
        
        return fileName;

    } catch (error) {
        console.error("Error exporting to Excel:", error);
        throw error;
    }
}

// Error handling middleware
app.use((err, req, res, next) => {
    console.error('Unhandled error:', err);
    res.status(500).json({
        success: false,
        message: 'Internal server error',
        error: err.message
    });
});

// Start server
app.listen(PORT, () => {
    console.log(`\nServer running on http://localhost:${PORT}`);
    console.log(`Access the scraper endpoint at: http://localhost:${PORT}/scrape-jobs`);
    console.log(`Press CTRL+C to stop\n`);
});

// Handle shutdown gracefully
process.on('SIGINT', () => {
    console.log('\nServer shutting down...');
    process.exit(0);
});