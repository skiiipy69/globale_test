const axios = require('axios');
const cheerio = require('cheerio');
const fs = require('fs');
const XLSX = require('xlsx');

async function scrapeWebsite(url) {
    try {
        const response = await axios.get(url);
        const $ = cheerio.load(response.data);
        const table = $('.table-results');
        const data = [];

        if (table.length) {
            const rows = table.find('tr');
            rows.slice(1).each((index, row) => {
                const cols = $(row).find('td');
                if (cols.length >= 5) {
                    const row_data = {
                        'URL': url,
                        'Entreprise': $(cols[0]).text().trim(),
                        'Enveloppes administratives': $(cols[1]).text().trim(),
                        'Enveloppes techniques': $(cols[2]).text().trim() || '',
                        'Enveloppes Financières Avant Correction': $(cols[3]).text().trim() || '',
                        'Enveloppes Financières Après Correction': $(cols[4]).text().trim()
                    };

                    const date_limite = $('#ctl0_CONTENU_PAGE_idEntrepriseConsultationSummary_dateHeureLimiteRemisePlis').text().trim();
                    const reference = $('#ctl0_CONTENU_PAGE_idEntrepriseConsultationSummary_reference').text().trim();
                    const objet = $('#ctl0_CONTENU_PAGE_idEntrepriseConsultationSummary_objet').text().trim();
                    
                    // Extracting referentiel_zone_text using Cheerio
                    const referentiel_zone_text = $('#ctl0_CONTENU_PAGE_idEntrepriseConsultationSummary_idReferentielZoneText_RepeaterReferentielZoneText_ctl0_labelReferentielZoneText').text().trim();

                    if (date_limite) row_data['Date limite'] = date_limite;
                    if (reference) row_data['Reference'] = reference;
                    if (objet) row_data['Objet'] = objet;
                    if (referentiel_zone_text) row_data['Estimation'] = referentiel_zone_text;

                    data.push(row_data);
                }
            });
        }
        return data;
    } catch (error) {
        console.error(`Error scraping ${url}: ${error}`);
        return [];
    }
}

// Read URLs from urls.txt
const urlList = fs.readFileSync('submission_links.txt', 'utf8').split('\n').filter(Boolean);

// Load existing data from Excel file, if any
let existingData = [];
try {
    const workbook = XLSX.readFile('scraped_data.xlsx');
    existingData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
} catch (error) {
    console.error('No existing data found.');
}

// Create a new workbook
const newWorkbook = XLSX.utils.book_new();

// Counter for sheet names
let sheetCounter = 1;

// Scrape data from each URL and add it to the workbook
(async () => {
    for (const url of urlList) {
        console.log('Scraping data from:', url);
        const data = await scrapeWebsite(url);
        
        // Convert data to XLSX format
        const worksheet = XLSX.utils.json_to_sheet(data);

        // Add the worksheet to the workbook with generic sheet name
        const sheetName = 'Sheet ' + sheetCounter++;
        XLSX.utils.book_append_sheet(newWorkbook, worksheet, sheetName);
        
        console.log('Data scraped from', url, 'and added to', sheetName);
    }

    // Write the workbook to a file
    XLSX.writeFile(newWorkbook, 'scraped_data_with_sheets.xlsx');
    console.log('Data saved to scraped_data_with_sheets.xlsx');
})();
