import * as ExcelJS from 'exceljs';
import * as readline from 'readline';

const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
});

interface SearchResult {
    row: number;
    column: number;
    data: String;
}

interface SearchResults {
    [keyword: string]: SearchResult[];
}

rl.question('FilePath or default : ', async (arg) => {

    const filePath = arg === 'default' ? __dirname + '/iw-tech-test-retailer-data.xlsx' : arg

    if (filePath && filePath?.split('.').pop() !== 'xlsx') {
        console.error('Only xlsx format is supported.')
        process.exit(1);
    }

    rl.question('Retailers : ', async (retailersString) => {
        rl.close();

        const retailersList = retailersString?.split(',')?.map(e => e?.trim())?.filter(e => e && e !== '')

        try {
            let data: SearchResults = await searchKeywords(filePath, retailersList)
            console.log('Search results:', data);
            process.exit(1);
        } catch (error: any) {
            console.error(error.message);
            process.exit(1);
        }
    })
});

async function searchKeywords(filePath: string, keywords: string[]): Promise<SearchResults> {
    const results: SearchResults = {};

    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);

        workbook.eachSheet((worksheet) => {
            worksheet.eachRow((row, rowIndex) => {
                row.eachCell((cell, colIndex) => {

                    for (const keyword of keywords) {
                        if (cell.value && cell.value.toString().includes(keyword)) {
                            if (!results[keyword]) {
                                results[keyword] = [];
                            }
                            results[keyword].push({ row: rowIndex, column: colIndex, data: cell.value.toString() });
                        }
                    }

                });
            });
        });
    } catch (error: any) {
        console.error(error.message);
        process.exit(1);
    }

    return results;
}

