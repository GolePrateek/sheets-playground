const GSheets = require('./gsheets');

async function main() {
    const gs = new GSheets();
    const spreadsheetId = '1CicXW85QnB71syn-HgzWj-UypJenWR6pdIRTvnUFaX0';
    const range = 'Class Data!A:Z';

    const rows = await gs.getValues(spreadsheetId, range);
    const rowIndicesToHighlight = rows
        .map((row, index) => ({ row, index }))
        .filter(({ row }) => row[1] === 'Male')
        .map(({ index }) => index);
    
    console.log(rowIndicesToHighlight);

    await gs.clearFormatting(spreadsheetId);

    await gs.highlightRows(spreadsheetId, rowIndicesToHighlight);

    // await gs.readAndHighlight(spreadsheetId, range);
}

main().catch(console.error);