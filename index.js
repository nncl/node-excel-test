const express = require('express')

const app = express()
const port = 3000

const ExcelJS = require('exceljs')

app.get('/', async (req, res) => {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Me';
    workbook.lastModifiedBy = 'Her';
    workbook.created = new Date(1985, 8, 30);
    workbook.modified = new Date();
    workbook.lastPrinted = new Date(2016, 9, 27);

    const worksheet = workbook.addWorksheet('My Sheet', {
        headerFooter: {firstHeader: "MY TITLE", firstFooter: "MY FOOTER TITLE"}
    });

    worksheet.columns = [
        {header: '', key: 'id'}, // , style: {font: {size: 14}}
        {header: '', key: 'value'}, // , style: {font: {size: 10}}
    ];

    worksheet.addRow({id: 'REPORT SALES', value: '25/01/2022 15h30'});
    worksheet.addRow({id: '', value: ''});
    worksheet.addRow({id: 'REGION', value: 'SP'});
    worksheet.addRow({id: 'PLACE', value: 'Petstore'});
    worksheet.addRow({id: 'STATUS', value: 'Active'});
    worksheet.addRow({id: 'PERIOD', value: '25/01/2022 to 26/01/2022'});

    worksheet.addTable({
        name: 'MyTable 2',
        ref: 'A10',
        headerRow: true,
        style: {
            theme: 'TableStyleDark3',
            showRowStripes: true,
        },
        columns: [
            {name: 'Date'},
            {name: 'Amount'},
            {name: 'Amount 2'},
            {name: 'Amount 3'},
            {name: 'Amount 4'},
            {name: 'Amount 5'},
        ],
        rows: [
            [new Date('2019-01-20'), 170.10, 3, 4, 5, 6],
            [new Date('2019-01-20'), 170.10, 3, 4, 5, 6],
            [new Date('2019-01-20'), 170.10, 3, 4, 5, 6],
            [new Date('2019-01-20'), 170.10, 3, 4, 5, 6],
            [new Date('2019-01-20'), 170.10, 3, 4, 5, 6],
        ],
    });

    res.attachment("report.xlsx");
    res.status(200);
    await workbook.xlsx.write(res)

    return res.end();
})

app.listen(port, () => {
    console.log(`Example app listening on port ${port}`)
})
