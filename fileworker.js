importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js', 'https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.29.1/moment.min.js');

self.onmessage = function(event) {
    const { file } = event.data;
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        let aggregatedData = [];
        workbook.SheetNames.forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName];
            const parsedData = XLSX.utils.sheet_to_json(sheet, { defval: '' });
            aggregatedData = aggregatedData.concat(parsedData);
        });

        const summaryData = computeTestingTimes(aggregatedData, [
            "INIT", "PRE-CHECK", "AST", "FLA", "IOT", "FCT", "FC2", "IST", "FPF", "NVL"
        ]);

        self.postMessage(summaryData);
    };

    reader.onerror = function() {
        self.postMessage([]);
    };

    reader.readAsArrayBuffer(file);
};

function computeTestingTimes(data, stageOrder) {
    const results = [];

    data.forEach(row => {
        const sku = row["SkuName"];
        const stage = row["Stage"];
        const startTime = parseExcelDate(row["StartTime"]);
        const endTime = parseExcelDate(row["EndTime"]);

        if (!startTime || !endTime || !stageOrder.includes(stage)) return;

        const duration = (new Date(endTime) - new Date(startTime)) / 60000; // in minutes
        const record = results.find(r => r.sku === sku) || { sku, Total: 0 };

        if (!record[stage]) {
            record[stage] = duration.toFixed(2);
            record.Total += duration;
        }

        if (!results.includes(record)) results.push(record);
    });

    return results.map(r => ({ ...r, Total: r.Total.toFixed(2) }));
}

function parseExcelDate(date) {
    return moment(date, [
        'YYYY-MM-DD HH:mm:ss', 'MM/DD/YYYY h:mm:ss A'
    ], true).isValid() ? moment(date).toDate() : null;
}
