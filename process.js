const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

// Configure multer for memory storage (Vercel doesn't support file system storage)
const upload = multer({
    storage: multer.memoryStorage(),
    fileFilter: (req, file, cb) => {
        const allowedTypes = ['.xlsx', '.xls'];
        const fileExt = path.extname(file.originalname).toLowerCase();
        if (allowedTypes.includes(fileExt)) {
            cb(null, true);
        } else {
            cb(new Error('只支持 .xlsx 和 .xls 文件格式'));
        }
    },
    limits: {
        fileSize: 50 * 1024 * 1024 // 50MB limit
    }
});

class ExcelProcessorVercel {
    // Process Excel file based on type
    processExcelFile(buffer, fileType) {
        try {
            const workbook = XLSX.read(buffer);
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(worksheet);

            console.log(`Processing ${fileType} with ${data.length} rows`);

            let result = {
                originalRows: data.length,
                worksheets: [],
                workbook: workbook
            };

            switch (fileType) {
                case 'Inventory Enquiry AU':
                    result = this.processInventoryAU(data, workbook);
                    break;
                case 'Inventory Enquiry NZ':
                    result = this.processInventoryNZ(data, workbook);
                    break;
                case 'Purchase Item AU':
                    result = this.processPurchaseAU(data, workbook);
                    break;
                case 'Purchase Item NZ':
                    result = this.processPurchaseNZ(data, workbook);
                    break;
                case 'Sales Item AU':
                    result = this.processSalesAU(data, workbook);
                    break;
                case 'Sales Item NZ':
                    result = this.processSalesNZ(data, workbook);
                    break;
                default:
                    throw new Error('不支持的文件类型');
            }

            return result;
        } catch (error) {
            console.error('Processing error:', error);
            throw error;
        }
    }

    processInventoryAU(data, workbook) {
        const result = {
            originalRows: data.length,
            worksheets: [],
            workbook: workbook
        };

        // QLD筛选
        const qldData = data.filter(row =>
            row['Brand Name'] === 'MG' &&
            row['Warehouse'] === 'CEVA QLD' &&
            row['Stock Location'] === 'Normal' &&
            (row['SOH'] || 0) > 0
        );
        this.addWorksheet(workbook, 'QLD', qldData);
        result.worksheets.push({ name: 'QLD', rows: qldData.length });

        // VIC&OFF筛选
        const vicOffData = data.filter(row =>
            row['Brand Name'] === 'MG' &&
            ['CEVA OFFSITE', 'CEVA VIC'].includes(row['Warehouse']) &&
            row['Stock Location'] === 'Normal' &&
            (row['SOH'] || 0) > 0
        );
        this.addWorksheet(workbook, 'VIC&OFF', vicOffData);
        result.worksheets.push({ name: 'VIC&OFF', rows: vicOffData.length });

        // VIC筛选
        const vicData = data.filter(row =>
            row['Brand Name'] === 'MG' &&
            row['Warehouse'] === 'CEVA VIC' &&
            row['Stock Location'] === 'Normal' &&
            (row['SOH'] || 0) > 0
        );
        this.addWorksheet(workbook, 'VIC', vicData);
        result.worksheets.push({ name: 'VIC', rows: vicData.length });

        // OFF筛选
        const offData = data.filter(row =>
            row['Brand Name'] === 'MG' &&
            row['Warehouse'] === 'CEVA OFFSITE' &&
            row['Stock Location'] === 'Normal' &&
            (row['SOH'] || 0) > 0
        );
        this.addWorksheet(workbook, 'OFF', offData);
        result.worksheets.push({ name: 'OFF', rows: offData.length });

        return result;
    }

    processInventoryNZ(data, workbook) {
        const result = {
            originalRows: data.length,
            worksheets: [],
            workbook: workbook
        };

        // NZ筛选
        const nzData = data.filter(row =>
            row['Brand Name'] === 'MG' &&
            row['Warehouse'] === 'CEVA AUC' &&
            row['Stock Location'] === 'Normal' &&
            (row['SOH'] || 0) > 0
        );
        this.addWorksheet(workbook, 'NZ', nzData);
        result.worksheets.push({ name: 'NZ', rows: nzData.length });

        return result;
    }

    processPurchaseAU(data, workbook) {
        const result = {
            originalRows: data.length,
            worksheets: [],
            workbook: workbook
        };

        // 添加Total列
        const dataWithTotal = data.map(row => ({
            ...row,
            'Total': (row['Inbound QTY'] || 0) + (row['Pending QTY'] || 0)
        }));

        const columnsToKeep = ['Brand Name', 'Order #', 'Purchase Group', 'Processed Part #',
                              'Inbound QTY', 'Pending QTY', 'Total', 'Shipment Mode', 'To Warehouse'];

        // QLD筛选
        const qldData = dataWithTotal
            .filter(row =>
                row['Brand Name'] === 'MG' &&
                (row['Total'] || 0) > 0 &&
                row['To Warehouse'] === 'CEVA QLD'
            )
            .map(row => this.selectColumns(row, columnsToKeep));
        this.addWorksheet(workbook, 'QLD', qldData);
        result.worksheets.push({ name: 'QLD', rows: qldData.length });

        // VIC筛选
        const vicData = dataWithTotal
            .filter(row =>
                row['Brand Name'] === 'MG' &&
                (row['Total'] || 0) > 0 &&
                row['To Warehouse'] === 'CEVA VIC'
            )
            .map(row => this.selectColumns(row, columnsToKeep));
        this.addWorksheet(workbook, 'VIC', vicData);
        result.worksheets.push({ name: 'VIC', rows: vicData.length });

        return result;
    }

    processPurchaseNZ(data, workbook) {
        const result = {
            originalRows: data.length,
            worksheets: [],
            workbook: workbook
        };

        // 添加Total列
        const dataWithTotal = data.map(row => ({
            ...row,
            'Total': (row['Inbound QTY'] || 0) + (row['Pending QTY'] || 0)
        }));

        const columnsToKeep = ['Brand Name', 'Order #', 'Purchase Group', 'Processed Part #',
                              'Inbound QTY', 'Pending QTY', 'Total', 'Shipment Mode', 'To Warehouse'];

        // NZ筛选
        const nzData = dataWithTotal
            .filter(row =>
                row['Brand Name'] === 'MG' &&
                (row['Total'] || 0) > 0 &&
                row['To Warehouse'] === 'CEVA AUC'
            )
            .map(row => this.selectColumns(row, columnsToKeep));
        this.addWorksheet(workbook, 'NZ', nzData);
        result.worksheets.push({ name: 'NZ', rows: nzData.length });

        return result;
    }

    processSalesAU(data, workbook) {
        const result = {
            originalRows: data.length,
            worksheets: [],
            workbook: workbook
        };

        const columnsToKeep = ['Brand Name', 'Processed Part #', 'Submitted Time', 'Warehouse', 'BO QTY'];

        // QLD筛选
        const qldData = data
            .filter(row =>
                row['Brand Name'] === 'MG' &&
                (row['BO QTY'] || 0) > 0 &&
                row['Warehouse'] === 'CEVA QLD'
            )
            .map(row => this.selectColumns(row, columnsToKeep));
        this.addWorksheet(workbook, 'QLD', qldData);
        result.worksheets.push({ name: 'QLD', rows: qldData.length });

        // VIC筛选
        const vicData = data
            .filter(row =>
                row['Brand Name'] === 'MG' &&
                (row['BO QTY'] || 0) > 0 &&
                ['CEVA VIC', 'CEVA OFFSITE'].includes(row['Warehouse'])
            )
            .map(row => this.selectColumns(row, columnsToKeep));
        this.addWorksheet(workbook, 'VIC', vicData);
        result.worksheets.push({ name: 'VIC', rows: vicData.length });

        return result;
    }

    processSalesNZ(data, workbook) {
        const result = {
            originalRows: data.length,
            worksheets: [],
            workbook: workbook
        };

        const columnsToKeep = ['Brand Name', 'Processed Part #', 'Submitted Time', 'Warehouse', 'BO QTY'];

        // NZ筛选
        const nzData = data
            .filter(row =>
                row['Brand Name'] === 'MG' &&
                (row['BO QTY'] || 0) > 0 &&
                row['Warehouse'] === 'CEVA AUC'
            )
            .map(row => this.selectColumns(row, columnsToKeep));
        this.addWorksheet(workbook, 'NZ', nzData);
        result.worksheets.push({ name: 'NZ', rows: nzData.length });

        return result;
    }

    selectColumns(row, columns) {
        const result = {};
        columns.forEach(col => {
            result[col] = row[col];
        });
        return result;
    }

    addWorksheet(workbook, sheetName, data) {
        // Remove existing sheet if it exists
        if (workbook.SheetNames.includes(sheetName)) {
            delete workbook.Sheets[sheetName];
            const index = workbook.SheetNames.indexOf(sheetName);
            workbook.SheetNames.splice(index, 1);
        }

        // Add new sheet
        const worksheet = XLSX.utils.json_to_sheet(data);
        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    }
}

const processor = new ExcelProcessorVercel();

export default async function handler(req, res) {
    // Enable CORS
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

    if (req.method === 'OPTIONS') {
        res.status(200).end();
        return;
    }

    if (req.method !== 'POST') {
        return res.status(405).json({ success: false, error: '只支持POST请求' });
    }

    try {
        // Use multer to parse the multipart form data
        const uploadMiddleware = upload.single('file');

        await new Promise((resolve, reject) => {
            uploadMiddleware(req, res, (err) => {
                if (err) reject(err);
                else resolve();
            });
        });

        if (!req.file) {
            return res.status(400).json({ success: false, error: '没有上传文件' });
        }

        const { fileType } = req.body;
        if (!fileType) {
            return res.status(400).json({ success: false, error: '缺少文件类型参数' });
        }

        console.log(`Processing file: ${req.file.originalname}, type: ${fileType}`);

        const result = processor.processExcelFile(req.file.buffer, fileType);

        // Convert workbook to buffer for download
        const outputBuffer = XLSX.write(result.workbook, { type: 'buffer', bookType: 'xlsx' });
        const base64Data = outputBuffer.toString('base64');

        res.json({
            success: true,
            data: {
                filename: `processed_${Date.now()}_${req.file.originalname}`,
                originalRows: result.originalRows,
                worksheets: result.worksheets,
                fileData: base64Data
            }
        });

    } catch (error) {
        console.error('Processing error:', error);
        res.status(500).json({
            success: false,
            error: error.message || '处理文件时发生错误'
        });
    }
}