const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static('.'));

// Configure multer for file uploads
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        const uploadDir = './uploads';
        if (!fs.existsSync(uploadDir)) {
            fs.mkdirSync(uploadDir);
        }
        cb(null, uploadDir);
    },
    filename: (req, file, cb) => {
        const timestamp = Date.now();
        const originalName = Buffer.from(file.originalname, 'latin1').toString('utf8');
        cb(null, `${timestamp}-${originalName}`);
    }
});

const upload = multer({
    storage: storage,
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

class ExcelProcessorServer {
    constructor() {
        this.processedFiles = new Map();
    }

    // Process Excel file based on type
    async processExcelFile(filePath, fileType) {
        try {
            const workbook = XLSX.readFile(filePath);
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

    saveProcessedFile(workbook, originalFilename) {
        const outputDir = './processed';
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir);
        }

        const timestamp = Date.now();
        const filename = `processed_${timestamp}_${originalFilename}`;
        const filepath = path.join(outputDir, filename);

        XLSX.writeFile(workbook, filepath);
        return { filename, filepath };
    }
}

const processor = new ExcelProcessorServer();

// Routes
app.post('/api/process', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ success: false, error: '没有上传文件' });
        }

        const { fileType } = req.body;
        if (!fileType) {
            return res.status(400).json({ success: false, error: '缺少文件类型参数' });
        }

        console.log(`Processing file: ${req.file.filename}, type: ${fileType}`);

        const result = await processor.processExcelFile(req.file.path, fileType);
        const savedFile = processor.saveProcessedFile(result.workbook, req.file.originalname);

        // Store processed file info for download
        processor.processedFiles.set(savedFile.filename, savedFile.filepath);

        // Clean up uploaded file
        fs.unlinkSync(req.file.path);

        res.json({
            success: true,
            data: {
                filename: savedFile.filename,
                originalRows: result.originalRows,
                worksheets: result.worksheets
            }
        });

    } catch (error) {
        console.error('Processing error:', error);

        // Clean up uploaded file if it exists
        if (req.file && fs.existsSync(req.file.path)) {
            fs.unlinkSync(req.file.path);
        }

        res.status(500).json({
            success: false,
            error: error.message || '处理文件时发生错误'
        });
    }
});

app.post('/api/download', (req, res) => {
    try {
        const { filename } = req.body;

        if (!filename || !processor.processedFiles.has(filename)) {
            return res.status(404).json({ success: false, error: '文件不存在' });
        }

        const filepath = processor.processedFiles.get(filename);

        if (!fs.existsSync(filepath)) {
            processor.processedFiles.delete(filename);
            return res.status(404).json({ success: false, error: '文件已被删除' });
        }

        res.download(filepath, filename, (err) => {
            if (err) {
                console.error('Download error:', err);
                res.status(500).json({ success: false, error: '下载失败' });
            } else {
                // Clean up processed file after download
                setTimeout(() => {
                    if (fs.existsSync(filepath)) {
                        fs.unlinkSync(filepath);
                        processor.processedFiles.delete(filename);
                    }
                }, 5000); // Delete after 5 seconds
            }
        });

    } catch (error) {
        console.error('Download error:', error);
        res.status(500).json({ success: false, error: '下载失败' });
    }
});

// Serve the main page
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// Error handling middleware
app.use((error, req, res, next) => {
    if (error instanceof multer.MulterError) {
        if (error.code === 'LIMIT_FILE_SIZE') {
            return res.status(400).json({ success: false, error: '文件大小超过限制 (50MB)' });
        }
    }

    console.error('Server error:', error);
    res.status(500).json({ success: false, error: '服务器内部错误' });
});

// Clean up old files periodically
setInterval(() => {
    const uploadDir = './uploads';
    const processedDir = './processed';

    [uploadDir, processedDir].forEach(dir => {
        if (fs.existsSync(dir)) {
            const files = fs.readdirSync(dir);
            const now = Date.now();

            files.forEach(file => {
                const filepath = path.join(dir, file);
                const stats = fs.statSync(filepath);
                const age = now - stats.mtime.getTime();

                // Delete files older than 1 hour
                if (age > 60 * 60 * 1000) {
                    fs.unlinkSync(filepath);
                    console.log(`Cleaned up old file: ${filepath}`);
                }
            });
        }
    });
}, 30 * 60 * 1000); // Run every 30 minutes

app.listen(PORT, () => {
    console.log(`Excel Processor Server running on http://localhost:${PORT}`);
    console.log('Ready to process Excel files with Apple-style interface!');
});