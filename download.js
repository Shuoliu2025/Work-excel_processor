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
        const { filename, fileData } = req.body;

        if (!filename || !fileData) {
            return res.status(400).json({ success: false, error: '缺少文件名或文件数据' });
        }

        // Convert base64 back to buffer
        const buffer = Buffer.from(fileData, 'base64');

        // Set headers for file download
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        res.setHeader('Content-Length', buffer.length);

        // Send the file
        res.send(buffer);

    } catch (error) {
        console.error('Download error:', error);
        res.status(500).json({ success: false, error: '下载失败' });
    }
}