class ExcelProcessor {
    constructor() {
        this.currentFile = null;
        this.fileType = null;
        this.processedData = null;
        this.baseURL = window.location.origin;
        this.initializeEventListeners();
    }

    initializeEventListeners() {
        const fileInput = document.getElementById('fileInput');
        const processBtn = document.getElementById('processBtn');
        const downloadBtn = document.getElementById('downloadBtn');
        const newFileBtn = document.getElementById('newFileBtn');

        fileInput.addEventListener('change', (e) => this.handleFileSelect(e));
        processBtn.addEventListener('click', () => this.processFile());
        downloadBtn.addEventListener('click', () => this.downloadFile());
        newFileBtn.addEventListener('click', () => this.resetApplication());
    }

    handleFileSelect(event) {
        const file = event.target.files[0];
        if (!file) return;

        this.currentFile = file;
        this.detectFileType(file.name);
        this.showFileInfo(file);
        this.showProcessingSection();
    }

    detectFileType(filename) {
        const fileTypes = {
            'Inventory Enquiry AU': {
                pattern: /Inventory.*Enquiry.*AU/i,
                description: '澳洲库存查询数据',
                rules: [
                    'QLD: MG品牌 + CEVA QLD仓库 + Normal库位 + SOH>0',
                    'VIC&OFF: MG品牌 + (CEVA OFFSITE或CEVA VIC)仓库 + Normal库位 + SOH>0',
                    'VIC: MG品牌 + CEVA VIC仓库 + Normal库位 + SOH>0',
                    'OFF: MG品牌 + CEVA OFFSITE仓库 + Normal库位 + SOH>0'
                ]
            },
            'Inventory Enquiry NZ': {
                pattern: /Inventory.*Enquiry.*NZ/i,
                description: '新西兰库存查询数据',
                rules: [
                    'NZ: MG品牌 + CEVA AUC仓库 + Normal库位 + SOH>0'
                ]
            },
            'Purchase Item AU': {
                pattern: /Purchase.*Item.*AU/i,
                description: '澳洲在途库存数据',
                rules: [
                    '添加Total列 (Inbound QTY + Pending QTY)',
                    'QLD: MG品牌 + Total>0 + CEVA QLD目标仓库',
                    'VIC: MG品牌 + Total>0 + CEVA VIC目标仓库'
                ]
            },
            'Purchase Item NZ': {
                pattern: /Purchase.*Item.*NZ/i,
                description: '新西兰在途库存数据',
                rules: [
                    '添加Total列 (Inbound QTY + Pending QTY)',
                    'NZ: MG品牌 + Total>0 + CEVA AUC目标仓库'
                ]
            },
            'Sales Item AU': {
                pattern: /Sales.*Item.*AU/i,
                description: '澳洲BO数据',
                rules: [
                    'QLD: MG品牌 + BO QTY>0 + CEVA QLD仓库',
                    'VIC: MG品牌 + BO QTY>0 + (CEVA VIC或CEVA OFFSITE)仓库'
                ]
            },
            'Sales Item NZ': {
                pattern: /Sales.*Item.*NZ/i,
                description: '新西兰BO数据',
                rules: [
                    'NZ: MG品牌 + BO QTY>0 + CEVA AUC仓库'
                ]
            }
        };

        for (const [type, config] of Object.entries(fileTypes)) {
            if (config.pattern.test(filename)) {
                this.fileType = type;
                this.updateProcessingUI(type, config);
                return;
            }
        }

        this.fileType = null;
        this.showError('文件名不符合处理规则，请检查文件名是否包含正确的关键字');
    }

    showFileInfo(file) {
        const fileInfo = document.getElementById('fileInfo');
        const fileName = document.getElementById('fileName');
        const fileSize = document.getElementById('fileSize');

        fileName.textContent = file.name;
        fileSize.textContent = this.formatFileSize(file.size);
        fileInfo.style.display = 'block';
    }

    formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    showProcessingSection() {
        document.getElementById('processingSection').style.display = 'block';
    }

    updateProcessingUI(type, config) {
        const fileTypeIndicator = document.getElementById('fileTypeIndicator');
        const detectedType = document.getElementById('detectedType');
        const rulesList = document.getElementById('rulesList');

        fileTypeIndicator.textContent = type;
        detectedType.textContent = config.description;

        rulesList.innerHTML = '';
        config.rules.forEach(rule => {
            const ruleItem = document.createElement('div');
            ruleItem.className = 'rule-item';
            ruleItem.textContent = rule;
            rulesList.appendChild(ruleItem);
        });
    }

    async processFile() {
        if (!this.currentFile || !this.fileType) {
            this.showError('请先选择有效的Excel文件');
            return;
        }

        this.showProcessingUI();
        this.updateProgress(0, '准备处理文件...');

        try {
            const formData = new FormData();
            formData.append('file', this.currentFile);
            formData.append('fileType', this.fileType);

            this.updateProgress(20, '上传文件中...');

            const response = await fetch(`${this.baseURL}/api/process`, {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                throw new Error(`服务器错误: ${response.status}`);
            }

            this.updateProgress(50, '处理数据中...');

            const result = await response.json();

            if (result.success) {
                this.updateProgress(80, '生成处理结果...');
                this.processedData = result.data;

                setTimeout(() => {
                    this.updateProgress(100, '处理完成');
                    this.showResults(result);
                }, 500);
            } else {
                throw new Error(result.error || '处理失败');
            }

        } catch (error) {
            this.showError(`处理文件时出错: ${error.message}`);
            this.hideProcessingUI();
        }
    }

    showProcessingUI() {
        document.getElementById('processingSection').style.display = 'none';
        document.getElementById('progressSection').style.display = 'block';

        const processBtn = document.getElementById('processBtn');
        const btnText = processBtn.querySelector('.btn-text');
        const btnLoader = processBtn.querySelector('.btn-loader');

        btnText.style.display = 'none';
        btnLoader.style.display = 'block';
        processBtn.disabled = true;
    }

    hideProcessingUI() {
        document.getElementById('progressSection').style.display = 'none';
        document.getElementById('processingSection').style.display = 'block';

        const processBtn = document.getElementById('processBtn');
        const btnText = processBtn.querySelector('.btn-text');
        const btnLoader = processBtn.querySelector('.btn-loader');

        btnText.style.display = 'block';
        btnLoader.style.display = 'none';
        processBtn.disabled = false;
    }

    updateProgress(percentage, status) {
        const progressFill = document.getElementById('progressFill');
        const progressPercentage = document.getElementById('progressPercentage');
        const progressStatus = document.getElementById('progressStatus');
        const logContent = document.getElementById('logContent');

        progressFill.style.width = `${percentage}%`;
        progressPercentage.textContent = `${percentage}%`;
        progressStatus.textContent = status;

        const timestamp = new Date().toLocaleTimeString();
        logContent.textContent += `[${timestamp}] ${status}\n`;
        logContent.scrollTop = logContent.scrollHeight;
    }

    showResults(result) {
        document.getElementById('progressSection').style.display = 'none';
        document.getElementById('resultSection').style.display = 'block';

        const resultSummary = document.getElementById('resultSummary');
        resultSummary.innerHTML = `
            <h4>处理结果摘要</h4>
            <div style="margin-top: 16px;">
                <p><strong>文件类型:</strong> ${this.fileType}</p>
                <p><strong>原始数据行数:</strong> ${result.originalRows || 0}</p>
                <p><strong>生成工作表数量:</strong> ${result.worksheets ? result.worksheets.length : 0}</p>
                ${result.worksheets ? result.worksheets.map(ws =>
                    `<p><strong>${ws.name}:</strong> ${ws.rows} 行数据</p>`
                ).join('') : ''}
            </div>
        `;
    }

    async downloadFile() {
        if (!this.processedData) {
            this.showError('没有可下载的文件');
            return;
        }

        try {
            // For Vercel, we have the file data directly in the response
            if (this.processedData.fileData) {
                // Convert base64 to blob
                const byteCharacters = atob(this.processedData.fileData);
                const byteNumbers = new Array(byteCharacters.length);
                for (let i = 0; i < byteCharacters.length; i++) {
                    byteNumbers[i] = byteCharacters.charCodeAt(i);
                }
                const byteArray = new Uint8Array(byteNumbers);
                const blob = new Blob([byteArray], {
                    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                });

                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = this.processedData.filename;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);
            } else {
                // Fallback for local server
                const response = await fetch(`${this.baseURL}/api/download`, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({
                        filename: this.processedData.filename,
                        fileData: this.processedData.fileData
                    })
                });

                if (!response.ok) {
                    throw new Error('下载失败');
                }

                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = this.processedData.filename;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);
            }

        } catch (error) {
            this.showError(`下载文件时出错: ${error.message}`);
        }
    }

    resetApplication() {
        this.currentFile = null;
        this.fileType = null;
        this.processedData = null;

        document.getElementById('fileInput').value = '';
        document.getElementById('fileInfo').style.display = 'none';
        document.getElementById('processingSection').style.display = 'none';
        document.getElementById('progressSection').style.display = 'none';
        document.getElementById('resultSection').style.display = 'none';

        const logContent = document.getElementById('logContent');
        if (logContent) {
            logContent.textContent = '';
        }

        const progressFill = document.getElementById('progressFill');
        const progressPercentage = document.getElementById('progressPercentage');
        if (progressFill && progressPercentage) {
            progressFill.style.width = '0%';
            progressPercentage.textContent = '0%';
        }
    }

    showError(message) {
        console.error('Error:', message);

        // Create a more user-friendly error display
        const errorDiv = document.createElement('div');
        errorDiv.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            background: linear-gradient(135deg, #FF3B30, #FF6B6B);
            color: white;
            padding: 16px 20px;
            border-radius: 12px;
            box-shadow: 0 4px 20px rgba(255, 59, 48, 0.3);
            z-index: 1000;
            max-width: 400px;
            font-family: 'SF Pro Display', -apple-system, BlinkMacSystemFont, sans-serif;
            font-size: 14px;
            line-height: 1.4;
        `;
        errorDiv.textContent = message;

        document.body.appendChild(errorDiv);

        // Auto remove after 5 seconds
        setTimeout(() => {
            if (errorDiv.parentNode) {
                errorDiv.parentNode.removeChild(errorDiv);
            }
        }, 5000);

        // Also show in console for debugging
        console.error(message);
    }
}

// Initialize the application when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    new ExcelProcessor();
});