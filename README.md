# Excel数据自动化处理工具 - 网页版

基于原Python程序开发的现代化网页版Excel处理工具，采用苹果科技风格界面设计。

## 功能特性

### 支持的文件类型
- **Inventory Enquiry AU** - 澳洲库存查询数据
- **Inventory Enquiry NZ** - 新西兰库存查询数据
- **Purchase Item AU** - 澳洲在途库存数据
- **Purchase Item NZ** - 新西兰在途库存数据
- **Sales Item AU** - 澳洲BO数据
- **Sales Item NZ** - 新西兰BO数据

### 处理规则
根据文件名自动识别类型并应用相应的筛选规则：

#### Inventory Enquiry AU
- QLD: MG品牌 + CEVA QLD仓库 + Normal库位 + SOH>0
- VIC&OFF: MG品牌 + (CEVA OFFSITE或CEVA VIC)仓库 + Normal库位 + SOH>0
- VIC: MG品牌 + CEVA VIC仓库 + Normal库位 + SOH>0
- OFF: MG品牌 + CEVA OFFSITE仓库 + Normal库位 + SOH>0

#### Purchase Item AU/NZ
- 自动添加Total列 (Inbound QTY + Pending QTY)
- 按目标仓库筛选数据

#### Sales Item AU/NZ
- 按BO数量和仓库筛选数据

## 技术栈

- **前端**: HTML5, CSS3, JavaScript (ES6+)
- **后端**: Node.js, Express
- **Excel处理**: xlsx库
- **文件上传**: multer
- **界面风格**: Apple Design System inspired

## 安装和运行

### 1. 安装依赖
```bash
npm install
```

### 2. 启动服务器
```bash
npm start
```

### 3. 开发模式 (自动重启)
```bash
npm run dev
```

### 4. 访问应用
打开浏览器访问: http://localhost:3000

## 使用方法

1. **选择文件**: 点击"选择文件"按钮上传Excel文件
2. **自动识别**: 系统根据文件名自动识别处理类型
3. **查看规则**: 确认处理规则和筛选条件
4. **开始处理**: 点击"开始处理"按钮
5. **查看进度**: 实时查看处理进度和日志
6. **下载结果**: 处理完成后下载处理后的文件

## 界面特色

### Apple风格设计
- 简洁现代的界面布局
- 流畅的动画过渡效果
- 优雅的渐变色彩搭配
- 响应式设计支持移动端

### 用户体验
- 拖拽上传文件支持
- 实时处理进度显示
- 详细的操作日志
- 智能错误提示

## 文件结构

```
excel_processor/
├── package.json          # 项目配置和依赖
├── server.js             # Express服务器
├── index.html            # 主页面
├── styles.css            # Apple风格样式
├── script.js             # 前端交互逻辑
├── uploads/              # 临时上传目录
├── processed/            # 处理结果目录
└── README.md            # 项目说明
```

## API接口

### POST /api/process
处理上传的Excel文件
- 参数: file (文件), fileType (文件类型)
- 返回: 处理结果和统计信息

### POST /api/download
下载处理后的文件
- 参数: filename (文件名)
- 返回: Excel文件流

## 安全特性

- 文件类型验证
- 文件大小限制 (50MB)
- 自动清理临时文件
- 错误处理和日志记录

## 浏览器兼容性

- Chrome 80+
- Firefox 75+
- Safari 13+
- Edge 80+

## 开发说明

基于原Python程序的完整功能移植，保持了所有数据处理逻辑的一致性，同时提供了更现代化的Web界面体验。