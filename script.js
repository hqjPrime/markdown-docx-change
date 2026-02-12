// 全局变量
let currentConverter = 'md-to-docx';
let currentInput = 'text';
let conversionResult = null;

// DOM 元素
const mdToDocxBtn = document.getElementById('md-to-docx');
const docxToMdBtn = document.getElementById('docx-to-md');
const inputFileBtn = document.getElementById('input-file');
const inputTextBtn = document.getElementById('input-text');
const fileInput = document.getElementById('file-input');
const textInput = document.getElementById('text-input');
const convertBtn = document.getElementById('convert-btn');
const status = document.getElementById('status');
const downloadBtn = document.getElementById('download-btn');
const outputContent = document.getElementById('output-content');

// 初始化事件监听器
function initEventListeners() {
    // 转换类型选择
    mdToDocxBtn.addEventListener('click', () => {
        setConverter('md-to-docx');
    });
    
    docxToMdBtn.addEventListener('click', () => {
        setConverter('docx-to-md');
    });
    
    // 输入方式选择
    inputFileBtn.addEventListener('click', () => {
        setInput('file');
    });
    
    inputTextBtn.addEventListener('click', () => {
        setInput('text');
    });
    
    // 文件上传
    fileInput.addEventListener('change', handleFileUpload);
    
    // 转换按钮
    convertBtn.addEventListener('click', handleConversion);
    
    // 下载按钮
    downloadBtn.addEventListener('click', downloadResult);
}

// 设置转换类型
function setConverter(type) {
    currentConverter = type;
    
    // 更新按钮状态
    mdToDocxBtn.classList.remove('active');
    docxToMdBtn.classList.remove('active');
    
    if (type === 'md-to-docx') {
        mdToDocxBtn.classList.add('active');
        fileInput.accept = '.md';
        textInput.placeholder = '在此输入 Markdown 内容';
    } else {
        docxToMdBtn.classList.add('active');
        fileInput.accept = '.docx';
        textInput.placeholder = '在此输入 Word 内容';
    }
    
    // 重置状态
    resetState();
}

// 设置输入方式
function setInput(type) {
    currentInput = type;
    
    // 更新按钮状态
    inputFileBtn.classList.remove('active');
    inputTextBtn.classList.remove('active');
    
    if (type === 'file') {
        inputFileBtn.classList.add('active');
        fileInput.style.display = 'block';
        textInput.style.display = 'none';
    } else {
        inputTextBtn.classList.add('active');
        fileInput.style.display = 'none';
        textInput.style.display = 'block';
    }
    
    // 重置状态
    resetState();
}

// 处理文件上传
function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    resetState();
    
    if (currentConverter === 'md-to-docx' && file.type !== 'text/markdown' && !file.name.endsWith('.md')) {
        showStatus('错误：请上传 .md 文件', 'error');
        return;
    }
    
    if (currentConverter === 'docx-to-md' && !file.name.endsWith('.docx')) {
        showStatus('错误：请上传 .docx 文件', 'error');
        return;
    }
    
    showStatus('文件已上传，准备转换...');
}

// 处理转换
function handleConversion() {
    resetState();
    showStatus('转换中，请稍候...');
    
    if (currentInput === 'file') {
        const file = fileInput.files[0];
        if (!file) {
            showStatus('错误：请先上传文件', 'error');
            return;
        }
        
        if (currentConverter === 'md-to-docx') {
            convertMdFileToDocx(file);
        } else {
            convertDocxFileToMd(file);
        }
    } else {
        const text = textInput.value;
        if (!text.trim()) {
            showStatus('错误：请输入内容', 'error');
            return;
        }
        
        if (currentConverter === 'md-to-docx') {
            convertMdTextToDocx(text);
        } else {
            convertDocxTextToMd(text);
        }
    }
}

// Markdown 文件转 Word
function convertMdFileToDocx(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        const mdText = e.target.result;
        convertMdTextToDocx(mdText);
    };
    reader.readAsText(file);
}

// Markdown 文本转 Word
function convertMdTextToDocx(mdText) {
    try {
        showProgress(20);
        
        const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType } = docx;
        
        showProgress(40);
        
        // 解析 Markdown 文本
        const lines = mdText.split('\n');
        const children = [];
        
        for (let line of lines) {
            line = line.trim();
            
            if (!line) {
                // 空行
                children.push(new Paragraph({ text: '' }));
                continue;
            }
            
            // 处理标题
            if (line.startsWith('# ')) {
                children.push(new Paragraph({
                    text: line.substring(2),
                    heading: HeadingLevel.HEADING_1,
                    spacing: { before: 240, after: 120 }
                }));
            } else if (line.startsWith('## ')) {
                children.push(new Paragraph({
                    text: line.substring(3),
                    heading: HeadingLevel.HEADING_2,
                    spacing: { before: 200, after: 100 }
                }));
            } else if (line.startsWith('### ')) {
                children.push(new Paragraph({
                    text: line.substring(4),
                    heading: HeadingLevel.HEADING_3,
                    spacing: { before: 160, after: 80 }
                }));
            } else if (line.startsWith('#### ')) {
                children.push(new Paragraph({
                    text: line.substring(5),
                    heading: HeadingLevel.HEADING_4,
                    spacing: { before: 140, after: 60 }
                }));
            } else if (line.startsWith('##### ')) {
                children.push(new Paragraph({
                    text: line.substring(6),
                    heading: HeadingLevel.HEADING_5,
                    spacing: { before: 120, after: 40 }
                }));
            } else if (line.startsWith('###### ')) {
                children.push(new Paragraph({
                    text: line.substring(7),
                    heading: HeadingLevel.HEADING_6,
                    spacing: { before: 100, after: 20 }
                }));
            } else if (line.startsWith('- ') || line.startsWith('* ')) {
                // 处理无序列表
                children.push(new Paragraph({
                    text: line.substring(2),
                    bullet: {
                        level: 0
                    },
                    spacing: { before: 100, after: 50 }
                }));
            } else if (line.match(/^\d+\. /)) {
                // 处理有序列表
                children.push(new Paragraph({
                    text: line.replace(/^\d+\. /, ''),
                    numbering: {
                        reference: 'default-numbering',
                        level: 0
                    },
                    spacing: { before: 100, after: 50 }
                }));
            } else if (line.startsWith('```')) {
                // 代码块（简单处理）
                continue;
            } else if (line.startsWith('> ')) {
                // 引用
                children.push(new Paragraph({
                    text: line.substring(2),
                    indent: { left: 720 },
                    spacing: { before: 100, after: 50 }
                }));
            } else if (line.startsWith('---') || line.startsWith('***')) {
                // 分隔线
                children.push(new Paragraph({
                    children: [
                        new TextRun({
                            text: '─'.repeat(50),
                            color: '999999'
                        })
                    ],
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 200, after: 200 }
                }));
            } else {
                // 普通段落
                children.push(new Paragraph({
                    text: line,
                    spacing: { before: 100, after: 100 }
                }));
            }
        }
        
        showProgress(60);
        
        // 创建文档
        const doc = new Document({
            sections: [{
                properties: {},
                children: children
            }]
        });
        
        showProgress(80);
        
        // 生成 Word 文档
        Packer.toBlob(doc).then(blob => {
            showProgress(100);
            setTimeout(() => {
                conversionResult = blob;
                showStatus('转换成功！');
                downloadBtn.disabled = false;
                outputContent.textContent = '转换结果已生成，点击下载按钮保存 Word 文档。';
            }, 500);
        }).catch(error => {
            showStatus('转换失败：' + error.message, 'error');
        });
    } catch (error) {
        showStatus('转换失败：' + error.message, 'error');
    }
}

// Word 文件转 Markdown
function convertDocxFileToMd(file) {
    showProgress(20);
    
    const reader = new FileReader();
    reader.onload = function(e) {
        showProgress(40);
        const arrayBuffer = e.target.result;
        
        try {
            showProgress(60);
            
            // 使用 mammoth.js 解析 Word 文档
            mammoth.convertToMarkdown({ arrayBuffer: arrayBuffer })
                .then(result => {
                    showProgress(80);
                    const mdText = result.value;
                    
                    showProgress(100);
                    setTimeout(() => {
                        conversionResult = mdText;
                        showStatus('转换成功！');
                        downloadBtn.disabled = false;
                        outputContent.textContent = mdText;
                    }, 500);
                    
                    // 检查是否有警告信息
                    if (result.messages && result.messages.length > 0) {
                        console.log('转换警告:', result.messages);
                    }
                })
                .catch(error => {
                    showStatus('转换失败：' + error.message, 'error');
                    console.error('详细错误:', error);
                });
        } catch (error) {
            showStatus('转换失败：' + error.message, 'error');
            console.error('详细错误:', error);
        }
    };
    reader.readAsArrayBuffer(file);
}

// Word 文本转 Markdown（模拟）
function convertDocxTextToMd(text) {
    // 这里只是简单处理，实际应用中需要更复杂的解析
    const mdText = text;
    conversionResult = mdText;
    showStatus('转换成功！');
    downloadBtn.disabled = false;
    outputContent.textContent = mdText;
}

// 下载结果
function downloadResult() {
    if (!conversionResult) return;
    
    if (currentConverter === 'md-to-docx') {
        // 下载 Word 文档
        saveAs(conversionResult, 'converted.docx');
    } else {
        // 下载 Markdown 文件
        const blob = new Blob([conversionResult], { type: 'text/markdown' });
        saveAs(blob, 'converted.md');
    }
}

// 显示状态信息
function showStatus(message, type = 'info') {
    status.textContent = message;
    status.className = `status ${type}`;
    
    // 添加样式
    if (type === 'error') {
        status.style.color = '#e74c3c';
        status.style.backgroundColor = '#fadbd8';
    } else {
        status.style.color = '#34495e';
        status.style.backgroundColor = '#f8f9fa';
    }
}

// 显示转换进度
function showProgress(percent) {
    status.innerHTML = `转换中... ${percent}%`;
    status.className = 'status progress';
    status.style.color = '#3498db';
    status.style.backgroundColor = '#ebf5fb';
}

// 重置状态
function resetState() {
    conversionResult = null;
    downloadBtn.disabled = true;
    outputContent.textContent = '';
    status.textContent = '';
    status.className = 'status';
}

// 初始化应用
function initApp() {
    initEventListeners();
    setConverter('md-to-docx');
    setInput('text');
}

// 启动应用
initApp();
